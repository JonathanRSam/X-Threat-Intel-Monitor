let capturing = false;
let capturedPosts = [];
let capturedKeys = new Set();
let detectedCount = 0;
let lastScanAt = null;
let scanTimer = null;

let runNewCount = 0;
let runDuplicateCount = 0;
let runDuplicateKeys = new Set();
let runImagePostCount = 0;
let baselineKeysAtRunStart = new Set();

let loadedFromStorage = false;
const STORAGE_KEY = "x_threat_intel_captured_posts_v1";
const SCAN_INTERVAL_MS = 500;

function storageGet(key) {
  return new Promise(resolve => chrome.storage.local.get([key], result => resolve(result[key])));
}
function storageSet(key, value) {
  return new Promise(resolve => chrome.storage.local.set({ [key]: value }, resolve));
}
async function loadStoredPosts() {
  if (loadedFromStorage) return;
  const stored = await storageGet(STORAGE_KEY);
  if (Array.isArray(stored)) {
    capturedPosts = stored;
    capturedKeys = new Set(stored.map(p => p.fingerprint || p.post_id || p.post_url).filter(Boolean));
  }
  loadedFromStorage = true;
}
async function saveStoredPosts() {
  await storageSet(STORAGE_KEY, capturedPosts);
}
function resetRunCounters() {
  runNewCount = 0;
  runDuplicateCount = 0;
  runDuplicateKeys = new Set();
  runImagePostCount = 0;
  baselineKeysAtRunStart = new Set(capturedKeys);
}
function isElementVisible(el) {
  if (!el) return false;
  const rect = el.getBoundingClientRect();
  const vh = window.innerHeight || document.documentElement.clientHeight;
  const vw = window.innerWidth || document.documentElement.clientWidth;
  return rect.bottom > 0 && rect.right > 0 && rect.top < vh && rect.left < vw && rect.width > 100 && rect.height > 40;
}
function normalizeSpace(text) {
  return (text || "").replace(/\s+/g, " ").trim();
}
function absoluteUrl(href) {
  if (!href) return "";
  try { return new URL(href, window.location.origin).toString(); } catch { return href; }
}
function extractStatusLinks(article) {
  return Array.from(article.querySelectorAll('a[href*="/status/"]'))
    .map(a => absoluteUrl(a.getAttribute("href") || a.href))
    .filter(Boolean)
    .filter(u => /\/status\/\d+/.test(u));
}
function pickCanonicalPostUrl(article) {
  const links = extractStatusLinks(article);
  if (!links.length) return "";
  const timeEl = article.querySelector("time");
  if (timeEl) {
    const parentA = timeEl.closest("a[href*='/status/']");
    if (parentA) return absoluteUrl(parentA.getAttribute("href") || parentA.href);
  }
  return links[0];
}
function extractPostId(postUrl) {
  const m = (postUrl || "").match(/\/status\/(\d+)/);
  return m ? m[1] : "";
}
function extractHandleFromUrl(postUrl) {
  try {
    const u = new URL(postUrl);
    const parts = u.pathname.split("/").filter(Boolean);
    if (parts.length >= 3 && parts[1] === "status") return "@" + parts[0];
  } catch {}
  return "";
}
function extractHandleFromText(text) {
  const m = (text || "").match(/@([A-Za-z0-9_]{1,20})\b/);
  return m ? "@" + m[1] : "";
}
function escapeRegExp(s) {
  return (s || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function extractDisplayName(article, handle) {
  const text = normalizeSpace(article.innerText || article.textContent || "");
  if (!text) return "";
  if (handle) {
    const escaped = escapeRegExp(handle);
    const re = new RegExp("(?:^|reposted\\s+)(.{1,80}?)\\s+" + escaped + "\\s+·", "i");
    const m = text.match(re);
    if (m && m[1]) return normalizeSpace(m[1].replace(/^.*reposted\s+/i, ""));
  }
  const lines = (article.innerText || "").split("\n").map(normalizeSpace).filter(Boolean);
  const handleNoAt = handle ? handle.slice(1) : "";
  for (const line of lines.slice(0, 8)) {
    if (!line || line.startsWith("@")) continue;
    if (handleNoAt && line.includes(handleNoAt)) continue;
    if (/^(reposted|replying to|show more)$/i.test(line)) continue;
    if (line.length <= 80) return line.replace(/reposted$/i, "").trim();
  }
  return "";
}
function extractTimestamp(article) {
  const timeEl = article.querySelector("time");
  return timeEl ? (timeEl.getAttribute("datetime") || "") : "";
}
function extractAllExternalLinks(article) {
  const links = [];
  for (const a of Array.from(article.querySelectorAll("a[href]"))) {
    const href = absoluteUrl(a.getAttribute("href") || a.href);
    if (!href) continue;
    if (href.includes("/status/")) continue;
    if (href.includes("x.com/") || href.includes("twitter.com/")) continue;
    links.push(href);
  }
  return Array.from(new Set(links));
}
function cleanPostText(rawText, displayName, handle) {
  let text = normalizeSpace(rawText);
  text = text.replace(/^.*?\s+reposted\s+/i, "");
  if (displayName && handle) {
    const escapedName = escapeRegExp(displayName);
    const escapedHandle = escapeRegExp(handle);
    text = text.replace(new RegExp("^" + escapedName + "\\s+" + escapedHandle + "\\s+·\\s+[^ ]+\\s*", "i"), "");
  }
  text = text.replace(/\s+(\d+[.,]?\d*[KMB]?\s*){2,5}$/i, "").trim();
  return text;
}
function getPageTitle() {
  const h1 = document.querySelector('h1[role="heading"], h1');
  const h1Text = h1 ? normalizeSpace(h1.innerText || h1.textContent || "") : "";
  if (h1Text) return h1Text;
  return normalizeSpace(document.title || "");
}
function extractPostImageUrls(article) {
  const candidates = [];
  for (const img of Array.from(article.querySelectorAll("img[src]"))) {
    const src = absoluteUrl(img.getAttribute("src") || img.src || "");
    if (!src) continue;
    // X attached media is normally hosted on pbs.twimg.com/media. Profile images and emoji are excluded.
    if (/pbs\.twimg\.com\/media\//i.test(src) || /pbs\.twimg\.com\/card_img\//i.test(src)) {
      candidates.push(src.replace(/&name=[^&]+/i, "&name=large"));
    }
  }
  return Array.from(new Set(candidates));
}
function extractImageAltTexts(article) {
  return Array.from(article.querySelectorAll("img[alt]"))
    .map(img => normalizeSpace(img.getAttribute("alt") || ""))
    .filter(Boolean)
    .filter(t => !/^(image|photo|profile picture)$/i.test(t))
    .slice(0, 12);
}

function extractPost(article) {
  const rawText = normalizeSpace(article.innerText || article.textContent || "");
  const postUrl = pickCanonicalPostUrl(article);
  const postId = extractPostId(postUrl);
  const handle = extractHandleFromUrl(postUrl) || extractHandleFromText(rawText);
  const displayName = extractDisplayName(article, handle);
  const publishedUtc = extractTimestamp(article);
  const externalLinks = extractAllExternalLinks(article);
  const postImageUrls = extractPostImageUrls(article);
  const imageAltTexts = extractImageAltTexts(article);
  const postText = cleanPostText(rawText, displayName, handle);
  const fingerprint = postId ? `x_post::${postId}` : `${handle}|${publishedUtc}|${postText.slice(0, 180)}`;
  return {
    platform: "X",
    display_name: displayName,
    handle,
    post_text: postText,
    raw_visible_text: rawText,
    post_url: postUrl,
    post_id: postId,
    published_utc: publishedUtc,
    external_links: externalLinks,
    post_image_urls: postImageUrls,
    post_image_count: postImageUrls.length,
    image_alt_texts: imageAltTexts,
    captured_from_url: window.location.href,
    captured_from_title: getPageTitle(),
    captured_at_utc: new Date().toISOString(),
    fingerprint
  };
}
async function scanVisiblePosts() {
  await loadStoredPosts();
  const articles = Array.from(document.querySelectorAll("article"));
  const visibleArticles = articles.filter(isElementVisible);
  detectedCount = visibleArticles.length;
  lastScanAt = new Date().toISOString();

  if (!capturing) return;

  let added = 0;
  for (const article of visibleArticles) {
    const post = extractPost(article);
    if (!post.fingerprint || !post.post_text || post.post_text.length < 10) continue;
    if (capturedKeys.has(post.fingerprint)) {
      if (baselineKeysAtRunStart.has(post.fingerprint) && !runDuplicateKeys.has(post.fingerprint)) {
        runDuplicateKeys.add(post.fingerprint);
        runDuplicateCount += 1;
      }
      continue;
    }
    capturedKeys.add(post.fingerprint);
    capturedPosts.push(post);
    runNewCount += 1;
    if ((post.post_image_urls || []).length > 0) runImagePostCount += 1;
    added += 1;
  }
  if (added > 0) await saveStoredPosts();
}
function getPublishedRange() {
  const times = capturedPosts
    .map(p => p.published_utc)
    .filter(Boolean)
    .map(t => Date.parse(t))
    .filter(n => !Number.isNaN(n));
  if (!times.length) return { newestPublishedUtc: "", oldestPublishedUtc: "" };
  return {
    newestPublishedUtc: new Date(Math.max(...times)).toISOString(),
    oldestPublishedUtc: new Date(Math.min(...times)).toISOString()
  };
}
async function purgeOldPosts(days) {
  await loadStoredPosts();
  const cutoffMs = Date.now() - (Number(days || 3) * 24 * 60 * 60 * 1000);
  const before = capturedPosts.length;
  capturedPosts = capturedPosts.filter(p => {
    const t = Date.parse(p.published_utc || "");
    if (Number.isNaN(t)) return true;
    return t >= cutoffMs;
  });
  capturedKeys = new Set(capturedPosts.map(p => p.fingerprint || p.post_id || p.post_url).filter(Boolean));
  resetRunCounters();
  await saveStoredPosts();
  return before - capturedPosts.length;
}
function startTimer() {
  if (scanTimer) clearInterval(scanTimer);
  scanTimer = setInterval(scanVisiblePosts, SCAN_INTERVAL_MS);
  scanVisiblePosts();
}
async function getStatus() {
  await loadStoredPosts();
  return {
    ok: true,
    capturing,
    detectedCount,
    storedCount: capturedPosts.length,
    runNewCount,
    runDuplicateCount,
    runImagePostCount,
    storedPostsWithImages: capturedPosts.filter(p => Array.isArray(p.post_image_urls) && p.post_image_urls.length > 0).length,
    lastScanAt,
    pageUrl: window.location.href,
    pageTitle: getPageTitle(),
    ...getPublishedRange()
  };
}
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || !message.type) return;
  (async () => {
    if (message.type === "START_CAPTURE") {
      await loadStoredPosts();
      resetRunCounters();
      capturing = true;
      startTimer();
      sendResponse(await getStatus());
      return;
    }
    if (message.type === "STOP_CAPTURE") {
      capturing = false;
      sendResponse(await getStatus());
      return;
    }
    if (message.type === "SCAN_NOW") {
      await scanVisiblePosts();
      sendResponse(await getStatus());
      return;
    }
    if (message.type === "RESET_RUN_COUNTERS") {
      await loadStoredPosts();
      resetRunCounters();
      sendResponse(await getStatus());
      return;
    }
    if (message.type === "PURGE_OLD_POSTS") {
      const purgedCount = await purgeOldPosts(message.days || 3);
      sendResponse({ ...(await getStatus()), purgedCount });
      return;
    }
    if (message.type === "CLEAR_CAPTURE") {
      capturedPosts = [];
      capturedKeys = new Set();
      detectedCount = 0;
      lastScanAt = null;
      resetRunCounters();
      await storageSet(STORAGE_KEY, []);
      sendResponse(await getStatus());
      return;
    }
    if (message.type === "EXPORT_CAPTURE") {
      await loadStoredPosts();
      sendResponse({ ok: true, posts: capturedPosts, ...(await getStatus()) });
      return;
    }
    if (message.type === "GET_STATUS") {
      sendResponse(await getStatus());
      return;
    }
  })();
  return true;
});
loadStoredPosts().then(startTimer);
console.log("[X Threat Intel Capture] Content script loaded on", window.location.href);
