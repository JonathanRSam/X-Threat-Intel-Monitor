async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs && tabs.length ? tabs[0] : null;
}

async function sendToContent(message) {
  const tab = await getActiveTab();
  if (!tab || !tab.id) return { ok: false, error: "No active tab" };
  try {
    return await chrome.tabs.sendMessage(tab.id, message);
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function fmtLocal(iso) {
  if (!iso) return "-";
  try { return new Date(iso).toLocaleString(); } catch { return iso; }
}

async function refreshStatus() {
  const resp = await sendToContent({ type: "GET_STATUS" });
  document.getElementById("status").textContent = resp?.capturing ? "Capturing" : "Idle";
  document.getElementById("detectedCount").textContent = resp?.detectedCount ?? 0;
  document.getElementById("storedCount").textContent = resp?.storedCount ?? 0;
  document.getElementById("runNewCount").textContent = resp?.runNewCount ?? 0;
  document.getElementById("runDuplicateCount").textContent = resp?.runDuplicateCount ?? 0;
  document.getElementById("runImagePostCount").textContent = resp?.runImagePostCount ?? 0;
  document.getElementById("storedPostsWithImages").textContent = resp?.storedPostsWithImages ?? 0;
  document.getElementById("newestPublished").textContent = fmtLocal(resp?.newestPublishedUtc);
  document.getElementById("oldestPublished").textContent = fmtLocal(resp?.oldestPublishedUtc);
  document.getElementById("lastScan").textContent = fmtLocal(resp?.lastScanAt);
  document.getElementById("pageStatus").textContent = resp?.pageTitle || (resp?.pageUrl ? new URL(resp.pageUrl).hostname : "No X content script response");
}

document.getElementById("startBtn").addEventListener("click", async () => {
  await sendToContent({ type: "START_CAPTURE" });
  await refreshStatus();
});

document.getElementById("stopBtn").addEventListener("click", async () => {
  await sendToContent({ type: "STOP_CAPTURE" });
  await refreshStatus();
});

document.getElementById("scanBtn").addEventListener("click", async () => {
  await sendToContent({ type: "SCAN_NOW" });
  await refreshStatus();
});

document.getElementById("resetRunBtn").addEventListener("click", async () => {
  await sendToContent({ type: "RESET_RUN_COUNTERS" });
  await refreshStatus();
});

document.getElementById("purgeOldBtn").addEventListener("click", async () => {
  const ok = confirm("Remove stored posts whose published timestamp is older than 3 days? This does not affect already exported JSON files.");
  if (!ok) return;
  const resp = await sendToContent({ type: "PURGE_OLD_POSTS", days: 3 });
  alert(`Purged ${resp?.purgedCount ?? 0} old posts. Remaining stored posts: ${resp?.storedCount ?? 0}`);
  await refreshStatus();
});

document.getElementById("clearBtn").addEventListener("click", async () => {
  const ok = confirm("Clear the entire capture buffer? Do this at the start of a new day, after you are sure yesterday's JSON/Excel is saved.");
  if (!ok) return;
  await sendToContent({ type: "CLEAR_CAPTURE" });
  await refreshStatus();
});

document.getElementById("exportBtn").addEventListener("click", async () => {
  const resp = await sendToContent({ type: "EXPORT_CAPTURE" });
  const posts = resp?.posts || [];
  const data = JSON.stringify(posts, null, 2);
  const blob = new Blob([data], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  a.href = url;
  a.download = `x_captured_posts_${stamp}.json`;
  a.click();
  URL.revokeObjectURL(url);
  await refreshStatus();
});

refreshStatus();
setInterval(refreshStatus, 700);
