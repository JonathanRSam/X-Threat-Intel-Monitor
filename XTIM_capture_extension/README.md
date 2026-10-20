# X Threat Intel Capture Extension

This Edge extension captures visible X/Twitter posts while the analyst manually scrolls approved X feeds/lists/searches.

Version 1.1.0 adds image capture support:

- stores post text
- stores external links
- stores attached X image URLs from post media
- stores image alt text when available
- exports all captured data as JSON for the Python processor

Daily usage:

1. Clear buffer at the start of the day.
2. Start capture.
3. Scroll X sources.
4. Export JSON.
5. Run Python processor.

If extension files are changed, reload the unpacked extension in `edge://extensions`.
