How to run ProAV Shoko (macOS)
==============================

1) Drag "__APP_NAME__" into your Applications folder
   (or double-click it to run straight from this disk image).

2) Double-click the ProAV Shoko icon to start the app.

   The app is free and unsigned, so macOS Gatekeeper may block the
   first launch with "cannot be opened because the developer cannot
   be verified". That is expected. Use one of these to open it:

   Option A - Right-click (or Control-click) the app icon and choose
              "Open", then click "Open" again in the dialog.
              (Do this only once - after that it opens normally.)

   Option B - System Settings -> Privacy & Security -> Security ->
              "Open Anyway" next to ProAV Shoko.

   Option C - Remove the download quarantine from Terminal:

              xattr -dr com.apple.quarantine "/Applications/__APP_NAME__"

3) Command line version (text UI):

   "/Applications/__APP_NAME__/Contents/MacOS/ProAV Shoko" --cli

Troubleshooting
---------------
- "ProAV Shoko is damaged and can't be opened": use Option C above,
  then double-click the app again.
- Blank/empty window after starting: click "Start Analysis".
- Works on macOS 10.15 and later, Intel and Apple Silicon.
