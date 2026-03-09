# Installation

Download the sideload manifest (this points to hosted HTTPS, so you won’t need localhost dev certificates):

- [manifest.sideload.xml](https://daedroth.github.io/name-indexer-word-add-in/dist/manifest.sideload.xml)

## Desktop Word (Windows/macOS)

### Windows

Option A (when available):

1. Download the XML linked above.
2. In Word: Insert → Get Add-ins (or Add-ins) → My Add-ins → Upload My Add-in.
3. Select the downloaded XML file.
4. Click Install.

Option B (official testing fallback when “Upload My Add-in” isn’t shown): install from a **Shared Folder** catalog.

1. Download the XML linked above and save it to a folder on your PC (for example `C:\OfficeAddins`).
2. Share that folder in Windows (Folder Properties → Sharing → Share) and note the network path (a UNC path like `\\YOUR-PC\OfficeAddins`).
3. In Word: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs.
4. Paste the UNC path into **Catalog Url**, choose **Add catalog**, check **Show in Menu**, then OK.
5. Restart Word.
6. In Word: Home → Add-ins → Advanced → **SHARED FOLDER** → select the add-in → Add.

Notes:

- This Shared Folder method is the Windows-only sideload approach documented by Microsoft for testing.
- The add-in itself still loads from the HTTPS URLs inside the manifest.

### macOS

On Mac, Microsoft’s supported sideload method is to copy the manifest into Word’s `wef` folder:

1. Download the XML linked above.
2. Quit Word.
3. Copy the XML to: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef` (create the `wef` folder if needed).
4. Reopen Word → open a document → Home → Add-ins → select your add-in.

## Word on the web

1. Open Word on the web and open a document.
2. Go to Insert → Add-ins → My Add-ins.
3. Choose Upload My Add-in (or Manage My Add-ins → Upload My Add-in).
4. Select the downloaded XML file.

If you don’t see an “Upload My Add-in” option, it’s usually disabled by your Microsoft 365 tenant admin.

If you’re using a personal Microsoft account and “Upload My Add-in” still isn’t available in desktop Word, use the Windows **Shared Folder** method above.
