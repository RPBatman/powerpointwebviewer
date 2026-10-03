# Live Website Viewer for PowerPoint

A PowerPoint content add-in that shows a live website inside a slide. Each inserted
viewer remembers its own URL inside the `.pptx`.

This is a personal fork of [peopleworks/powerpointwebviewer](https://github.com/peopleworks/powerpointwebviewer)
(MIT, by Pedro Hernandez), hosted on GitHub Pages:

| What | URL |
|------|-----|
| Viewer page | https://rpbatman.github.io/powerpointwebviewer/viewer.html |
| Manifest | https://rpbatman.github.io/powerpointwebviewer/manifest.xml |

## What changed from upstream

- **Own fixed add-in Id** `f4013c8d-7276-4ecc-af95-cd22c0783867`. Do not change it: slides that
  contain the add-in find it again by this Id.
- **Hosted on my GitHub Pages**, deployed by `.github/workflows/pages.yml` on every push to `main`.
- **Cache-busting**: `SourceLocation` ends in `?v=<Version>`, so a new release is a new URL and
  PowerPoint's WebView cache can't serve an old `viewer.html`.
- **Safer saving**: the `saveAsync` result is checked. You see a small *"Saved — remember to save
  the file (Ctrl+S)"* note, or a red error that stays on screen if saving failed.
- **More reliable restore**: the saved URL is read in `Office.onReady`. A URL entered before Office
  finished starting is still saved. Stray `about:blank` load events are ignored.
- **Clear startup errors**: you get a message with a Reload button, not a blank frame, if
  `office.js` can't be downloaded or PowerPoint never finishes starting the add-in (15 s).
- PeopleWorks branding footer, credit line and "P" icon monogram removed.

## Install on PowerPoint desktop (Windows) via a Trusted Add-in Catalog

PowerPoint desktop has no "Upload My Add-in" button. A shared-folder catalog is the supported
way to sideload, and it **survives restarts**.

### 1. Create the catalog folder and share it

1. Create a local folder **outside OneDrive**, e.g. `C:\AddinCatalog`.
2. Copy `manifest.xml` (from this repo, or download it from the Manifest URL above) into it.
3. Right-click the folder > **Properties** > **Sharing** > **Share...**
4. Add your own user with **Read** permission > **Share** > **Done**.
5. Note the **network path** shown, e.g. `\\YOUR-PC\AddinCatalog`.
   (Run `hostname` in a terminal if you're unsure of the PC name.)

   Or, from an **admin** PowerShell:

   ```powershell
   New-Item -ItemType Directory -Force C:\AddinCatalog
   New-SmbShare -Name AddinCatalog -Path C:\AddinCatalog -ReadAccess "$env:USERDOMAIN\$env:USERNAME"
   ```

### 2. Trust the catalog in PowerPoint

1. **File** > **Options** > **Trust Center** > **Trust Center Settings...** > **Trusted Add-in Catalogs**.
2. In **Catalog Url**, enter the UNC path `\\YOUR-PC\AddinCatalog`. A drive letter like `C:\...`
   won't work; it must start with `\\`.
3. Click **Add catalog** and tick **Show in Menu**.
4. **OK** > **OK**, then **close and reopen PowerPoint**.

> Add the catalog **once** and leave it alone. Each catalog entry gets its own internal Id, and
> slides record which catalog their add-in came from. Removing and re-adding the catalog can
> orphan existing slides.

### 3. Insert the viewer

1. **Home** > **Add-ins** > **More Add-ins** (or **Advanced**) > **SHARED FOLDER** tab.
2. Choose **Live Website Viewer** > **Add**.
3. Enter a URL and press **Enter**. Wait for the green *Saved* note.
4. **Save the presentation (Ctrl+S).** The URL is only written to disk when the file is saved.

### Migrating slides made with the upstream add-in

Slides inserted with the upstream add-in point at its Id (`e2b7c1a0-1234-...`) and its server.
Delete those viewers, insert this one, and enter the URLs again.

## Releasing an update

1. Edit `viewer.html`.
2. Bump the version in **three** places to the same value, e.g. `1.4.1.0`:
   - `manifest.xml` > `<Version>`
   - `manifest.xml` > `SourceLocation` `?v=` query
   - `viewer.html` > `VIEWER_VERSION`

   The Pages workflow fails if these don't match.
3. Commit and push to `main`. GitHub Pages redeploys automatically.
4. Copy the new `manifest.xml` into `C:\AddinCatalog`, then restart PowerPoint.
5. Hover the DESKTOP badge in a viewer: its tooltip shows the version actually running.

If PowerPoint still shows an old version, close PowerPoint and clear the Office add-in cache,
then reopen:

```powershell
Remove-Item -Recurse -Force "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*"
```

## Troubleshooting: viewer is blank after closing and reopening the .pptx

The add-in itself is just a web page. The `.pptx` stores two things: a reference to the add-in
(its Id plus the store or catalog it came from) and the saved URL setting. If either is missing on
reopen, you get an empty frame. Inspect a deck like this:

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead("C:\path\to\deck.pptx")
$zip.Entries | Where-Object FullName -like 'ppt/webextensions/webextension*.xml' |
  ForEach-Object { (New-Object IO.StreamReader($_.Open())).ReadToEnd() }
$zip.Dispose()
```

| What you see in the XML | Meaning | Fix |
|---|---|---|
| `<we:reference id="e2b7c1a0-1234-…">` | Slide uses the upstream add-in, not this one | Re-insert this add-in |
| `storeType`/`store` doesn't match your shared-folder catalog | Add-in was loaded some temporary way (Online upload, dev sideload, a catalog since removed) and PowerPoint can't find the manifest again | Install via the catalog above and re-insert |
| No `<we:property name="webViewerUrl" …>` | URL was never saved into the file (closed without saving) | Enter the URL, wait for *Saved*, press Ctrl+S |
| Everything present, still blank | Website loaded but refuses to be framed, or a stale cached `viewer.html` | Try the **Window** button; clear the `Wef` cache (above) |

The viewer now shows an explicit error if `office.js` fails or Office never initialises, so a
completely white pane points to the add-in not being loaded at all (rows 1–2).

## Local testing

```bash
python -m http.server 8765
# open http://127.0.0.1:8765/viewer.html  (badge shows STANDALONE; URLs aren't saved outside PowerPoint)
```

## Files

| File | Description |
|------|-------------|
| `viewer.html` | The add-in page shown inside the slide |
| `manifest.xml` | Office add-in manifest (copy this into your catalog folder) |
| `icon-32.png`, `icon-80.png`, `icon.svg` | Add-in icons |
| `.github/workflows/pages.yml` | GitHub Pages deployment |
| `powerpointwebviewer.html`, `web.config`, `*.docx` | Upstream docs and IIS config, not deployed |

## Limitations

- Sites that send `X-Frame-Options` / CSP `frame-ancestors` (Google, Facebook, X…) can't be framed.
  Use the **Window** button for those.
- YouTube and Vimeo links are converted to their embeddable player automatically.
- Needs an internet connection: both `office.js` and the viewer are loaded online.

## License

MIT. Original work © Pedro Hernandez / PeopleWorks Services.
