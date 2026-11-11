# Deli Elements: PowerPoint add-in for The Deli Group

Adds a **Deli** tab to the PowerPoint ribbon:

- **Elements**: opens the side panel.
- **Align**: Left, Center, Right, Top, Middle, Bottom. With one shape selected, it aligns to the slide.
- **Arrange**: Distribute across or down, Same size (matches the first selected shape), Swap two shapes.

The side panel has four tabs:

| Tab | What it does |
|---|---|
| Insert | 26 one-click elements, grouped and on brand: Harvey balls, checks, traffic lights, status pills, number markers, stickers, takeaway and note boxes, source line, stat cards, chevron process, section tracker, 2x2 matrix, 12-month Gantt. It also inserts full slides from the Deli Elements library (agenda, process, Gantt, frameworks, driver tree, options, KPIs, icons, cover, closing). |
| Layout | Align, distribute, stack with 0.2" gaps, same width/height/size, swap, snap to grid, fit to content width |
| Tables & charts | Branded table (green header, Deli fonts); chart slides with native Deli-colored charts (column, line, stacked, doughnut, waterfall); number formatting for selected text (1,234,567 · Rp1.2bn · 12.5%) |
| Check | Brand check across the deck: off-brand fonts and colors, text under 9pt, shapes off the slide, leftover placeholders ([ ], TBD, TODO). One click fixes the fonts and colors. |

It works in PowerPoint for Microsoft 365 on Windows, Mac and the web. Branded tables and grouping elements need a current version of PowerPoint (PowerPointApi 1.8). On older versions everything else still works, and elements are inserted ungrouped.

---

## Step 1: Host the files (one time, about 10 minutes)

Office add-ins load from an HTTPS web address. The free route is GitHub Pages:

1. Create a GitHub repository, for example `deli-elements`. It can be private if your GitHub plan allows Pages on private repos; otherwise use a public repo. The files contain no data, only the toolkit.
2. Upload **everything inside the `src/` folder**, keeping the `assets/` subfolder, to the repository root.
3. In the repository, go to Settings → Pages → "Deploy from a branch" → `main` / root → Save.
4. After a minute your address is `https://<your-github-name>.github.io/deli-elements`. Open `.../taskpane.html` in a browser to check that it loads.

Any other HTTPS host works too (Netlify, Azure Static Web Apps, your company web server).

## Step 2: Point the manifest at your address

Open `manifest.xml` in a text editor and replace every `https://YOUR-HOST.example/deli-elements` with your address from Step 1, without a trailing slash. Save the file.

If you have Node.js, you can do the same with: `npm install && node tools/build.js https://<your-github-name>.github.io/deli-elements`.

## Step 3: Install in PowerPoint

**For the whole team (recommended).** A Microsoft 365 admin goes to admin.microsoft.com → Settings → Integrated apps → Upload custom apps. They choose "Office Add-in", upload `manifest.xml`, and assign users or groups. The Deli tab appears for everyone, usually within a few hours.

**On one Mac.** Create the folder `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` and copy `manifest.xml` into it. Restart PowerPoint. Then go to Insert → Add-ins → My Add-ins and choose Deli Elements.

**On one Windows PC.**
1. Put `manifest.xml` in a shared network folder.
2. In PowerPoint, go to File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs.
3. Add the folder's network path, tick "Show in Menu", and restart PowerPoint.
4. Go to Insert → My Add-ins → Shared Folder → Deli Elements.

**PowerPoint on the web.** Go to Insert → Add-ins → My Add-ins → Upload My Add-in, and choose `manifest.xml`.

---

## Notes and limits

- **Charts:** PowerPoint's add-in interface can't create charts, so chart buttons insert library slides that already contain native Deli charts. Right-click a chart → Edit Data to use your own numbers.
- **75% Harvey ball:** on current PowerPoint it is drawn the standard way (empty quarter top-left). Older versions can't rotate shapes, so the empty quarter shows top-right.
- **Library slides:** these are inserted after the slide you have selected and take on the deck's theme. Use the Deli template so colors stay right.
- **Updating:** change the files on your host. Everyone gets the new version the next time PowerPoint opens; nothing is reinstalled.

## For developers

```
src/core.js          pure logic: layout math, number formats, brand check, element definitions
src/taskpane.*       side panel (Office.js)
src/commands.*       ribbon buttons that run without the panel
tools/build.js       generates icons + manifest.xml for a given host
test/                node test/core.test.js · node test/taskpane.mock.test.js (needs jsdom)
```

Add a new element by adding an entry to `ELEMENTS` in `src/core.js`. The panel button appears automatically.
