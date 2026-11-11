/*
 * Builds icons and manifest.xml.  Usage:  node tools/build.js https://your-host.example/path
 * The host URL must be HTTPS and serve the contents of src/ at that path.
 */
const fs = require("fs");
const path = require("path");
const sharp = require("sharp");

const ROOT = path.join(__dirname, "..");
const SRC = path.join(ROOT, "src");
const host = (process.argv[2] || "https://YOUR-HOST.example/deli-elements").replace(/\/$/, "");
const GREEN = "#17362E", GOLD = "#C9A46A", CREAM = "#FAF8F4";

// Simple 32-unit glyphs for ribbon buttons.
const bar = (x, y, w, h, c = GREEN) => `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="1.5" fill="${c}"/>`;
const vline = (x) => `<rect x="${x - 0.75}" y="3" width="1.5" height="26" fill="${GOLD}"/>`;
const hline = (y) => `<rect x="3" y="${y - 0.75}" width="26" height="1.5" fill="${GOLD}"/>`;
const GLYPHS = {
  alignLeft: vline(5) + bar(7, 7, 18, 6) + bar(7, 18, 11, 6),
  alignCenter: vline(16) + bar(6, 7, 20, 6) + bar(10, 18, 12, 6),
  alignRight: vline(27) + bar(7, 7, 18, 6) + bar(14, 18, 11, 6),
  alignTop: hline(5) + bar(7, 7, 6, 18) + bar(18, 7, 6, 11),
  alignMiddle: hline(16) + bar(7, 6, 6, 20) + bar(18, 10, 6, 12),
  alignBottom: hline(27) + bar(7, 7, 6, 18) + bar(18, 14, 6, 11),
  distH: bar(3, 9, 6, 14) + bar(13, 9, 6, 14) + bar(23, 9, 6, 14),
  distV: bar(9, 3, 14, 6) + bar(9, 13, 14, 6) + bar(9, 23, 14, 6),
  sameSize: bar(4, 8, 10, 16) + bar(18, 8, 10, 16, GOLD),
  swapShapes: bar(3, 10, 9, 12) + bar(20, 10, 9, 12, GOLD) + `<path d="M13 13h6l-2-2M19 19h-6l2 2" stroke="${GREEN}" stroke-width="1.6" fill="none"/>`,
};

async function glyphIcon(name, size) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 32 32">${GLYPHS[name]}</svg>`;
  await sharp(Buffer.from(svg)).png().toFile(path.join(SRC, "assets", `${name}-${size}.png`));
}

async function logoIcon(size) {
  const r = Math.round(size * 0.18);
  const bg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}"><rect width="${size}" height="${size}" rx="${r}" fill="${GREEN}"/></svg>`;
  const lw = Math.round(size * 0.8), lh = Math.round(lw * 319 / 438);
  const logo = await sharp(path.join(SRC, "assets", "logo-cream.png")).resize(lw, lh).png().toBuffer();
  await sharp(Buffer.from(bg)).composite([{ input: logo, left: Math.round((size - lw) / 2), top: Math.round((size - lh) / 2) }]).png().toFile(path.join(SRC, "assets", `icon-${size}.png`));
}

const BUTTONS = [
  ["alignLeft", "Left", "Align selected shapes to the leftmost edge (one shape: to the slide)."],
  ["alignCenter", "Center", "Center selected shapes horizontally."],
  ["alignRight", "Right", "Align selected shapes to the rightmost edge."],
  ["alignTop", "Top", "Align selected shapes to the top edge."],
  ["alignMiddle", "Middle", "Center selected shapes vertically."],
  ["alignBottom", "Bottom", "Align selected shapes to the bottom edge."],
  ["distH", "Across", "Equal horizontal gaps between 3 or more shapes."],
  ["distV", "Down", "Equal vertical gaps between 3 or more shapes."],
  ["sameSize", "Same size", "Resize to match the first selected shape."],
  ["swapShapes", "Swap", "Swap the positions of two shapes."],
];

function manifest() {
  const img = (id, file) => `        <bt:Image id="${id}" DefaultValue="${host}/assets/${file}"/>`;
  const images = [16, 32, 64, 80].map((s) => img(`Icon.${s}`, `icon-${s}.png`))
    .concat(BUTTONS.flatMap(([n]) => [16, 32, 80].map((s) => img(`${n}.${s}`, `${n}-${s}.png`)))).join("\n");
  const button = ([n, label]) => `
                <Control xsi:type="Button" id="Deli.${n}">
                  <Label resid="${n}.Label"/>
                  <Supertip><Title resid="${n}.Label"/><Description resid="${n}.Tip"/></Supertip>
                  <Icon><bt:Image size="16" resid="${n}.16"/><bt:Image size="32" resid="${n}.32"/><bt:Image size="80" resid="${n}.80"/></Icon>
                  <Action xsi:type="ExecuteFunction"><FunctionName>${n}</FunctionName></Action>
                </Control>`;
  const group = (id, label, names) => `
              <Group id="Deli.${id}">
                <Label resid="${id}.Label"/>
                <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>${BUTTONS.filter((b) => names.includes(b[0])).map(button).join("")}
              </Group>`;
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f1d2b7e-3c4a-4d8e-9a5b-de11e1e2027a</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>The Deli Group</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Deli Elements"/>
  <Description DefaultValue="On-brand slide elements, layout tools, branded tables and a brand check for The Deli Group."/>
  <IconUrl DefaultValue="${host}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${host}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="${host}/taskpane.html"/>
  <AppDomains>
    <AppDomain>${new URL(host).origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.4"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${host}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Desc"/>
            <LearnMoreUrl resid="Taskpane.Url"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Deli.Tab">
              <Group id="Deli.Main">
                <Label resid="Main.Label"/>
                <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
                <Control xsi:type="Button" id="Deli.Open">
                  <Label resid="Open.Label"/>
                  <Supertip><Title resid="Open.Label"/><Description resid="Open.Tip"/></Supertip>
                  <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
                  <Action xsi:type="ShowTaskpane"><TaskpaneId>DeliElementsPane</TaskpaneId><SourceLocation resid="Taskpane.Url"/></Action>
                </Control>
              </Group>${group("Align", "Align", ["alignLeft", "alignCenter", "alignRight", "alignTop", "alignMiddle", "alignBottom"])}${group("Arrange", "Arrange", ["distH", "distV", "sameSize", "swapShapes"])}
              <Label resid="Tab.Label"/>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
${images}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${host}/taskpane.html"/>
        <bt:Url id="Commands.Url" DefaultValue="${host}/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Tab.Label" DefaultValue="Deli"/>
        <bt:String id="Main.Label" DefaultValue="Deli Elements"/>
        <bt:String id="Align.Label" DefaultValue="Align"/>
        <bt:String id="Arrange.Label" DefaultValue="Arrange"/>
        <bt:String id="Open.Label" DefaultValue="Elements"/>
        <bt:String id="GetStarted.Title" DefaultValue="Deli Elements is ready"/>
${BUTTONS.map(([n, l]) => `        <bt:String id="${n}.Label" DefaultValue="${l}"/>`).join("\n")}
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Open.Tip" DefaultValue="Open the Deli Elements panel: insert elements, layout tools, branded tables and a brand check."/>
        <bt:String id="GetStarted.Desc" DefaultValue="Open the Deli tab on the ribbon and click Elements."/>
${BUTTONS.map(([n, , t]) => `        <bt:String id="${n}.Tip" DefaultValue="${t}"/>`).join("\n")}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

(async () => {
  for (const s of [16, 32, 64, 80]) await logoIcon(s);
  for (const n of Object.keys(GLYPHS)) for (const s of [16, 32, 80]) await glyphIcon(n, s);
  fs.writeFileSync(path.join(ROOT, "manifest.xml"), manifest());
  console.log("manifest.xml written for", host);
})();
