#!/bin/bash
set -eo pipefail
cd /home/claude/mac
VER=2.1.0; ID=com.thedeligroup.delielements
rm -rf work out; mkdir -p work/root "work/root/Library/Application Support/Deli Elements" work/flat/Resources work/flat/delielements.pkg out
D="work/root/Library/Application Support/Deli Elements"
cp /home/claude/exe/manifest.xml "$D/deli-elements.xml"
cp "/home/claude/exe/Deli presentation.potx" "/home/claude/exe/The Deli Group.potx" "$D/"
cp src/uninstall.command "$D/Uninstall Deli Elements.command"
chmod 755 "$D/Uninstall Deli Elements.command"; chmod 644 "$D"/*.xml "$D"/*.potx
find work/root -type d -exec chmod 755 {} +
P=work/flat/delielements.pkg
( cd work/root && find . | sort | cpio -o --format odc --owner 0:80 --quiet ) | gzip -9 > $P/Payload
mkdir -p work/scripts && cp src/scripts/* work/scripts/
( cd work/scripts && find . | cpio -o --format odc --owner 0:80 --quiet ) | gzip -9 > $P/Scripts
/tmp/bomutils/build/bin/mkbom -u 0 -g 80 work/root $P/Bom
NFILES=$(find work/root | wc -l); KB=$(du -sk work/root | cut -f1)
cat > $P/PackageInfo <<X
<?xml version="1.0" encoding="utf-8"?>
<pkg-info format-version="2" identifier="$ID" version="$VER" install-location="/" auth="root" overwrite-permissions="false">
  <payload numberOfFiles="$NFILES" installKBytes="$KB"/>
  <scripts><preinstall file="./preinstall"/><postinstall file="./postinstall"/></scripts>
  <bundle-version/>
</pkg-info>
X
cp src/resources/* work/flat/Resources/
cat > work/flat/Distribution <<X
<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="2">
  <title>Deli Elements for PowerPoint</title>
  <organization>com.thedeligroup</organization>
  <welcome file="welcome.html" mime-type="text/html"/>
  <conclusion file="conclusion.html" mime-type="text/html"/>
  <options customize="never" require-scripts="false" hostArchitectures="x86_64,arm64"/>
  <domains enable_localSystem="true"/>
  <choices-outline><line choice="default"><line choice="$ID"/></line></choices-outline>
  <choice id="default"/>
  <choice id="$ID" visible="false"><pkg-ref id="$ID"/></choice>
  <pkg-ref id="$ID" version="$VER" onConclusion="none" installKBytes="$KB">#delielements.pkg</pkg-ref>
</installer-gui-script>
X
( cd work/flat && xar --compression none -cf "../../out/Deli-Elements-Mac.pkg" * )
ls -la out
