#!/bin/bash
# Removes Deli Elements and the Deli templates from PowerPoint for Mac
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
TPL="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized"
osascript -e 'tell application "Microsoft PowerPoint" to quit' >/dev/null 2>&1
rm -f "$WEF/deli-elements.xml" "$TPL/Deli presentation.potx" "$TPL/The Deli Group.potx"
echo "Removing the shared installer files needs your Mac password:"
sudo rm -rf "/Library/Application Support/Deli Elements"
sudo pkgutil --forget com.thedeligroup.delielements >/dev/null 2>&1
echo "Deli Elements has been removed. You can close this window."
