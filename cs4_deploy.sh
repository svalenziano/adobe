#!/bin/bash
dest="/mnt/c/Program Files (x86)/Adobe/Adobe Illustrator CS4/Presets/en_US/Scripts"
src="$(dirname "$0")"
for f in "$src"/*.jsx; do
  cp -f "$f" "$dest"
done
