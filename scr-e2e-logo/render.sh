#!/bin/sh
# Render the SVG previews to PNG with the preinstalled Chromium (2x for crisp output).
CH=$(ls /opt/pw-browsers/chromium-*/chrome-linux/chrome | head -1)
for f in preview/slide1 preview/slide2 preview/slide3 SCR_E2E_Logo; do
  w=$(grep -o 'width="[0-9]*"' $f.svg | head -1 | tr -dc 0-9); h=$(grep -o 'height="[0-9]*"' $f.svg | head -1 | tr -dc 0-9)
  printf '<html><body style="margin:0;overflow:hidden;background:transparent">%s</body></html>' "$(cat $f.svg | sed 's/<svg /<svg style="display:block" /')" > /tmp/scr_preview.html
  $CH --headless --no-sandbox --disable-gpu --hide-scrollbars --force-device-scale-factor=2.5 --default-background-color=00000000 \
    --window-size=$w,$((h + 200)) --screenshot=$PWD/$f.png file:///tmp/scr_preview.html 2>/dev/null
  convert $PWD/$f.png -crop $((w * 5 / 2))x$((h * 5 / 2))+0+0 +repage $PWD/$f.png  # headless window is taller than its viewport
done
