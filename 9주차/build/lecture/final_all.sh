set -e
cd /home/claude/lec
SK=$(ls -d /root/.claude/skills/synced/*/pptx | head -1)
for b in build_week09_lecture build_week09_practice build_week10_lecture build_week10_practice; do node $b.mjs | tail -1; done
rm -rf final && mkdir final
for f in output/*.pptx; do
  n=$(basename "$f" .pptx); d=final/$n; mkdir -p "$d"; cp "$f" "$d/deck.pptx"
  python3 $SK/scripts/office/validate.py "$f" | tail -1
  timeout 300 python3 lo_export.py "$d/deck.pptx" "$d/deck.pdf" $SK | tail -1
  (cd "$d" && pdftoppm -png -r 192 deck.pdf slide)
  python3 build_image_pdfs_practice.py "$d" output "$n"
done
