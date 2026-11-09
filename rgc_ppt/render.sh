#!/bin/bash
# Uso: bash render.sh CR-L1032  -> work/r/<op>/a-NN.jpg (solo láminas visibles) y grillas gNN.jpg
HERE=$(cd "$(dirname "$0")" && pwd); W=${RGC_WORK:-$HERE/work}
SOFF=$(ls /mnt/skills/public/pptx/scripts/office/soffice.py 2>/dev/null || find /root/.claude/skills -path '*pptx/scripts/office/soffice.py' | head -1)
op=$1; mkdir -p $W/r/$op; cd $W/r/$op && rm -f *.jpg *.pdf
python3 $SOFF --headless --convert-to pdf $W/out/$op.pptx >/dev/null 2>&1
pdftoppm -jpeg -r 60 $op.pdf a
python3 -c "
from PIL import Image;import glob
fs=sorted(glob.glob('a-*.jpg'));ims=[Image.open(f) for f in fs];w,h=ims[0].size
for k in range(0,len(ims),6):
  g=Image.new('RGB',(w*2,h*3),'white')
  for i,im in enumerate(ims[k:k+6]): g.paste(im,((i%2)*w,(i//2)*h))
  g.save(f'g{k//6:02d}.jpg')
print(len(ims), 'láminas visibles')"
