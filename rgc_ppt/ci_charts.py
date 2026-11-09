"""Recorta la gráfica 'Desembolsos y Proyecciones' del informe de Cartera Inteligente (PDF de imagen).
Uso: python3 -I ci_charts.py "<informe CI.pdf>" <página 1-based> <código de 4 dígitos del préstamo, p.ej. 5823>
Guarda work/charts/<código>.png (la usa la lámina 'Desembolsos 2026'). La región está calibrada para el
informe extendido de CI en ancho 935 px: (228, 385, 612, 640); si el formato del informe cambia, ajústela."""
import sys, os, subprocess, tempfile
from PIL import Image
HERE = os.path.dirname(os.path.abspath(__file__)); WORK = os.environ.get('RGC_WORK', os.path.join(HERE, 'work'))
pdf, page, code = sys.argv[1], int(sys.argv[2]), sys.argv[3]
with tempfile.TemporaryDirectory() as t:
    subprocess.run(['pdftoppm', '-jpeg', '-r', '200', '-f', str(page), '-l', str(page), pdf, f'{t}/p'], check=True)
    im = Image.open(next(os.path.join(t, f) for f in os.listdir(t)))
k = im.width / 935
box = tuple(int(v * k) for v in (228, 385, 612, 640))
os.makedirs(f'{WORK}/charts', exist_ok=True)
im.crop(box).save(f'{WORK}/charts/{code}.png'); print('ok', f'{WORK}/charts/{code}.png', im.crop(box).size)
