# Generador de presentaciones RGC (sesiones previas)

Genera una presentación por operación de la agenda a partir del template `template_RGC.pptx`
(versión v2 con definiciones, sin refrendos) y los datos de la Torre de Control CCR.
Lo usa la skill `torre-control-actualizar-ppt`; ver su SKILL.md para el flujo completo y las reglas de datos.

Orden de ejecución (desde esta carpeta, con el repo stageado en RGC_BASE):
1. `work/agenda.json` – operaciones de la agenda, portada (META) y correcciones manuales (MANUAL).
2. `work/ci.json` – datos de Cartera Inteligente por préstamo (curados a mano desde el OCR del informe).
3. `python3 ci_charts.py "<CI.pdf>" <pág> <código>` – gráficas de desembolsos → `work/charts/`.
4. `python3 -I extract.py` – Torre + agreement_report + PMR_Historico + Raw Data + Contracts and Amendments → `work/data.json`.
5. `python3 -I pmr_extract.py` – informes PMR (PDF), simuladores PMR y OPS1 → `work/pmr.json`, `work/ops.json`.
6. `python3 -I gen.py [CR-XXXX ...]` – presentaciones en `work/out/`.
7. `bash render.sh CR-XXXX` – render para QA en `work/r/<op>/`.

Variables: RGC_BASE (repo stageado), RGC_WORK (datos/salidas, por defecto ./work), RGC_CORTE (fecha de corte,
por defecto la de agenda.json), RGC_TEMPLATE, PPTX_SCRIPTS (scripts de la skill pptx; se detectan solos).
