# Generador de presentaciones RGC (sesiones previas)

Genera una presentación por operación de la agenda a partir del template `template_RGC.pptx`
(versión v2 con definiciones, sin refrendos) y los datos de la Torre de Control CCR.
Lo usa la skill `torre-control-actualizar-ppt`; ver su SKILL.md para el flujo completo y las reglas de datos.

Orden de ejecución (desde esta carpeta, con el repo stageado en RGC_BASE):
1. `work/agenda.json` – operaciones de la agenda, portada (META), correcciones manuales (MANUAL) y supuestos de análisis
   por operación (`supuestos`, p. ej. `{"CR-J0002": {"ext_meses": 18}}` para calcular el ritmo requerido con una extensión
   supuesta; la lámina lo rotula como supuesto).
2. `work/ci.json` – datos de Cartera Inteligente por préstamo (curados a mano desde el OCR del informe). Cada operación
   lleva `fecha` (creación del informe, ISO): es la fecha de corte que muestran la portada y las láminas 4 y 6.
3. `python3 ci_charts.py "<CI.pdf>" <pág> <código>` – gráficas de desembolsos → `work/charts/`.
4. `python3 -I extract.py` – Torre + agreement_report + PMR_Historico + Raw Data + Contracts and Amendments → `work/data.json`.
5. `python3 -I pmr_extract.py` – informes PMR (PDF), simuladores PMR y OPS1 → `work/pmr.json`, `work/ops.json`.
6. `python3 -I gen.py [CR-XXXX ...]` – presentaciones en `work/out/`.
7. `bash render.sh CR-XXXX` – render para QA en `work/r/<op>/`.

Variables: RGC_BASE (repo stageado), RGC_WORK (datos/salidas, por defecto ./work), RGC_CORTE (fecha de corte,
por defecto la de agenda.json), RGC_TEMPLATE, PPTX_SCRIPTS (scripts de la skill pptx; se detectan solos).

Reglas de cálculo incorporadas (sesión previa CR-J0002, 16/11/2026):
- Contratado (láminas 5 y 14): por proceso contratado se toma el mayor entre el plan de adquisiciones (monto del contrato
  o, si no hay, el estimado, x % BID; misma base que Cartera Inteligente) y la suma de sus contratos (Raw Data + Contracts
  and Amendments, sin rescindidos). Los procesos con varias filas cuentan el contrato una sola vez. "Otros costos" del plan
  se asocia al componente de administración de CI. Si el resultado es menor que lo desembolsado, la lámina queda con [__].
- Diagrama por objetivo: cada lámina muestra solo los indicadores con línea hacia algún producto de esa lámina.
- Productos logrados: sin riesgo de atraso en el detalle ni en el diagrama.
