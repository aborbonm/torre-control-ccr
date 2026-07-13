# Torre de Control — CCR

Dashboard interactivo para el monitoreo de operaciones del Grupo BID en Costa Rica.
Permite visualizar el estado de cláusulas contractuales, productos críticos, adquisiciones
críticas, equipo, documentos y calificaciones PMR de forma centralizada y actualizada.

## Acceso

🔗 [aborbonm.github.io/torre-control-ccr](https://aborbonm.github.io/torre-control-ccr/)

## Cómo funciona

El dashboard lee sus datos al cargarse:

- **`torre_de_control_CCR.xlsx`** — contiene información de productos críticos, adquisiciones
  críticas, equipo, documentos, resultados, PMR, comentarios de cláusulas y change log.
- **`Status Date Clause Operation.csv`** — contiene el estado de cláusulas contractuales,
  descargado desde el tablero de PowerBI del BID.
- **`Raw Data - <préstamo>.xlsx`** — uno por operación (opcional), con el detalle de los
  procesos de adquisición de esa operación. Ver [Adquisiciones críticas por operación](#adquisiciones-críticas-por-operación).
- **`PMR_Historico.xlsx`** — opcional, con el histórico de ciclos PMR por operación.
  Ver [Histórico de PMR por operación](#histórico-de-pmr-por-operación).

Los archivos se leen directamente desde este repositorio usando la API de GitHub.
No se requiere instalación ni configuración — basta con abrir el link.

Los archivos terminados en `- Copy` son respaldos de la última versión estable.

Los datos se guardan en caché local (IndexedDB) para que la página cargue
instantáneamente. Al abrir la app, se muestra la versión en caché y se actualiza
en segundo plano. El botón **Actualizar Datos** fuerza una recarga inmediata.

## Cómo actualizar los datos

### Excel (`torre_de_control_CCR.xlsx`)

La información se actualiza directamente en el archivo Excel. Reglas importantes:
- No cambiar el nombre de las hojas
- No cambiar el orden ni el nombre de las columnas
- La hoja `comentarios_clausulas` permite registrar comentarios sobre cláusulas
  vencidas (ver sección abajo)

### CSV (`Status Date Clause Operation.csv`)

El archivo se descarga desde el tablero de cláusulas contractuales del BID:
- URL: https://app.powerbi.com/groups/d9b55ff5-07fe-4a88-8995-d94ed53cd7cf/rdlreports/429c4230-9c53-46ba-8c96-94ec89aa1cd2
- Filtros antes de descargar: **Country = Costa Rica**, **Clause Status = ALL**, **Operation Stage = Disbursing**

### Raw Data (`Raw Data - <préstamo>.xlsx`)

El archivo se descarga desde el [Portal de Cliente del BID](https://clientportal.iadb.org/):
- Entra a la operación correspondiente y haz clic en el enlace de descarga de esa data.
- El archivo descargado **usualmente** ya trae el nombre correcto, pero hay que
  validarlo: si el navegador ya tenía un archivo con ese nombre, puede agregar un
  `(2)` o `(3)` al final — hay que renombrarlo para que calce con la convención
  (ver [Adquisiciones críticas por operación](#adquisiciones-críticas-por-operación)).
- Sube el archivo **sin modificar nada del raw data**.

### Pasos para subir archivos

1. En este repositorio, haz clic en el archivo que deseas reemplazar
2. Clic en el ícono de lápiz → **"Upload a new version"** → sube el archivo
3. Clic en **"Commit changes"**

El dashboard reflejará los datos nuevos en la próxima visita o al presionar **Actualizar Datos**.

## Adquisiciones críticas por operación

El tab **Productos Críticos** y el popup de producto en **Resultados** pueden mostrar,
para cada producto crítico, el nombre de su(s) adquisición(es) crítica(s) asociada(s)
como un chip clicable — con detalle de ID de proceso, monto estimado, estado del
proceso, método de adquisición y riesgo al hacer clic.

Esto depende de dos piezas de datos:

1. **Hoja `productos_adquisiciones`** (en `torre_de_control_CCR.xlsx`): relaciona cada
   producto crítico con el o los procesos de adquisición asociados.
   Columnas: `Operacion | Producto_Critico | ID_Proceso | Adquisicion_Critica | Riesgos_Comentarios`.
   - `Producto_Critico` debe coincidir exactamente con el texto de la columna
     `Producto Crítico` de la hoja `productos criticos`.
   - `Adquisicion_Critica` = `Sí` marca esa fila como la adquisición crítica a mostrar
     en el tab; `No` la deja registrada pero oculta (por ejemplo, procesos de apoyo o
     de menor monto que no ameritan seguimiento destacado).
   - `Riesgos_Comentarios` solo debería llenarse en filas `Sí` — es lo que se muestra
     en el popup de detalle de esa adquisición.

2. **Archivos `Raw Data - <préstamo>.xlsx`** (uno por operación, en la raíz del repo):
   exportados del portal de cliente del BID (reporte "Plan de Adquisiciones"). El
   dashboard usa la hoja `Plan de Adquisiciones` de cada archivo para obtener el
   nombre, monto estimado, estado del proceso y método de adquisición de cada
   `ID_Proceso`.
   - **Convención de nombre**: `Raw Data - <aprobación con "/" reemplazado por "_">.xlsx`,
     donde `<aprobación>` es el número de préstamo tal como aparece en la columna
     `aprobacion` de la hoja `datos_ops` (ej. `6061/OC-CR` → `Raw Data - 6061_OC-CR.xlsx`).

**Operaciones con raw data hoy**: CR-L1032 (`3071_OC-CR`), CR-L1139 (`4864_OC-CR`),
CR-L1137 (`4871_OC-CR`), CR-J0002 (`5777_GR-CR`), CR-L1151 (`5823_OC-CR`), CR-L1157
(`6061_OC-CR`). Para el resto de las operaciones (sin raw data todavía), el tab sigue
mostrando el texto libre de las columnas `Adquisición Crítica` / `Riesgos/Comentarios`
de `productos criticos`, como fallback automático.

### Cómo agregar el raw data de una operación nueva

1. Sube el archivo `Raw Data - <préstamo>.xlsx` a la raíz del repo, con el nombre
   exacto según la convención de arriba (mismos pasos de "Pasos para subir archivos").
2. En `productos_adquisiciones`, agrega una fila por cada (producto crítico, proceso)
   que quieras asociar, indicando `Sí`/`No` en `Adquisicion_Critica` y, si aplica,
   el riesgo en `Riesgos_Comentarios` (solo en filas `Sí`).
3. Sube el Excel actualizado. El dashboard detecta automáticamente que la operación
   ya tiene filas en `productos_adquisiciones` y empieza a buscar su raw data.

## Histórico de PMR por operación

En el tab **PMR**, las tarjetas de las operaciones que tienen histórico cargado
muestran una insignia **"🕑 Histórico"** — al hacer clic en la tarjeta se abre un
popup con la tabla de todos los ciclos PMR registrados para esa operación.

Esto depende del archivo:

**`PMR_Historico.xlsx`** (raíz del repo, opcional): un archivo con **una hoja por
operación**, nombrada exactamente igual al código de operación (`datos_ops.codigo`,
ej. `CR-L1032`). Cada hoja tiene una fila por ciclo PMR, con columnas:
`Ciclo | Estado del PMI | CPI | SPI | CPI(a) | SPI(a) | Disb | T3 | Indicador Sintético | Clasificación Auto calculada | Clasificación Validada`.

No hay una cantidad fija de ciclos por operación — cada hoja tiene tantas filas
como ciclos tenga esa operación (varía según su antigüedad), y el dashboard ordena
los ciclos cronológicamente sin importar el orden de las filas en el archivo.

Hoy el archivo trae **una sola fila semilla por operación** (el ciclo actual,
"Ciclo I-2026"), tomada de lo que ya existía en la hoja `PMR` — a la espera de
que se consiga la fuente real con el histórico completo (2015 en adelante, según
el reporte de origen) para completar cada hoja con sus ciclos anteriores.

### Cómo agregar ciclos históricos

1. Abre `PMR_Historico.xlsx` y ve a la hoja de la operación correspondiente
   (créala si la operación todavía no tiene hoja, usando su código como nombre).
2. Agrega una fila por cada ciclo adicional, con las columnas de arriba.
3. Sube el archivo actualizado (ver "Pasos para subir archivos").

Si una operación no tiene hoja en `PMR_Historico.xlsx` (o el archivo no está
disponible), su tarjeta en el tab PMR se ve igual que hoy, sin insignia ni clic.

## Cómo registrar comentarios en cláusulas vencidas

1. Abre `torre_de_control_CCR.xlsx` y ve a la hoja **`comentarios_clausulas`**
2. Agrega una fila con los siguientes campos:

| Columna | Descripción |
|---------|-------------|
| `Clause_ID` | ID exacto de la cláusula (tomarlo del CSV) |
| `Fecha_Comentario` | Fecha en que se registra el comentario |
| `Autor` | Nombre de quien registra |
| `Comentario` | Texto del comentario |
| `Estado_Gestion` | `En gestión` / `Resuelto` / `Escalado` |
| `Fecha_Compromiso` | Fecha estimada de resolución (opcional) |

3. Sube el archivo actualizado al repositorio (ver pasos anteriores)

El dashboard mostrará el comentario más reciente por cláusula en el tab **Cláusulas**.

## Tabs del dashboard

| Tab | Fuente | Descripción |
|-----|--------|-------------|
| Resumen | xlsx + csv | KPIs generales: cláusulas vencidas, próximas, productos retrasados, operaciones activas. Tabla de estado por operación y alertas de vencimiento próximo (≤30 días) |
| Cláusulas | csv + xlsx | Dos secciones: **Vencidas** (con comentarios) y **Próximos 180 días** con barra de urgencia |
| Resultados | xlsx + raw data | Indicadores de la Matriz de Resultados agrupados por operación y objetivo, con probabilidad Alta/Media/Baja. Los chips de producto crítico muestran su(s) adquisición(es) crítica(s) al hacer clic (ver [Adquisiciones críticas por operación](#adquisiciones-críticas-por-operación)) |
| Productos Críticos | xlsx + raw data | Estado de productos con filtro por operación (Logrado / En tiempo / Retrasado). Cada producto muestra el nombre de su(s) adquisición(es) crítica(s) como chip clicable, con detalle de ID de proceso, monto, estado y método |
| Equipo | xlsx | Composición del equipo por operación con alerta de roles faltantes |
| Documentos | xlsx | Matriz de documentos clave con enlaces a SharePoint por tipo y operación |
| PMR | xlsx + PMR_Historico | Calificaciones e indicadores técnicos del ciclo PMR I-2026 por programa. Las tarjetas con histórico cargado son clicables y muestran la tabla de ciclos anteriores (ver [Histórico de PMR por operación](#histórico-de-pmr-por-operación)) |
| Change Log | xlsx | Historial de cambios registrados en el Excel (más reciente primero) |

## Estructura del repositorio

| Archivo | Descripción |
|---------|-------------|
| `index.html` | Dashboard completo (aplicación de una sola página, sin servidor) |
| `torre_de_control_CCR.xlsx` | Datos: productos críticos, adquisiciones críticas, equipo, PMR, documentos, resultados, comentarios de cláusulas y changelog |
| `Status Date Clause Operation.csv` | Estado de cláusulas contractuales (fuente: PowerBI BID) |
| `Raw Data - <préstamo>.xlsx` | Detalle del Plan de Adquisiciones por operación (uno por préstamo con datos cargados; ver [Adquisiciones críticas por operación](#adquisiciones-críticas-por-operación)) |
| `PMR_Historico.xlsx` | Histórico de ciclos PMR, una hoja por operación (ver [Histórico de PMR por operación](#histórico-de-pmr-por-operación)) |
| `README.md` | Este archivo |

### Hojas del Excel

| Hoja | Contenido |
|------|-----------|
| `productos criticos` | Estado de productos por operación |
| `productos_adquisiciones` | Relación producto crítico ↔ ID de proceso de adquisición, con marca de criticidad y riesgo |
| `equipo` | Roles y personas asignadas por operación |
| `documentos` | Matriz de documentos con links a SharePoint |
| `PMR` | Calificaciones y plan de acción PMR |
| `datos_ops` | Catálogo de operaciones (código, préstamo y nombre) |
| `resultados` | Indicadores de la Matriz de Resultados |
| `productos_resultados` | Relación indicador ↔ producto |
| `comentarios_clausulas` | Comentarios asociados a cláusulas vencidas |
| `change log` | Historial de cambios |

## Stack tecnológico

- HTML5 + JavaScript (ES6+), sin framework ni servidor
- [SheetJS v0.18.5](https://sheetjs.com/) — lectura de Excel en el navegador
- IndexedDB — caché local de datos (Excel principal, CSV y raw data por operación)
- GitHub Pages — hosting estático
