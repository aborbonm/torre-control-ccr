# Torre de Control — CCR

Dashboard interactivo para el monitoreo de operaciones del Grupo BID en Costa Rica.
Permite visualizar el estado de cláusulas contractuales, productos críticos, equipo,
documentos y calificaciones PMR de forma centralizada y actualizada.

## Acceso

🔗 [aborbonm.github.io/torre-control-ccr](https://aborbonm.github.io/torre-control-ccr/)

## Cómo funciona

El dashboard lee dos archivos de datos al cargarse:

- **`torre_de_control_CCR.xlsx`** — contiene información de productos críticos, equipo,
  documentos, resultados, PMR, comentarios de cláusulas y change log.
- **`Status Date Clause Operation.csv`** — contiene el estado de cláusulas contractuales,
  descargado desde el tablero de PowerBI del BID.

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

### Pasos para subir archivos

1. En este repositorio, haz clic en el archivo que deseas reemplazar
2. Clic en el ícono de lápiz → **"Upload a new version"** → sube el archivo
3. Clic en **"Commit changes"**

El dashboard reflejará los datos nuevos en la próxima visita o al presionar **Actualizar Datos**.

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
| Resultados | xlsx | Indicadores de la Matriz de Resultados agrupados por operación y objetivo, con probabilidad Alta/Media/Baja |
| Productos Críticos | xlsx | Estado de productos con filtro por operación (Logrado / En tiempo / Retrasado) |
| Equipo | xlsx | Composición del equipo por operación con alerta de roles faltantes |
| Documentos | xlsx | Matriz de documentos clave con enlaces a SharePoint por tipo y operación |
| PMR | xlsx | Calificaciones e indicadores técnicos del ciclo PMR I-2026 por programa |
| Change Log | xlsx | Historial de cambios registrados en el Excel (más reciente primero) |

## Estructura del repositorio

| Archivo | Descripción |
|---------|-------------|
| `index.html` | Dashboard completo (aplicación de una sola página, sin servidor) |
| `torre_de_control_CCR.xlsx` | Datos: productos críticos, equipo, PMR, documentos, resultados, comentarios de cláusulas y changelog |
| `Status Date Clause Operation.csv` | Estado de cláusulas contractuales (fuente: PowerBI BID) |
| `README.md` | Este archivo |

### Hojas del Excel

| Hoja | Contenido |
|------|-----------|
| `productos criticos` | Estado de productos por operación |
| `equipo` | Roles y personas asignadas por operación |
| `documentos` | Matriz de documentos con links a SharePoint |
| `PMR` | Calificaciones y plan de acción PMR |
| `datos_ops` | Catálogo de operaciones (código y nombre) |
| `resultados` | Indicadores de la Matriz de Resultados |
| `productos_resultados` | Relación indicador ↔ producto |
| `comentarios_clausulas` | Comentarios asociados a cláusulas vencidas |
| `change log` | Historial de cambios |

## Stack tecnológico

- HTML5 + JavaScript (ES6+), sin framework ni servidor
- [SheetJS v0.18.5](https://sheetjs.com/) — lectura de Excel en el navegador
- IndexedDB — caché local de datos
- GitHub Pages — hosting estático
