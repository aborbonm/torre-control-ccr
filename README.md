\# Torre de Control — CCR



Dashboard interactivo para el monitoreo de operaciones del Grupo BID en Costa Rica.

Permite visualizar el estado de cláusulas contractuales, productos críticos, equipo,

documentos y calificaciones PMR de forma centralizada y actualizada.



\## Acceso



🔗 \[aborbonm.github.io/torre-control-ccr](https://aborbonm.github.io/torre-control-ccr/)



\## Cómo funciona



El dashboard lee dos archivos de datos al cargarse:



\- \*\*`torre_de_control_CCR.xlsx`\*\* — contiene información de productos críticos, equipo,

&#x20; documentos, PMR y change log.

\- \*\*`Status Date Clause Operation.csv`\*\* — contiene el estado de cláusulas contractuales,

&#x20; usado para los tabs de Cláusulas Vencidas y Próximos 180 días.

Los archivos terminados en '- Copy' son respaldos de la última versión estable que permiten volver a ellas en caso de un error.



Los archivos se leen directamente desde este repositorio. No se requiere instalación

ni configuración adicional — basta con abrir el link.



\## Cómo actualizar los datos



1\. Prepara la nueva versión del archivo (`.xlsx` o `.csv`). 

\- \*\*`torre_de_control_CCR.xlsx`\*\* — La información se actualiza directamente en el excel. No cambiar el orden de las columnas en el excel, ni el nombre de las hojas.

\- \*\*`Status Date Clause Operation.csv`\*\* — el archivo se descarga en formato CSV desde el tablero de cláusulas contractuales del BID (https://app.powerbi.com/groups/d9b55ff5-07fe-4a88-8995-d94ed53cd7cf/rdlreports/429c4230-9c53-46ba-8c96-94ec89aa1cd2?experience=power-bi). Antes de descargar se aplican filtro de> Country=Costa Rica; Clause Status=ALL; Operation Stage=Disbursing. 

&#x20;

2\. En este repositorio, haz clic en el archivo que deseas reemplazar

3\. Clic en el ícono de lápiz (editar) → \*\*"Upload a new version"\*\* → sube el archivo

4\. Clic en \*\*"Commit changes"\*\*



El dashboard reflejará los datos nuevos en la próxima visita o al hacer clic

en el botón \*\*Actualizar Datos\*\*.



\## Estructura del repositorio



| Archivo                        | Descripción |
|--------------------------------|---------------------------------------------|
| `index.html`                   | Dashboard interactivo (no requiere servidor) |
| `torre\_de\_control\_CCR.xlsx` | Datos de productos criticos, equipo, PMR, documentos y changelog |
| `Status Date Clause Operation.csv` | Estado de cláusulas contractuales |
| `README.md`                    | Este archivo |



\## Tabs del dashboard



| Tab                | Fuente | Descripción |

|--------------------|------ -|-------------|

| Resumen            | xlsx + csv | KPIs generales y alertas de vencimiento próximo |

| Cláusulas Vencidas | csv    | Cláusulas con status TRACK vencidas a la fecha |

| Próximos 180 días  | csv    | Cláusulas TRACK por vencer en los próximos 180 días |

| Productos Críticos | xlsx   | Estado de productos con filtro por operación y estado |

| Equipo             | xlsx   | Composición del equipo por operación con alertas de roles faltantes |

| Documentos         | xlsx   | Matriz de documentos clave con enlaces a SharePoint |

| PMR                | xlsx   | Calificaciones y plan de acción del ciclo PMR I-2026 |

| Change Log         | xlsx   | Historial de cambios registrados en el Excel |





