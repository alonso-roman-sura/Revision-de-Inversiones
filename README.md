# Revision-de-Inversiones

Automatización en VBA para la revisión de inversiones en Excel, con soporte para importación de archivos fuente, transformación mediante Power Query, contraste de operaciones de Constitución de Títulos Únicos contra tasas pasivas publicadas por la SBS y limpieza integral del entorno de trabajo.

## Objetivo

Este proyecto concentra el código necesario para replicar una solución de revisión de inversiones en un entorno Excel compatible con VBA y Power Query.

La automatización permite:

- importar archivos de inversiones con estructura validada;
- reconstruir tablas, consultas y vistas derivadas;
- generar una vista específica para instrumentos de tipo Acciones y ETF;
- importar tasas pasivas por tipo de persona;
- contrastar la tasa de operaciones de Constitución de Títulos Únicos contra la tasa de referencia SBS aplicable según banco, moneda, duración y fecha de operación;
- refrescar el entorno de forma controlada;
- reiniciar el entorno de trabajo mediante una rutina de limpieza explícita;
- documentar la dependencia externa que abastece la base de tasas pasivas.

## Alcance del repositorio

Este repositorio versiona el código VBA y la lógica de transformación asociada al proceso. Su finalidad es servir como base técnica para replicar la solución en un entorno propio, sin exponer archivos operativos ni información confidencial.

La lógica está organizada en cuatro módulos funcionales:

- `modActualizar.bas`
- `modInversiones.bas`
- `modTasas.bas`
- `modLimpiarDatos.bas`

## Estructura

```text
Revision-de-Inversiones/
├─ Excel/
│  └─ Módulos/
│     ├─ modActualizar.bas
│     ├─ modInversiones.bas
│     ├─ modTasas.bas
│     └─ modLimpiarDatos.bas
└─ README.md
````

## Arquitectura funcional

El flujo del proyecto está dividido en cuatro bloques principales:

1. actualización controlada del entorno;
2. importación y transformación de inversiones;
3. importación y evaluación de tasas SBS;
4. limpieza integral del entorno para reiniciar el proceso.

Cada módulo cumple una responsabilidad específica y desacoplada, lo que facilita el mantenimiento, la trazabilidad y la réplica del proceso.

## Módulos

### `modActualizar.bas`

Contiene la rutina de actualización general del entorno de trabajo.

**Macro pública principal**

* `ActualizarLibroActivo`

**Responsabilidad**

Refrescar consultas, tablas y conexiones del libro de forma controlada, reduciendo efectos colaterales sobre la sesión de Excel.

**Comportamiento general**

La macro:

1. desactiva actualización de pantalla, eventos y cálculo automático;
2. ejecuta `RefreshAll`;
3. espera la finalización de consultas asíncronas cuando corresponde;
4. restituye la configuración original de Excel.

Este módulo cumple una función transversal y puede ejecutarse como utilitario de refresco cuando el proceso requiere sincronizar salidas, consultas y conexiones.

### `modInversiones.bas`

Es el módulo principal del flujo de importación y transformación de inversiones.

**Macro pública principal**

* `ImportarInversionesDesdeXls`

**Responsabilidad**

Validar el archivo fuente de inversiones, reconstruir la capa RAW, generar consultas de transformación y cargar las tablas finales utilizadas por el proceso.

**Proceso implementado**

La macro realiza, en términos generales, las siguientes acciones:

1. solicita al usuario un archivo de inversiones;
2. abre el archivo en modo lectura;
3. rechaza archivos de texto, HTML u otros formatos que no correspondan a un libro Excel nativo;
4. valida la estructura del origen contra un conjunto estricto de encabezados;
5. copia la data al entorno de trabajo;
6. crea o reconstruye la tabla `Inversiones_Raw`;
7. genera o actualiza consultas de Power Query;
8. carga la tabla final `Inversiones`;
9. limpia espacios y caracteres no separables;
10. corrige fechas en formato texto cuando corresponde;
11. aplica formato numérico a columnas relevantes;
12. construye la hoja `Datos`;
13. genera la vista `Acciones ETF`;
14. reordena hojas principales y deja el entorno listo para revisión.

**Formato esperado del archivo de inversiones**

El formato de entrada se encuentra definido y validado por el propio código. El archivo de origen debe contener los siguientes encabezados:

1. `Portafolio`
2. `Codigo de Orden`
3. `Fecha de Operacion`
4. `Fecha Liquidacion`
5. `Fecha fin Contrato`
6. `Codigo ISIN`
7. `Codigo SBS`
8. `Monto de Operacion Original`
9. `Monto de Operacion ML`
10. `Cantidad`
11. `Precio`
12. `Codigo de Emisor`
13. `Operacion`
14. `Moneda`
15. `Nemonico`
16. `Codigo de Tercero`
17. `Tercero`
18. `Monto Nominal Operacion Original`
19. `Monto Nominal Operacion ML`
20. `Total de Comisiones`
21. `Plaza`
22. `Tipo Tasa`
23. `Porcentaje Tasa`

La validación admite dos tolerancias específicas:

* acepta la variante `Pocentaje Tasa` además de `Porcentaje Tasa`;
* admite que la fila de encabezados se encuentre en la fila 1 o en la fila 3.

**Objetos construidos o actualizados**

**Tablas principales**

* `Inversiones_Raw`
* `Inversiones`

**Vistas derivadas**

* hoja `Datos`
* tabla de trabajo en `Datos`
* hoja `Acciones ETF`
* tabla `Acciones ETF`

**Transformaciones relevantes**

Entre las transformaciones implementadas por el módulo se incluyen:

* normalización de texto y recorte de espacios;
* eliminación de caracteres no imprimibles o no separables;
* corrección de fechas importadas como texto;
* preservación como texto de columnas que no deben forzarse a tipo numérico o fecha;
* aplicación de formato a montos, cantidades, precios y porcentaje de tasa;
* construcción programática de consultas Power Query;
* carga de resultados mediante `Microsoft.Mashup.OleDb.1`.

### `modTasas.bas`

Este módulo gestiona la importación de tasas pasivas por tipo de persona y la evaluación de operaciones de Constitución de Títulos Únicos con base en información SBS.

**Macro pública principal**

* `ImportarTasasPasivasTipoPersona`

**Responsabilidad**

Incorporar la base de tasas al entorno de trabajo, reconstruir las consultas asociadas y generar una capa de validación para operaciones que requieren contraste contra la tasa de referencia.

**Proceso implementado**

La macro ejecuta, de forma resumida, las siguientes acciones:

1. solicita un archivo Excel con tasas pasivas;
2. exige que el archivo contenga una tabla llamada `TasaPasivaTipoPersona`;
3. copia la data origen como tabla RAW;
4. genera o actualiza la consulta `Tasas Pasivas Tipo Persona`;
5. genera o actualiza la consulta `Constitución Títulos Únicos`;
6. carga ambas salidas en hojas del entorno;
7. agrega columnas calculadas y validaciones sobre la hoja de contraste.

**Formato esperado del archivo de tasas**

El archivo seleccionado debe contener una tabla llamada exactamente:

* `TasaPasivaTipoPersona`

A partir de esa tabla, el proceso genera la capa RAW y reconstruye las salidas necesarias para el contraste de operaciones de Constitución de Títulos Únicos.

**Objetos construidos o actualizados**

**Entrada intermedia**

* `TasaPasivaTipoPersona_Raw`

**Salida de tasas**

* hoja `Tasas Pasivas Tipo Persona`
* tabla `Tasas_Pasivas_Tipo_Persona`

**Salida de contraste**

* hoja `Constitución Títulos Únicos`
* tabla `Constitucion_Titulos_Unicos`

**Lógica aplicada a `Constitución Títulos Únicos`**

La consulta asociada parte de la tabla `Inversiones` y selecciona únicamente las filas cuya operación sea `CONSTITUCION TITULOS UNICOS`. Posteriormente elimina columnas no necesarias y ordena el resultado por fecha de operación y código de orden.

Sobre la salida final se generan estas columnas calculadas:

* `Duracion Dias`
* `Duracion Intervalo`
* `Tasa SBS`
* `Resultado`

**Criterio de duración**

La duración de cada operación se clasifica en los siguientes tramos:

* `Hasta 30 días`
* `31-90 días`
* `91-180 días`
* `181-360 días`
* `Más de 360 días`

**Funciones públicas relevantes**

**`TasaSBS_Valor(Tercero, Moneda, Duracion, FechaOperacion)`**

Obtiene la tasa de referencia aplicable según:

* banco normalizado a partir del tercero;
* moneda;
* tramo de duración;
* fecha de operación.

Puede devolver:

* una tasa numérica;
* `El banco no es peruano`;
* `No encontrado`.

**`ResultadoTasaSBS(tasa, valor)`**

Compara la tasa SBS obtenida contra `Porcentaje Tasa` de la operación evaluada.

Puede devolver:

* `Dentro del rango`
* `Fuera del rango`
* `El banco no es peruano`
* `No encontrado`

El umbral actualmente aplicado corresponde a un margen de ±1 punto respecto de la tasa encontrada.

### `modLimpiarDatos.bas`

Este módulo incorpora una rutina de limpieza integral para reiniciar el entorno de trabajo antes de una nueva carga o cuando se requiere eliminar resultados previos.

**Macro pública principal**

* `LimpiarDatos`

**Responsabilidad**

Eliminar de forma controlada todas las salidas del proceso, preservando únicamente las hojas base del entorno.

**Comportamiento general**

La macro:

1. verifica que existan las hojas `Inicio` y `ListaBancos`;
2. solicita confirmación explícita al usuario antes de ejecutar la limpieza;
3. elimina todas las hojas, excepto `Inicio` y `ListaBancos`;
4. elimina todas las consultas del libro;
5. elimina todas las conexiones del libro;
6. elimina nombres definidos rotos que contengan referencias inválidas;
7. restituye el estado de Excel al finalizar.

**Alcance de la limpieza**

La rutina está diseñada para dejar el entorno en un estado base, conservando únicamente:

* `Inicio`
* `ListaBancos`

Todo el resto del contenido generado por el proceso se considera regenerable y es eliminado durante la ejecución.

**Consideración operativa**

Se trata de una limpieza deliberadamente agresiva. Antes de utilizarla, debe asumirse que las salidas construidas por el flujo serán regeneradas posteriormente a partir de nuevas importaciones.

## Dependencias entre módulos, hojas, tablas, consultas y conexiones

Esta sección resume el contrato técnico entre los distintos componentes del proceso.

### Dependencias del flujo de inversiones

`modInversiones.bas`:

* consume un archivo de inversiones con encabezados válidos;
* genera la tabla `Inversiones_Raw`;
* crea o actualiza la lógica Power Query asociada a inversiones;
* produce la tabla `Inversiones`;
* construye la hoja `Datos`;
* construye la hoja `Acciones ETF`.

### Dependencias del flujo de tasas

`modTasas.bas`:

* consume un archivo que contiene la tabla `TasaPasivaTipoPersona`;
* genera la tabla `TasaPasivaTipoPersona_Raw`;
* crea o actualiza la salida `Tasas_Pasivas_Tipo_Persona`;
* depende de la existencia de la tabla `Inversiones`;
* genera la hoja y tabla de `Constitución Títulos Únicos`;
* agrega fórmulas y validaciones sobre la salida final.

### Dependencias de actualización

`modActualizar.bas`:

* no genera estructuras nuevas;
* opera sobre las consultas, tablas y conexiones ya existentes;
* se utiliza como mecanismo de sincronización del entorno.

### Dependencias de limpieza

`modLimpiarDatos.bas`:

* asume la existencia de `Inicio` y `ListaBancos`;
* elimina el resto de hojas;
* elimina consultas del libro;
* elimina conexiones del libro;
* elimina nombres definidos rotos;
* deja el entorno preparado para reconstrucción posterior.

## Dependencia externa documentada

Este repositorio utiliza como insumo operativo una fuente externa para la generación de tasas pasivas. La integración se documenta con fines de trazabilidad técnica.

### Repositorio externo

`Hardnyx/webscrapper`

### Ruta relevante

`SBS/Tasa pasiva`

### Archivos relevantes

* `script.py`
* `empaquetar.py`
* `empaquetar.bat`

### Rol de esta dependencia

El componente externo se encarga de obtener y consolidar la información de tasas pasivas de depósitos a plazo desde la SBS, exportándola a archivos Excel que sirven como insumo para el flujo implementado en `modTasas.bas`.

### Contrato funcional de integración

La relación entre este repositorio y el scraper es documental y funcional, no estructural dentro de Git.

Este repositorio no incorpora el scraper como submódulo ni como subtree. En su lugar, documenta la dependencia externa y asume como contrato de integración la estabilidad del archivo de salida consumido por el proceso de Excel.

Bajo este enfoque, eventuales cambios en la fuente original, como ajustes en la página de la SBS, deben resolverse dentro del scraper sin alterar necesariamente el flujo del lado de Excel, siempre que la estructura de salida continúe siendo compatible con la lógica implementada en `modTasas.bas`.

## Flujo sugerido de uso

### 1. Preparar o actualizar la base de tasas

Ejecutar el scraper externo o su ejecutable para obtener el archivo Excel de tasas pasivas actualizado.

### 2. Preparar el entorno de trabajo en Excel

Disponer de un entorno Excel compatible con VBA y Power Query en el que se replique la estructura requerida por los módulos.

### 3. Importar inversiones

Ejecutar la macro:

```vb
ImportarInversionesDesdeXls
```

### 4. Importar tasas pasivas

Ejecutar la macro:

```vb
ImportarTasasPasivasTipoPersona
```

### 5. Refrescar consultas, tablas y conexiones cuando sea necesario

Ejecutar la macro:

```vb
ActualizarLibroActivo
```

### 6. Reiniciar el entorno cuando se requiera una nueva ejecución limpia

Ejecutar la macro:

```vb
LimpiarDatos
```

## Requisitos

### En Excel

* Microsoft Excel con soporte para VBA.
* Power Query habilitado.
* Permisos para consultas y conexiones.
* Un entorno de trabajo propio en el que se replique la estructura requerida por los módulos.

### Para el scraper externo

* Google Chrome instalado.
* Python y dependencias, si se ejecuta como script.
* O bien el ejecutable generado desde el proceso de empaquetado.

## Uso del código

Este repositorio contiene el código necesario para replicar la automatización del proceso en un entorno propio.

Los módulos pueden utilizarse como base para reconstruir la solución, siempre que se respete la estructura de hojas, tablas, consultas, conexiones y encabezados esperados por la lógica implementada.

Los componentes principales del repositorio son:

1. `modActualizar.bas`
2. `modInversiones.bas`
3. `modTasas.bas`
4. `modLimpiarDatos.bas`

## Consideraciones operativas

* El flujo de inversiones depende de una validación estricta de encabezados y de una estructura de entrada controlada por código.
* El flujo de tasas depende de la existencia de una tabla fuente llamada `TasaPasivaTipoPersona`.
* La evaluación de Constitución de Títulos Únicos se construye sobre la tabla transformada `Inversiones`, no sobre el archivo bruto original.
* La rutina `LimpiarDatos` elimina de forma intencional hojas, consultas y conexiones para restaurar un estado base del entorno.
* La dependencia externa del scraper debe mantenerse alineada con el formato de salida esperado por el proceso.
* Esta documentación describe la lógica versionada en los módulos publicados en el repositorio y la ampliación funcional prevista con `modLimpiarDatos.bas`.

## Mejoras futuras sugeridas

* ampliar la documentación de los objetos intermedios de Power Query, en especial cuando se incorporen nuevas consultas o derivaciones;
* documentar con más detalle las reglas de negocio adicionales que puedan añadirse al contraste de tasas o a la clasificación de instrumentos.

## Estado del proyecto

Repositorio orientado a la automatización operativa de revisión de inversiones y validación de tasas de referencia SBS, con dependencia externa documentada para la obtención de la base de tasas pasivas y con una rutina explícita de limpieza para reinicio controlado del entorno.

```

La parte verificable en el repo público confirma la estructura `Excel/Módulos`, la presencia de `modActualizar.bas`, `modInversiones.bas` y `modTasas.bas`, la macro `ImportarInversionesDesdeXls`, la creación de `Inversiones_Raw`, la vista `Acciones ETF`, la exigencia de la tabla `TasaPasivaTipoPersona`, la salida de `Constitución Títulos Únicos`, los tramos de duración y los estados como `Dentro del rango`, `Fuera del rango`, `El banco no es peruano` y `No encontrado`. También se ve que el scraper externo usa `tkinter`, `selenium`, `BeautifulSoup` y `openpyxl`, y que el empaquetado contempla `PyInstaller`. :contentReference[oaicite:1]{index=1}