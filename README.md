# Revisión de Inversiones

Automatización en Excel, VBA y Python para importar archivos de inversiones, reconstruir tablas y consultas de Power Query, generar vistas de revisión y contrastar operaciones de Constitución de Títulos Únicos contra tasas pasivas de referencia de la SBS.

## Propósito del repositorio

Este repositorio está documentado principalmente para programadores o mantenedores técnicos que necesiten:

- entender la arquitectura general del proceso;
- importar la lógica a un libro propio;
- adaptar módulos, tablas, consultas y validaciones;
- mantener el flujo sin romper dependencias internas.

También incluye una sección breve de uso operativo para ejecutar el proceso en el orden correcto, pero no está planteado como un manual completo para usuario final.

## Qué resuelve

La solución automatiza un flujo de revisión en Excel que cubre cuatro frentes principales:

- importación y validación de archivos fuente de inversiones;
- reconstrucción de tablas y consultas de Power Query;
- evaluación de operaciones de Constitución de Títulos Únicos contra tasas pasivas SBS;
- limpieza del entorno para reiniciar una nueva ejecución.

## Alcance

El repositorio publica la lógica del proceso y su estructura técnica.

No incluye archivos operativos, datos confidenciales ni libros de trabajo productivos.

## Estructura del repositorio

```text
Revision-de-Inversiones/
├─ Excel/
│  └─ Módulos/
│     ├─ modActualizar.bas
│     ├─ modInversiones.bas
│     ├─ modTasas.bas
│     └─ modLimpiarDatos.bas
├─ Python/
│  └─ Tasa pasiva SBS/
│     ├─ empaquetar.bat
│     ├─ empaquetar.py
│     └─ script.py
└─ README.md
````

## Arquitectura funcional

El flujo está dividido en cuatro bloques:

1. actualización controlada del libro;
2. importación y transformación de inversiones;
3. importación y evaluación de tasas pasivas;
4. limpieza integral del entorno.

Cada módulo tiene una responsabilidad específica. La intención es mantener el proceso desacoplado para facilitar mantenimiento, trazabilidad y ajustes puntuales.

## Módulos VBA

### `modActualizar.bas`

**Macro pública principal:** `ActualizarLibroActivo`

**Responsabilidad**

Refrescar consultas, tablas y conexiones del libro de forma controlada.

**Rol dentro del proceso**

* ejecuta la actualización global del libro;
* sincroniza salidas ya construidas;
* sirve como utilitario transversal cuando el entorno requiere refresco sin reconstrucción completa.

### `modInversiones.bas`

**Macro pública principal:** `ImportarInversionesDesdeXls`

**Responsabilidad**

Validar el archivo fuente de inversiones, reconstruir la capa RAW, regenerar consultas y cargar tablas finales para la revisión.

**Qué hace en términos generales**

* solicita un archivo Excel de inversiones;
* valida encabezados y estructura;
* copia la data al entorno de trabajo;
* crea o reconstruye `Inversiones_Raw`;
* genera o actualiza consultas de Power Query;
* carga la tabla final `Inversiones`;
* limpia espacios y caracteres no separables;
* corrige fechas importadas como texto;
* construye la hoja `Datos`;
* genera la vista `Acciones ETF`.

**Supuestos importantes**

* el origen debe ser un libro Excel válido;
* no se aceptan archivos de texto o HTML como sustituto del origen principal;
* la validación de encabezados es estricta;
* se admite la variante `Pocentaje Tasa` además de `Porcentaje Tasa`;
* los encabezados pueden encontrarse en la fila 1 o en la fila 3.

### `modTasas.bas`

**Macro pública principal:** `ImportarTasasPasivasTipoPersona`

**Responsabilidad**

Importar la base de tasas pasivas por tipo de persona, reconstruir la salida asociada y generar la validación de Constitución de Títulos Únicos.

**Qué hace en términos generales**

* incorpora la tabla RAW de tasas al libro;
* genera la salida de tasas pasivas;
* construye la hoja de contraste para operaciones relevantes;
* agrega columnas de cálculo y validación como duración, tasa de referencia y resultado.

**Dependencia clave**

Este módulo parte de la información ya transformada de inversiones. No debe considerarse un flujo completamente independiente.

### `modLimpiarDatos.bas`

**Macro pública principal:** `LimpiarDatos`

**Responsabilidad**

Eliminar salidas regenerables y devolver el libro a un estado base para una nueva corrida.

**Qué limpia**

Según la lógica implementada, elimina elementos reconstruibles como hojas derivadas, consultas, conexiones y objetos auxiliares, preservando la base mínima del entorno.

**Supuesto importante**

El entorno base conserva las hojas `Inicio` y `ListaBancos`.

## Componente auxiliar en Python

La carpeta `Python/Tasa pasiva SBS` contiene una utilidad separada del núcleo VBA.

### Archivos

* `script.py`: obtiene y exporta la información de tasas pasivas;
* `empaquetar.py`: construye un ejecutable del script;
* `empaquetar.bat`: lanza el proceso de empaquetado.

### Rol dentro del proyecto

Este componente existe para abastecer o facilitar la obtención de la base de tasas pasivas usada por la lógica en Excel.

## Objetos y salidas relevantes

Entre los objetos que el flujo crea, reconstruye o utiliza se encuentran:

* `Inversiones_Raw`;
* `Inversiones`;
* hoja `Datos`;
* vista `Acciones ETF`;
* `TasaPasivaTipoPersona_Raw`;
* hoja `Tasas Pasivas Tipo Persona`;
* hoja `Constitución Títulos Únicos`.

## Dependencias internas

Antes de modificar el proyecto, conviene tener en cuenta lo siguiente:

* `modTasas.bas` depende del resultado del flujo de inversiones;
* `ActualizarLibroActivo` refresca estructuras existentes, pero no reemplaza el proceso de construcción inicial;
* `LimpiarDatos` debe ejecutarse con cautela, porque elimina salidas regenerables para volver al estado base;
* cambiar nombres de hojas, tablas, consultas o conexiones puede romper referencias internas.

## Requisitos técnicos

### Excel

* Microsoft Excel con soporte para VBA;
* Power Query habilitado;
* permisos para consultas y conexiones;
* un libro de trabajo compatible donde se importen los módulos.

### Python, solo para el auxiliar de tasas

* Python 3;
* dependencias requeridas por `script.py`;
* `PyInstaller` si se desea generar ejecutable mediante `empaquetar.py`.

## Uso básico

Esta sección resume el orden operativo general. No sustituye una revisión del código cuando se desee adaptar el proyecto.

### 1. Importar los módulos al libro de trabajo

Incorporar los archivos `.bas` al proyecto VBA del libro donde se ejecutará la solución.

### 2. Cargar inversiones

Ejecutar:

```vb
ImportarInversionesDesdeXls
```

Resultado esperado:

* creación o reconstrucción de la capa RAW;
* generación de la tabla `Inversiones`;
* creación de vistas derivadas como `Datos` y `Acciones ETF`.

### 3. Cargar tasas pasivas

Ejecutar:

```vb
ImportarTasasPasivasTipoPersona
```

Resultado esperado:

* incorporación de la base de tasas;
* generación de la salida de contraste para Constitución de Títulos Únicos.

### 4. Refrescar el libro

Ejecutar:

```vb
ActualizarLibroActivo
```

Usar esta macro cuando se necesite sincronizar consultas, tablas y conexiones ya construidas.

### 5. Reiniciar el entorno, cuando corresponda

Ejecutar:

```vb
LimpiarDatos
```

Usar esta macro solo cuando se quiera limpiar salidas regenerables y volver al estado base del libro.

## Qué debe leer cada tipo de lector

### Si se va a mantener o modificar el proyecto

Revisar principalmente:

* estructura del repositorio;
* módulos y responsabilidades;
* dependencias internas;
* objetos construidos y nombres esperados.

### Si solo se va a ejecutar el flujo

Basta con revisar:

* requisitos mínimos;
* sección de uso básico;
* orden de ejecución de macros.

[1]: https://github.com/alonso-roman-sura/Revision-de-Inversiones "GitHub - alonso-roman-sura/Revision-de-Inversiones · GitHub"
