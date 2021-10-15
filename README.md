## Excel a T-SQL Scripts
La idea de este proyecto, es poder contar con algunas funciones en Excel que permitan la generación de scripts SQL para las sentencias más utilizadas, en particular en la resolución de tickets. A través de las celdas necesarias generará, por ejemplo, un string de la siguiente forma:

``` sql
INSERT INTO NombreDeTabla (valor1, valor2) VALUES (100, 200);
```


## Instrucciones de Instalación
Primero que todo se debe verificar qué tipo de instalación de Excel tenemos en nuestro equipo. Por lo general corresponde a la versión de **32-bits**. En caso de contar con la instalación de 64bits se deben escoger, para todos los artefactos, los con nomenclatura **x64**.


1. Descargar [`ExcelDNA.Intellisense.xll`](https://github.com/Excel-DNA/IntelliSense#getting-started) según versión (32/64bits). **Sin la instalación de este complemento no es posible visualizar en Excel la documentación de ayuda al querer usar las funciones, si bien es opcional, se recomienda**

2. Descargar una de las versiones de `TestDNA` desde [releases](https://github.com/danieldiazastudillo/sql-excel-helpers/releases)
3. Instalar archivo XLL como complemento en Excel. Instrucciones [aquí](https://support.microsoft.com/es-es/office/agregar-o-quitar-complementos-en-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460)

---


# Funciones Expuestas

## UPDATES Simples (Celda Única)
---


### `SQLUPDATEFECHASINGLE`

Genera script para ACTUALIZAR un campo de fecha Excel, formateando a _safe SQL_ 

```sql
CAST('20211015' as datetime)
```

Este formato es seguro para cualquier versión del motor de base de datos con cualquier lenguaje y/o cultura (ES-EN)


Parámetro                  | Descripción
-------------------------- | -------------
Nombre Tabla               | Tabla en la cuál se generará el UPDATE
Campo fecha a modificar    | Nombre del campo que se desea modificar. ej.: fechaDeclaracion
Nueva Fecha                | Valor a actualizar de tipo fecha
Nombre Campo Identificador | Campo Identificador fila. ej.: ProductoID
Valor Campo Identificador  | VALOR del campo Identificador fila. ej.: 15987



### `SQLUPDATEBOOLEANOSINGLE`
Actualiza campos booleanos. Traduce desde inglés o castellano (VERDADERO, FALSO, TRUE, FALSE) o valores numéricos de SQL

Parámetro                  | Descripción
-------------------------- | -------------
Nombre Tabla               | Tabla en la cuál se generará el UPDATE
Campo a modificar          | Nombre del campo que se desea modificar. ej.: `utilizado`
Nuevo Booleano             | Valor a actualizar de tipo booleano (texto, SQL o Access)
Nombre Campo Identificador | Campo Identificador fila. ej.: ProductoID
Valor Campo Identificador  | VALOR del campo Identificador fila. ej.: 123456


### `SQLUPDATEGENERICOSINGLE`
Actualiza campos de cualquier tipo: numéricos, textos, etc. No generará ningún tipo de conversión.

Parámetro                  | Descripción
-------------------------- | -------------
Nombre Tabla               | Tabla en la cuál se generará el UPDATE
Campo a modificar          | Nombre del campo que se desea modificar. ej.: `pesoRestante`
Nuevo Valor                | Texto, número, etc.
Nombre Campo Identificador | Campo Identificador fila. ej.: ProductoID (PK generalmente)
Valor Campo Identificador  | VALOR del campo Identificador fila. ej.: 123456

## Inserciones
Estos mismos tipos (fecha, booleano & genérico) cuentan con un método `INSERT INTO` que sólo requiere del nombre del campo, el valor del campo y el nombre de la tabla con el fin de elaborar el siguiente script:

``` sql
INSERT INTO NombreDeTabla (valor1) VALUES (100);
```

Todos solicitan el nombre de la tabla en la cual insertar el valor, el nombre de la columna y el valor de la columna


## UPDATES & Inserciones por Rango (Múltiples Valores)
Permite la actualización e inserción de múltiples campos en un sólo statement. El sistema NO MANEJA muy bien los nulos por lo que se aconseja limpiar o acomodar los datos según corresponda

### `SQLINSERTINTORANGO`
Genera una instrucción `INSERT INTO` con los datos de un rango. Se pueden anidar funciones auxiliares (Fecha & String SQL) para poder dejar los datos como corresponde. Esta función traduce los valores para booleanos

Parámetro                      | Descripción
------------------------------ | -------------
Nombre Tabla                   | Tabla en la cuál se generará el UPDATE
Rango con NOMBRES de columnas  | Nombre del campo que se desea modificar. ej.: `pesoRestante`
Rango con VALORES de columnas  | Texto, número, etc.


### `SQLUPDATERANGO`
Genera una instrucción `UPDATE` con los datos de un rango. Se pueden anidar funciones auxiliares (Fecha & String SQL) para poder dejar los datos como corresponde. Esta función traduce los valores para booleanos

Parámetro                      | Descripción
------------------------------ | -------------
Nombre Tabla                   | Tabla en la cuál se generará el UPDATE
Rango con NOMBRES de columnas  | Nombre del campo que se desea modificar. ej.: `pesoRestante`
Rango con VALORES de columnas  | Texto, número, etc.
Nombre Campo Identificador     | Campo Identificador fila. ej.: ProductoID (PK generalmente)
Valor Campo Identificador      | VALOR del campo Identificador fila. ej.: 123456

## Eliminación (DELETE)
### `SQLDELETESINGLE`
Genera instrucción `DELETE FROM` de una tabla en particular con su identificador

Parámetro                  | Descripción
-------------------------- | -------------
Nombre Tabla               | Tabla en la cuál se generará el DELETE
Nombre Campo Identificador | Campo Identificador fila. ej.: ProductoID (PK generalmente)
Valor Campo Identificador  | VALOR del campo Identificador fila. ej.: 123456



## Funciones Auxiliares & Anidables


### `SQLFECHAEXCELASAFESQL`
Toma una fecha en formato Excel (cualquiera sea su tipo o cultura) y retorna string con el formato seguro de SQL string, compatible con cualquier versión del motor:

```sql
CAST('20211015' as datetime)
```

Esta función se puede (y debe) anidar en los funciones anteriores


### `SQLTEXTOEXCELASTRINGSQL`
Genera string SQL concatenando comillas simples al inicio y fin del texto

Si el valor en la celda era `HOLA` esta función generará `'HOLA'`



## Solicitud de Cambios
Cualquier solicitud de cambios por favor realizarla vía issues