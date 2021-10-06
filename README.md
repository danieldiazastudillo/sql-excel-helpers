## Excel a T-SQL Scripts
La idea de este proyecto, es poder contar con algunas funciones en Excel que permitan la generación de scripts SQL para las sentencias más utilizadas, en particular en la resolución de tickets. A través de las celdas necesarias generará, por ejemplo, un string de la siguiente forma:

``` sql
INSERT INTO NombreDeTabla (valor1, valor2) VALUES (100, 200);
```


## Instrucciones de Instalación
Primero que todo se debe verificar qué tipo de instalación de Excel tenemos en nuestro equipo. Por lo general corresponde a la versión de **32-bits**. En caso de contar con la instalación de 64bits se deben escoger, para todos los artefactos, los con nomenclatura **x64**.


1. Descargar ExcelDNA.Intellisense.xll según versión (32/64bits)
2. Descargar una de las versiones de TestDNA desde la sección releases
3. Instalar como complemento en Excel

## Funciones Expuestas

### `SQLUPDATEFECHA`
---
