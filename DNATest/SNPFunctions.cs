using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.ObjectModel;

namespace DNATest
{
    public class SNPFunctions
    {
        /// <summary>
        /// Diccionario con los valores que corresponden a booleanos. Permite revisar si el string en cuestión contiene alguno de estos datos
        /// </summary>
        public static readonly IList<string> valoresBoolean = new ReadOnlyCollection<string> (new List<string> { "VERDADERO", "TRUE", "FALSO", "FALSE" });



        [ExcelFunction(Description = "Genera T-SQL para UPDATE con un valor simple de tipo FECHA")]
        public static string SQLUPDATEFECHA(
            [ExcelArgument(Name = "Nombre Base de Datos", Description = "Base de datos en la cuál se generará el UPDATE")] string nombreBD,
            [ExcelArgument(Name = "Campo fecha a modificar", Description = "Nombre del campo que se desea modificar. ej.: fechaDeclaracion")] string campoModificacion,
            [ExcelArgument(Name ="Nueva Fecha", Description = "Valor a actualizar de tipo fecha")] DateTime nuevaFecha,                        
            [ExcelArgument(Name = "Nombre Campo Identificador", Description = "Campo Identificador fila. ej.: ProductoID")] string campoID,
            [ExcelArgument(Name = "valor Campo Identificador", Description = "VALOR del campo Identificador fila. ej.: 15987")] string valorCampoID)
        {                        
            return $"UPDATE {nombreBD.Trim()} SET {campoModificacion.Trim()} = {StringToSafeSQLDate(nuevaFecha)} WHERE {campoID.Trim()} =  {valorCampoID.Trim()};";
        }

        [ExcelFunction(Description = "Genera T-SQL para INSERT con un valor simple de tipo FECHA")]
        public static string SQLINSERTFECHA(
            [ExcelArgument(Name = "Nombre Base de Datos", Description = "Sin prefijos como dbo o dbo_*. Ej.: SigcuePlantaDestino")] string nombreBD,
            [ExcelArgument(Name = "Nombre Columna", Description = "Nombre columna donde se hará la inserción")] string nombreColumna,
            [ExcelArgument(Name = "Nueva Fecha", Description = "Valor para ser agregado al INSERT")] DateTime nuevafecha)
        {
            return $"INSERT INTO {nombreBD.Trim()} ({nombreColumna}) VALUES ({StringToSafeSQLDate(nuevafecha)});";
        }

        [ExcelFunction(Description = "Genera T-SQL para UPDATE con un valor simple de tipo BOOLEANO (VERDADERO/FALSO)")]
        public static string SQLUPDATEBOOLEANO(
            [ExcelArgument(Name = "Nombre Base de Datos", Description = "Base de datos en la cuál se generará el UPDATE")] string nombreBD,
            [ExcelArgument(Name = "Campo a Modificar", Description = "Nombre del campo que se desea modificar. ej.: utilizado")] string campoModificacion,
            [ExcelArgument(Name = "Nuevo Booleano", Description = "Puede ser texto VERDADERO/TRUE o FALSO/FALSE")] string nuevoBooleano,                        
            [ExcelArgument(Name = "Nombre Campo Identificador", Description = "Campo Identificador fila. ej.: ProductoID")] string campoID,
            [ExcelArgument(Name = "valor Campo Identificador", Description = "VALOR del campo Identificador fila. ej.: 15987")] string valorCampoID)
        {
            string boolAsSQLString = translateBoolean(nuevoBooleano);

            return $"UPDATE {nombreBD.Trim()} SET {campoModificacion.Trim()} = {boolAsSQLString} WHERE {campoID.Trim()} =  {valorCampoID.Trim()};";
        }

        [ExcelFunction(Description = "Genera T-SQL para INSERT con un valor simple de tipo BOOLEANO (VERDADERO/FALSO)")]
        public static string SQLINSERTBOOLEANO(
            [ExcelArgument(Name = "Nombre Base de Datos", Description = "Base de datos en la cuál se generará el INSERT")] string nombreBD,
            [ExcelArgument(Name = "Nombre Columna", Description = "Nombre de la columna donde se hará la inserción")] string nombreColumna,
            [ExcelArgument(Name = "Nuevo Booleano", Description = "Puede ser texto VERDADERO/TRUE o FALSO/FALSE")] string nuevoBooleano)                                    
        {
            string boolAsSQLString = translateBoolean(nuevoBooleano);

            return $"INSERT INTO {nombreBD.Trim()} ({nombreColumna.Trim()}) VALUES ({boolAsSQLString});";
        }


        [ExcelFunction(Description = "Genera T-SQL para UPDATE con un valor simple de cualquier tipo, no generará conversiones")]
        public static string SQLUPDATEGENERICO(
            [ExcelArgument(Name = "Nombre Base de Datos", Description = "Base de datos en la cuál se generará el UPDATE")] string nombreBD,
            [ExcelArgument(Name = "Campo a Modificar", Description = "Nombre del campo que se desea modificar. ej.: utilizado")] string campoModificacion,
            [ExcelArgument(Name = "Nuevo Valor", Description = "Permite texto, números o cualquier campo que no requiera modificación")] string nuevoValor,                        
            [ExcelArgument(Name = "Nombre Campo Identificador", Description = "Campo Identificador fila. ej.: ProductoID")] string campoID,
            [ExcelArgument(Name = "Valor Campo Identificador", Description = "VALOR del campo Identificador fila. ej.: 15987")] string valorCampoID)
        {            
            return $"UPDATE {nombreBD.Trim()} SET {campoModificacion.Trim()} = {nuevoValor} WHERE {campoID.Trim()} =  {valorCampoID.Trim()};";
        }


        [ExcelFunction(Description = "Genera T-SQL para INSERT con un valor simple de cualquier tipo, no generará conversiones")]
        public static string SQLINSERTGENERICO(
            [ExcelArgument(Name = "Nombre Base de Datos", Description = "Base de datos en la cuál se generará el INSERT")] string nombreBD,
            [ExcelArgument(Name = "Nombre Columna", Description = "Nombre de la columna donde se insertará el dato")] string nombreColumna,
            [ExcelArgument(Name = "Nuevo Valor", Description = "Permite texto, números o cualquier campo que no requiera modificación")] string nuevoValor)            
        {


            return $"INSERT INTO {nombreBD.Trim()} ({nombreColumna.Trim()}) VALUES ({nuevoValor.Trim()});";
        }


        [ExcelFunction(Description = "Genera T-SQL con conversión de FECHA a string, el string generado es seguro para la inserción en cualquier servidor")]
        public static string FECHASAFESQL([ExcelArgument(Name = "Valor de tipo FECHA", Description = "Debe ser de tipo FECHA")] DateTime fechaExcel)
        {
            return StringToSafeSQLDate(fechaExcel);
        }


        [ExcelFunction(Description = "Genera T-SQL para INSERT con múltiples valores, no generará conversiones")]
        public static string SQLINSERTINTOMULTIPLE(
            [ExcelArgument(Name ="Nombre BD", Description = "Nombre Base de para realizar el INSERT")] string nombreBD,
            [ExcelArgument(AllowReference = true, Name = "Rango con NOMBRES de columnas")] object nombresColumnas,
            [ExcelArgument(AllowReference = true, Name = "Rango con VALORES de columnas")] object valoresColumnas)
        {
            var stringColumnas = ExcelReferenceToString(nombresColumnas).statement;
            var stringValores = ExcelReferenceToString(valoresColumnas).statement;
            

            return $"INSERT INTO {nombreBD.Trim()} ({stringColumnas}) VALUES ({stringValores});";            
        }


        [ExcelFunction(Description = "Genera T-SQL para UPDATE con múltiples valores, no generará conversiones")]
        public static string SQLUPDATEMULTIPLE(
            [ExcelArgument(Name = "Nombre BD", Description = "Nombre Base de para realizar el UPDATE")] string nombreBD,
            [ExcelArgument(AllowReference = true, Name = "Rango con NOMBRES de columnas")] object nombreColumnas,
            [ExcelArgument(AllowReference = true, Name = "Rango con VALORES de columnas")] object valoresColumnas,
            [ExcelArgument(Name = "Nombre columna con ID", Description = "WHERE {columnaId} = 1234")] string columnaID,
            [ExcelArgument(Name = "Valor columna ID", Description = "WHERE ColumnaID = {valorID}")] string valorID)
        {
            var nomColRef = ExcelReferenceToString(nombreColumnas);
            var valColRef = ExcelReferenceToString(valoresColumnas);

            List<string> updatePropsVal = new List<string>();
            
            
            if (nomColRef.objList.Count() == valColRef.objList.Count())
            {

                int counter = 0;

                foreach(var col in nomColRef.objList)
                {
                    string currentValue = $"{col} = {valColRef.objList[counter]}";
                    updatePropsVal.Add(currentValue);
                    counter++;
                }

                return $"UPDATE {nombreBD.Trim()} SET {string.Join(", ", updatePropsVal)} WHERE {columnaID.Trim()} = {valorID.Trim()};";

            }
            else
            {
                throw new Exception("Los rangos de nombres de columna y valores no contienen la misma cantidad de items");
            }
            
        }


        /// <summary>
        /// Convierte un object reference correspondiente a un RANGO de Excel en un satatement T-SQL & una lista con los valores
        /// </summary>
        /// <param name="obj">Rango Excel (object reference true)</param>
        /// <returns>Tupla con statement T-SQL y lista con valores</returns>
        private static (string statement, List<string> objList) ExcelReferenceToString(object obj)
        {
            if(obj is ExcelReference target)
            {
                if (target.GetValue() is object[,] res)
                {
                    return (Obj2dToString(res), Obj2dToStringList(res));
                }
                else
                {
                    return (ObjToString(target.GetValue()), new List<string>{ target.GetValue() as string });
                }                
                
            }
            else
            {
                throw new ArgumentException("Invalid argument in ExcelReferenceToString method");
            }
        }
        
        
        /// <summary>
        /// Convierte un object reference de Excel valores separados por comas
        /// </summary>
        /// <param name="obj">Rango Excel</param>
        /// <param name="separator">Separador para valores</param>
        /// <returns></returns>
        private static string Obj2dToString(object[,] obj, string separator = ", ")
        {
            List<string> result = Obj2dToStringList(obj);
            return string.Join(separator, result);
        }


        /// <summary>
        /// Entrega string con T-SQL. Además verifica si alguno de los valores es BOOLEANO y genera la traducción pertinente
        /// de VERDADERO/TRUE = 1 && FALSO/FALSE = 0
        /// </summary>
        /// <param name="obj">Rango de Excel</param>
        /// <returns></returns>
        private static List<string> Obj2dToStringList(object[,] obj)
        {
            var result = new List<string>();

            foreach (var v in obj)
            {
                if (valoresBoolean.Any(i => v.ToString().Trim().ToUpper().Contains(i)))
                {
                    result.Add(translateBoolean(v.ToString()));
                }
                else
                {
                    result.Add(ObjToString(v));                    
                }
            }

            return result;
        }

        /// <summary>
        /// Convierte una sola celda de Excel (object) en string T-SQL. Verifica nulos.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private static string ObjToString(object obj)
        {            
            return (obj == null) ? "NULL" : obj.ToString();
        }


        /// <summary>
        /// Genera desde un DateTime (celda tipo fecha) una instrucción T-SQL con safe-date (funciona en cualquier collation)
        /// </summary>
        /// <param name="excelDate">Fecha desde Excel</param>
        /// <returns></returns>
        private static string StringToSafeSQLDate(DateTime excelDate)
        {
            return $"CAST('{excelDate.Year}{excelDate.Month:00}{excelDate.Day:00}' as datetime)";

        }


        /// <summary>
        /// Traduce booleano desde VERDADERO/TRUE/FALSO/FALSE a los valores T-SQL correspondientes
        /// </summary>
        /// <param name="val">String (palabra) correspondiente a un booleano</param>
        /// <returns>VERDADERO = 1 || FALSO = 0</returns>
        private static string translateBoolean(string val)
        {
            string newVal = val.Trim().ToUpper();

            // castellano
            if (newVal == "VERDADERO" || newVal == "TRUE")
            {
                return "1";
            }

            // ingles
            if (newVal == "FALSO" || newVal == "FALSE")
            {
                return "0";
            }

            return "";

        }


        //private static Range ToRange(ExcelReference reference)
        //{
        //    var xlApp = ExcelDnaUtil.Application as Application;
        //    var item = XlCall.Excel(XlCall.xlSheetNm, reference) as string;
        //    int index = item.LastIndexOf(']');
        //    item = item.Substring(index + 1);
        //    var ws = xlApp.Sheets[item] as Worksheet;
        //    var target = xlApp.Range[
        //        ws.Cells[reference.RowFirst + 1, reference.ColumnFirst + 1],
        //        ws.Cells[reference.RowLast + 1, reference.ColumnLast + 1]] as Range;

        //    return target;
        //}
    }
}
