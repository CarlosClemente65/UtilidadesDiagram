using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;


namespace UtilidadesDiagram
{
    /// <summary>
    /// Utilidades que se utilizan para las aplicaciones Diagram
    /// </summary>
    public static class Utilidades
    {
        // Extensiones Excel validas
        private static readonly HashSet<string> extensionesValidas = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx",
            ".xls"
        };

        /// <summary>
        /// Metodo para quitar los acentos a las vocales y simbolos raros
        /// </summary>
        /// <param name="cadena"></param>
        /// <returns>Devuelve una cadena en la que se han sustituido los caracteres raros</returns>
        public static string QuitaRaros(string cadena)
        {
            Dictionary<char, char> caracteresReemplazo = new Dictionary<char, char>
            {
                {'á', 'a'}, {'é', 'e'}, {'í', 'i'}, {'ó', 'o'}, {'ú', 'u'},
                {'Á', 'A'}, {'É', 'E'}, {'Í', 'I'}, {'Ó', 'O'}, {'Ú', 'U'},
                {'¥', 'Ñ'}, {'¤', 'ñ'},
                {'´', '\'' } // Reemplaza el simbolo ´por el '
                //{'\u00AA', '.'}, {'ª', '.'}, {'\u00BA', '.'}, {'°', '.' }
            };

            StringBuilder resultado = new StringBuilder(cadena.Length);
            foreach(char c in cadena)
            {
                if(caracteresReemplazo.TryGetValue(c, out char reemplazo))
                {
                    resultado.Append(reemplazo);
                }
                else
                {
                    resultado.Append(c);
                }
            }

            return resultado.ToString();
        }

        /// <summary>
        /// Se borra el fichero pasado, comprobando antes si existe
        /// </summary>
        /// <param name="fichero"></param>
        /// <returns>Devuelve 'true' si se ha borrado el fichero</returns>
        //Controla si existe el fichero para borrarlo
        public static bool ControlFicheros(string fichero)
        {
            bool resultado = false;
            if(File.Exists(fichero))
            {
                File.Delete(fichero);
                resultado = true;
            }
            return resultado;
        }

        /// <summary>
        /// Metodo para grabar un texto en un fichero con la codificacion 'Default' que es la que se usa en el basico
        /// </summary>
        /// <param name="fichero"></param>
        /// <param name="texto"></param>
        /// <returns>Devuelve 'true' si se ha grabado el fichero</returns>
        //Metodo para grabar el fichero en la ruta que se pase
        public static bool GrabarFichero(string fichero, string texto)
        {
            bool resultado = false;
            try
            {
                File.WriteAllText(fichero, texto, Encoding.Default);
                resultado = true;
            }
            catch
            {
                resultado = false;
            }
            return resultado;
        }

        /// <summary>
        /// Permite dividir una cadena por el caracter pasado, dividiendola en un maximo de 2 partes (divide desde el primer divisor que encuentra)
        /// </summary>
        /// <param name="cadena"></param>
        /// <param name="divisor"></param>
        /// <returns>Devuelve dos cadenas con el 'atributo' y 'valor'</returns>
        public static (string, string) DivideCadena(string cadena, char divisor)
        {
            string atributo = string.Empty;
            string valor = string.Empty;
            string[] partes = cadena.Split(new[] { divisor }, 2);
            if(partes.Length == 2)
            {
                atributo = partes[0].Trim();
                valor = partes[1].Trim();
            }

            return (atributo, valor);
        }

        /// <summary>
        /// Metodo que chequea si la version de .NET Framework pasada es inferior a la instalada
        /// Versiones disponibles: 4.7 (numero 460798), 4.7.1 (numero 461308), 4.7.2 = 461808, 4.8 = (numero 528040) y 4.8.1 (numero 533320)
        /// </summary>
        /// <param name="clave"></param>
        /// <returns>Devuelve 'true' si la version instalada es superior a la necesaria</returns>
        public static bool ChequeoFramework(string clave)
        {
            //Diccionario con los valores de las distintas versiones
            Dictionary<string, int> Versiones = new Dictionary<string, int>()
            {
                {"4.7", 460798 },
                {"4.7.1", 461308 },
                {"4.7.2", 461808 },
                {"4.8", 528040},
                {"4.8.1", 53320 }
            };

            //Se inicializa a cero por si no esta instalado el Framework en modo 'Release'
            int versionNecesaria = 0;

            //Se obtiene el valor de la version del diccionario segun la cadena pasada
            Versiones.TryGetValue(clave, out versionNecesaria);

            //Se inicializa la version instalada antes de leerla del registro
            int versionInstalada = 0;
            using(RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
            .OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"))
            {
                //Si el valor no es nulo, y aparece la clave 'Release'
                if(ndpKey != null && ndpKey.GetValue("Release") != null)
                {
                    versionInstalada = (int)ndpKey.GetValue("Release");
                }
            }
            return versionInstalada >= versionNecesaria;
        }

        /// <summary>
        /// Permite borrar todos los ficheros sin tener en cuenta la extension del fichero pasado, como si fuera 'delete fichero.*'; puede ser util para procesos que pueden generar ficheros de varios tipos (html, pdf, txt, etc).
        /// </summary>
        /// <param name="fichero">Nombre del fichero a borrar</param>
        /// <returns>Devuelve la lista de ficheros eliminados</returns>
        public static StringBuilder BorrarFicheros(string fichero)
        {
            StringBuilder ficheros = new StringBuilder();
            string rutaSalida = string.Empty;
            if(!string.IsNullOrEmpty(fichero))
            {
                rutaSalida = Path.GetDirectoryName(fichero);
            }

            //Si no se ha pasado la ruta completa del fichero, se carga la de ejecucion del programa
            if(string.IsNullOrEmpty(rutaSalida))
            {
                rutaSalida = Directory.GetCurrentDirectory();
            }

            //Carga todos los ficheros sin importar la extension
            string patronFicheros = Path.GetFileNameWithoutExtension(fichero) + ".*";
            string[] elementos = Directory.GetFiles(rutaSalida, patronFicheros);

            foreach(string elemento in elementos)
            {
                File.Delete(elemento);
                ficheros.AppendLine(elemento);
            }
            return ficheros;
        }

        /// <summary>
        /// Permite convertir un fichero.xls (Excel 97-2003) a fichero.xlsx (Excel 2007)
        /// </summary>
        /// <param name="ficheroXls">Ruta del fichero a convertir</param>
        /// <returns>Devuelve un "MemoryStream" con el contenido convertido</returns>
        public static MemoryStream ConvertirAXlsx(string ficheroXls)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook libro = excelApp.Workbooks.Open(ficheroXls);

            // Crear un archivo temporal en formato .xlsx
            string tempFilePath = Path.GetTempFileName() + ".xlsx";
            libro.SaveAs(tempFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook);

            // Cerrar y liberar recursos
            libro.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(libro);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            // Leer el archivo temporal en un MemoryStream
            MemoryStream ms = new MemoryStream(File.ReadAllBytes(tempFilePath));

            // Eliminar el archivo temporal
            File.Delete(tempFilePath);

            return ms;

        }


        /// <summary>
        /// Chequea si el fichero pasado es un Excel (extension xlsx o xls)
        /// </summary>
        /// <param name="fichero"></param>
        /// <returns>true si el fichero es un Excel</returns>
        public static bool EsFicheroExcel(string fichero)
        {
            if(string.IsNullOrEmpty(fichero))
            {
                return false;
            }

            return extensionesValidas.Contains(Path.GetExtension(fichero));

        }
    }
}
