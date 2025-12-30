using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Ranges;
using PruebaArbol;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Xml.Linq;
using System.Xml;
using DataRow = PruebaArbol.DataRow;
using System.IO;
using System.Linq;

namespace crucia
{
    public static class Utilidades
    {
        public static void EscribirEnLog(string mensaje)
        {
            // La ruta del log sigue hardcodeada (revisar)
            string rutaLog = "E:\\DescarConector\\Intermedio_log.txt";

            try
            {
                File.AppendAllText(rutaLog, $"{DateTime.Now} - {mensaje}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al escribir en el log: {ex.Message}");
            }
        }

        public static void LimpiarCarpeta(string rutaCarpeta)
        {
            try
            {
                // Asegura que la carpeta de destino exista
                if (!Directory.Exists(rutaCarpeta))
                {
                    Directory.CreateDirectory(rutaCarpeta);
                    Utilidades.EscribirEnLog($"La carpeta de salida {rutaCarpeta} no existía y fue creada.");
                }

                // Limpia el contenido
                foreach (string archivo in Directory.GetFiles(rutaCarpeta))
                {
                    File.Delete(archivo);
                }
                Utilidades.EscribirEnLog($"Carpeta de salida {rutaCarpeta} limpiada con éxito.");
            }
            catch (Exception ex)
            {
                Utilidades.EscribirEnLog($"Error al limpiar o crear la carpeta de salida {rutaCarpeta}: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Utilidades.EscribirEnLog("\n\n----------------------------------------------------------------------------------------------------------------------------------\n\n");

            // 1. VALIDACIÓN Y OBTENCIÓN DE PARÁMETROS
            if (args.Length < 2) // Solo se requieren 2 argumentos ahora
            {
                Utilidades.EscribirEnLog("ERROR: Se requieren dos argumentos:\n1) Ruta del XML de entrada (MBOM)\n2) Ruta de la carpeta de salida (XMLs BOPs).");
                return;
            }

            // Argumento 1: Ruta del archivo XML de entrada (MBOM)
            string archivo = args[0];

            // Argumento 2: Ruta de la carpeta de salida de los XMLs (AllXmls)
            string allXmlsPath = args[1];

            // Aseguramos que la ruta de carpeta termine con un separador
            if (!allXmlsPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                allXmlsPath += Path.DirectorySeparatorChar;
            }

            if (!File.Exists(archivo))
            {
                Utilidades.EscribirEnLog($"ERROR: El archivo XML de entrada no se encontró en la ruta especificada: {archivo}");
                return;
            }

            // Rutas base (HARDCODEADAS por ser configuración de Teamcenter y BAT)
            string directoryBat = "E:\\DescarConector";
            string baseBatName = "script_export_";
            const int NUM_BATCH_FILES = 10; // Número de archivos .bat para paralelización

            // Crear un objeto XDocument y cargar el contenido desde el archivo
            XDocument xdoc = XDocument.Load(archivo);

            // Base para la configuración de Teamcenter en el .bat
            string tcConfigContent = @"@echo off
                                    cd /d ""E:\Siemens\Teamcenter14\tc_menu""
                                    call ""E:\Siemens\Teamcenter14\tc_menu\tc_config1.bat""";

            // Limpiar la carpeta de destino de los XMLs antes de empezar la exportación
            Utilidades.LimpiarCarpeta(allXmlsPath);

            // Lista para almacenar todos los comandos de exportación generados
            List<string> exportCommands = new List<string>();



            // 1. Obtener todos los elementos Product (incluyendo el raíz de la MBOM)
            var allProductElements = xdoc.Descendants("{http://www.plmxml.org/Schemas/PLMXMLSchema}Product");

            // 2. Usar un HashSet para garantizar la unicidad y evitar procesar duplicados
            var uniqueProductsToExport = new HashSet<(string productId, string subType)>();

            foreach (XElement productElement in allProductElements)
            {
                string productId = productElement.Attribute("productId")?.Value;
                string subtype = productElement.Attribute("subType")?.Value;

                if (!string.IsNullOrEmpty(productId))
                {
                    uniqueProductsToExport.Add((productId, subtype));
                }
            }

            // 3. Generar los comandos de exportación a partir de los items únicos
            foreach (var productInfo in uniqueProductsToExport)
            {
                string productId = productInfo.productId;
                string subtype = productInfo.subType;
                string itemToExport = "";

                // Aplicar la lógica de transformación del ID
                if (subtype == "Agm4_Pieza")
                {
                    itemToExport = "P-" + productId;
                }
                else if (subtype == "Agm4_SubCon")
                {
                    if (productId.StartsWith('E'))
                    {
                        itemToExport = "P-" + productId.Substring(1);
                    }
                    else
                    {
                        itemToExport = "P-" + productId;
                    }
                }

                if (!string.IsNullOrEmpty(itemToExport))
                {
                    // Usa la ruta parametrizada para la salida de los XMLs
                    string command = $"\r\nplmxml_export -u=infodba -p=infodba -g=dba -item={itemToExport} -rev_rule=\"Latest Working\" -export_bom=yes -transfermode=ConfiguredDataExportDefault -xml_file=\"{allXmlsPath}{itemToExport}.xml\"";
                    exportCommands.Add(command);
                }
            }

            Utilidades.EscribirEnLog($"Se encontraron {uniqueProductsToExport.Count} ítems únicos (incluyendo el padre) para exportar.");


            Utilidades.EscribirEnLog($"Se generaron {exportCommands.Count} comandos de exportación para distribuir en {NUM_BATCH_FILES} archivos.");

            List<Process> runningProcesses = new List<Process>();

            for (int i = 0; i < NUM_BATCH_FILES; i++)
            {
                string batFileName = $"{baseBatName}{i + 1}.bat";
                string rutaBat = Path.Combine(directoryBat, batFileName);
                string contenidoBat = tcConfigContent;

                // Distribución equitativa: Asignar al bat 'i' los comandos 'i', 'i+N', 'i+2N', etc.
                for (int j = i; j < exportCommands.Count; j += NUM_BATCH_FILES)
                {
                    contenidoBat += exportCommands[j];
                }

                contenidoBat += "\nexit";

                // Escribir el contenido en el nuevo archivo .bat
                File.WriteAllText(rutaBat, contenidoBat);
                Utilidades.EscribirEnLog($"Archivo {batFileName} creado. Comandos asignados: {exportCommands.Count / NUM_BATCH_FILES + (i < exportCommands.Count % NUM_BATCH_FILES ? 1 : 0)}");


                // 3. EJECUCIÓN PARALELA DE CADA BATCH FILE
                var ProcesosStarInfo = new ProcessStartInfo();
                ProcesosStarInfo.FileName = rutaBat;
                ProcesosStarInfo.WorkingDirectory = directoryBat;
                ProcesosStarInfo.UseShellExecute = false;
                ProcesosStarInfo.RedirectStandardOutput = true;

                var proceso = new Process();
                proceso.StartInfo = ProcesosStarInfo;
                proceso.EnableRaisingEvents = true;

                // Manejo de la salida estándar (cada proceso tendrá su log)
                proceso.OutputDataReceived += (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        Utilidades.EscribirEnLog($"[{batFileName}] Salida del proceso: {e.Data}");
                    }
                };

                proceso.Start();
                proceso.BeginOutputReadLine();
                runningProcesses.Add(proceso);
            }

            // 4. ESPERAR A QUE TODOS LOS PROCESOS TERMINEN
            foreach (var proceso in runningProcesses)
            {
                proceso.WaitForExit();
                Utilidades.EscribirEnLog($"Proceso {proceso.StartInfo.FileName} ha terminado con código {proceso.ExitCode}.");
            }

            Utilidades.EscribirEnLog("Todos los procesos de exportación paralelos han finalizado.");


        }


        // --------------------------------------------- Metodos no utilizados por el Main() (Carga en Base de datos de BOP) ---------------------------------------------


        static bool ParseNode(XmlNode node, Dictionary<string, List<DataRow>> groupedDataRows, string parentNodeName = "")
        {
            // Crear una lista de nombres de nodos a ignorar
            var listaIgnorados = new List<string> { "ApplicationRef", "AssociatedDataSet", "AttributeContext", "DataSet",
                                                     "ExternalFile", "Folder", "InstanceGraph", "ProductDef", "ProductInstance",
                                                     "ProductRevisionView", "RevisionRule", "Site", "Transform", "View" };
            try
            {
                if (node.NodeType == XmlNodeType.Element && !listaIgnorados.Contains(node.Name))
                {
                    string nodeName = node.Name;

                    DataRow dataRow = new DataRow();
                    dataRow.NombreNodo = nodeName;


                    dataRow.Atributos = new List<string>();

                    foreach (XmlAttribute attribute in node.Attributes)
                    {
                        dataRow.Atributos.Add(attribute.Name);
                    }

                    dataRow.XmlNode = node;
                    string tableName = GetTableName(nodeName, dataRow.Atributos, parentNodeName);

                    if (!groupedDataRows.ContainsKey(tableName))
                    {
                        groupedDataRows[tableName] = new List<DataRow>();
                    }
                    groupedDataRows[tableName].Add(dataRow);

                    foreach (XmlNode childNode in node.ChildNodes)
                    {
                        ParseNode(childNode, groupedDataRows, nodeName); //recursividad
                    }
                    return true;
                }
                return false;
            }
            catch (Exception ea)
            {
                Utilidades.EscribirEnLog("Excepcion controlada en el metodo ParseNode: " + ea.Message);
                return false;
            }

        }

        static string GetTableName(string nodeName, List<string> attributes, string parentNodeName)
        {
            string tableName = nodeName;
            if (!attributes.Contains("id") && tableName != "PLMXML")
            {

                tableName = $"{nodeName}_{parentNodeName}";
            }
            return tableName;
        }




        
    }


}