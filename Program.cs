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
            // La ruta del log sigue hardcodeada, típicamente se deja así o se saca de un archivo de configuración.
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


            //// Iterar sobre todos los elementos Product y generar comandos
            foreach (XElement productElement in xdoc.Descendants("{http://www.plmxml.org/Schemas/PLMXMLSchema}Product"))
            {
                string productId = productElement.Attribute("productId")?.Value;
                string subtype = productElement.Attribute("subType")?.Value;
                string itemToExport = "";

                if (string.IsNullOrEmpty(productId)) continue;

                if (subtype == "Agm4_Pieza")
                {
                    itemToExport = "P-" + productId;
                }
                else if (subtype == "Agm4_SubCon")
                {
                    if (productId.Contains('E'))
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

            // 2. DISTRIBUCIÓN EQUITATIVA Y GENERACIÓN DE MÚLTIPLES BATCH FILES
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
        static void BorrarTabla(SqlConnection connection, Dictionary<string, List<DataRow>> groupedDataRows)
        {
            foreach (var group in groupedDataRows)
            {
                try
                {
                    string tableName = group.Key;
                    string deleteTableQuery = $"IF OBJECT_ID('[{tableName}]', 'U') IS NOT NULL DROP TABLE [{tableName}]";
                    using (SqlCommand command = new SqlCommand(deleteTableQuery, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ea)
                {
                    Utilidades.EscribirEnLog($"Error al intentar borrar la tabla para su sobreescritura - Error: {ea.Message}");
                }
            }
        }

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

        static void CreateTable(SqlConnection connection, Dictionary<string, List<DataRow>> groupedDataRows)
        {
            foreach (var group in groupedDataRows)
            {
                string tableName = group.Key;
                if (tableName == "PLMXML")
                {
                    continue;
                }
                string createTableQuery = $"IF OBJECT_ID('[{tableName}]', 'U') IS NOT NULL DROP TABLE [{tableName}]; CREATE TABLE [{tableName}] (id INT IDENTITY(1,1) PRIMARY KEY, contenido NVARCHAR(MAX)";
                List<string> additionalAttributes = new List<string>();
                bool hasIdAttribute = false;

                foreach (DataRow dataRow in group.Value)
                {
                    foreach (string attribute in dataRow.Atributos)
                    {
                        if (!additionalAttributes.Contains(attribute) && attribute != "id")
                        {
                            additionalAttributes.Add(attribute);
                        }
                        if (attribute == "id")
                        {
                            hasIdAttribute = true;
                        }
                    }
                }
                if (hasIdAttribute)
                {
                    createTableQuery += ", id_Table INT";
                }
                else
                {
                    createTableQuery += ", id_Father INT";
                }
                foreach (string columnName in additionalAttributes)
                {
                    if (columnName != "id")
                    {
                        createTableQuery += $", [{columnName}] NVARCHAR(MAX)";
                    }
                }
                createTableQuery += ");";
                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        static void AlterTable(SqlConnection connection, string tableName, string columnName, string columnType)
        {
            try
            {
                string alterTableQuery = $"IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName}' AND COLUMN_NAME = '{columnName}') " +
                $"ALTER TABLE [{tableName}] ADD [{columnName}] {columnType};";

                using (SqlCommand command = new SqlCommand(alterTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ea)
            {
                Utilidades.EscribirEnLog($"Excepcion controlada en el metodo AlterTable: {ea.Message}");
                throw;
            }

        }

        static void InsertData(SqlConnection connection, Dictionary<string, List<DataRow>> groupedDataRows, string xml, int contadorXmls)
        {
            try
            {
                foreach (var group in groupedDataRows)
                {
                    string tableName = group.Key;

                    foreach (DataRow dataRow in group.Value)
                    {
                        if (dataRow.NombreNodo == "PLMXML")
                            continue;
                        string insertQuery = $"INSERT INTO [{tableName}] (";
                        List<string> columnNames = new List<string>();
                        List<string> parameterNames = new List<string>();
                        List<SqlParameter> parameters = new List<SqlParameter>();
                        bool hasIdAttribute = false;


                        foreach (string columnName in dataRow.Atributos)
                        {

                            if (columnName == "id" || columnName == "instancedRef" || columnName == "masterRef" || columnName == "parentRef" || columnName == "instanceRefs")
                            {
                                string attributeValue1 = dataRow.XmlNode.Attributes[columnName]?.Value;

                                if (columnName == "id" && !string.IsNullOrEmpty(attributeValue1) && attributeValue1.Length > 2)
                                {
                                    hasIdAttribute = true;
                                    columnNames.Add("[id_Table]");
                                    parameterNames.Add("@id");
                                    attributeValue1 = attributeValue1.Substring(2);
                                    parameters.Add(new SqlParameter("@id", attributeValue1));
                                }
                                if (columnName == "instancedRef" && !string.IsNullOrEmpty(attributeValue1) && attributeValue1.Length > 2)
                                {
                                    columnNames.Add("[instancedRef]");
                                    parameterNames.Add("@instancedRef");
                                    attributeValue1 = attributeValue1.Substring(3);
                                    parameters.Add(new SqlParameter("@instancedRef", attributeValue1));
                                }
                                if (columnName == "masterRef" && !string.IsNullOrEmpty(attributeValue1) && attributeValue1.Length > 2)
                                {
                                    columnNames.Add("[masterRef]");
                                    parameterNames.Add("@masterRef");
                                    attributeValue1 = attributeValue1.Substring(3);
                                    parameters.Add(new SqlParameter("@masterRef", attributeValue1));
                                }
                                if (columnName == "parentRef" && !string.IsNullOrEmpty(attributeValue1) && attributeValue1.Length > 2)
                                {
                                    columnNames.Add("[parentRef]");
                                    parameterNames.Add("@parentRef");
                                    attributeValue1 = attributeValue1.Substring(3);
                                    parameters.Add(new SqlParameter("@parentRef", attributeValue1));
                                }
                                if (columnName == "instanceRefs" && !string.IsNullOrEmpty(attributeValue1) && attributeValue1.Length > 2)
                                {
                                    columnNames.Add("[instanceRefs]");
                                    parameterNames.Add("@instanceRefs");
                                    attributeValue1 = attributeValue1.Substring(3);
                                    parameters.Add(new SqlParameter("@instanceRefs", attributeValue1));
                                }
                                continue;
                            }
                            AlterTable(connection, tableName, columnName, "NVARCHAR(MAX)");
                            columnNames.Add($"[{columnName}]");
                            parameterNames.Add($"@{columnName}");
                            string attributeValue = dataRow.XmlNode.Attributes[columnName]?.Value;
                            attributeValue = attributeValue.Replace("'", "''");
                            parameters.Add(new SqlParameter($"@{columnName}", attributeValue));

                        }
                        columnNames.Add("[contenido]");
                        parameterNames.Add("@contenido");
                        parameters.Add(new SqlParameter("@contenido", dataRow.XmlNode.InnerText));

                        if (!hasIdAttribute)
                        {
                            columnNames.Add("[id_Father]");
                            parameterNames.Add("@idFather");
                            XmlNode parentNode = dataRow.XmlNode.ParentNode;
                            string parentAttributeValue = parentNode?.Attributes["id"]?.Value;
                            string parentAttributeId = parentAttributeValue?.Substring(2) ?? "0";
                            parameters.Add(new SqlParameter("@idFather", parentAttributeId));
                        }
                        columnNames.Add("[idXml]");
                        parameterNames.Add("@idXml");
                        parameters.Add(new SqlParameter("@idXml", contadorXmls));

                        insertQuery += string.Join(", ", columnNames) + ") VALUES (";
                        insertQuery += string.Join(", ", parameterNames) + ");";

                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            command.Parameters.AddRange(parameters.ToArray());
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ea)
            {
                Utilidades.EscribirEnLog($"Excepcion controlada en el metodo InsertData {ea.Message}");
            }

        }
    }
}