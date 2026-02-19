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
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;

namespace crucia
{
    public static class Utilidades
    {
        private static readonly object _lock = new object();

        // Base fija (misma que hoy)
        private static readonly string _logDir = @"E:\DescarConector";

        // Cache diario
        private static string _rutaLogActual = null;
        private static string _fechaCache = null; // "dd_MM_yy"
        public static string SanitizarNombreParaRuta(string input, int maxLen = 80)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "SIN_NOMBRE";

            // 1) Quitar diacríticos (Ñ -> N, á -> a, etc.)
            string normalized = input.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(normalized.Length);

            foreach (char c in normalized)
            {
                var cat = CharUnicodeInfo.GetUnicodeCategory(c);
                if (cat != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }

            string sinDiacriticos = sb.ToString().Normalize(NormalizationForm.FormC);

            // 2) Por las dudas, normalizar ñ/Ñ explícitamente también
            sinDiacriticos = sinDiacriticos.Replace('ñ', 'n').Replace('Ñ', 'N');

            // 3) Reemplazar caracteres inválidos para nombre de carpeta/archivo
            string safe = sinDiacriticos;
            foreach (char invalid in Path.GetInvalidFileNameChars())
                safe = safe.Replace(invalid, '_');

            // 4) Compactar espacios y limpiar extremos raros
            safe = Regex.Replace(safe, @"\s+", "_").Trim(' ', '.', '_');

            if (string.IsNullOrWhiteSpace(safe))
                safe = "SIN_NOMBRE";

            if (safe.Length > maxLen)
                safe = safe.Substring(0, maxLen);

            return safe;
        }

        /// <summary>
        /// Sanitiza SOLO el último segmento de la ruta (la carpeta final).
        /// Ej: E:\...\CAÑO1\  => E:\...\CANO1\
        /// </summary>
        public static string SanitizarUltimaCarpetaDeRuta(string ruta)
        {
            if (string.IsNullOrWhiteSpace(ruta))
                return ruta;

            char sep = Path.DirectorySeparatorChar;

            string original = ruta;
            string trimmed = ruta.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            string? parent = Path.GetDirectoryName(trimmed);
            string leaf = Path.GetFileName(trimmed);

            // Si no hay "parent" (casos raros), no tocamos
            if (string.IsNullOrWhiteSpace(parent) || string.IsNullOrWhiteSpace(leaf))
                return original.EndsWith(sep.ToString()) ? original : (original + sep);

            string leafSan = SanitizarNombreParaRuta(leaf);

            string nueva = Path.Combine(parent, leafSan) + sep;

            if (!string.Equals(original, nueva, StringComparison.OrdinalIgnoreCase))
                EscribirEnLog($"Sanitización de ruta de salida: '{original}' => '{nueva}'");

            return nueva;
        }

        private static string ObtenerRutaLogDelDia()
        {
            string hoy = DateTime.Now.ToString("dd_MM_yy");

            // Si cambió el día, actualizamos ruta
            if (!string.Equals(_fechaCache, hoy, StringComparison.Ordinal))
            {
                _fechaCache = hoy;
                _rutaLogActual = Path.Combine(_logDir, $"Intermedio_log_{hoy}.txt");
            }

            return _rutaLogActual ?? Path.Combine(_logDir, $"Intermedio_log_{hoy}.txt");
        }

        public static void EscribirEnLog(string mensaje)
        {
            try
            {
                Directory.CreateDirectory(_logDir);

                string rutaLog = ObtenerRutaLogDelDia();
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {mensaje}{Environment.NewLine}";

                lock (_lock)
                {
                    File.AppendAllText(rutaLog, line);
                }
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
                if (!Directory.Exists(rutaCarpeta))
                {
                    Directory.CreateDirectory(rutaCarpeta);
                    Utilidades.EscribirEnLog($"La carpeta de salida {rutaCarpeta} no existía y fue creada.");
                }

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

            // Sanitiza la carpeta final (por ej. "CAÑO1" -> "CANO1")
            allXmlsPath = Utilidades.SanitizarUltimaCarpetaDeRuta(allXmlsPath);

            // Aseguramos separador al final (por si vino sin \)
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

            // ====================================================================================================================
            // === CÓDIGO CORREGIDO: Recolectar Items únicos (incluyendo el raíz) para garantizar la exportación del padre.
            // ====================================================================================================================

            // 1. Obtener todos los elementos Product (incluyendo el raíz de la MBOM)
            // 1) Construir diccionario productId -> subType
            Dictionary<string, string> productos = ConstruirDiccionarioProductos(xdoc);
            Utilidades.EscribirEnLog($"Se encontraron {productos.Count} Product únicos (productId) en el XML.");

            // 2) Generar los comandos de exportación iterando el diccionario y aplicando sanitización actual
            foreach (var kv in productos)
            {
                string productId = kv.Key;
                string subType = kv.Value;

                string itemToExport = SanitizarItemAExportar(productId, subType);

                Utilidades.EscribirEnLog($"Producto leído: productId={productId} subType={subType} => itemToExport={itemToExport}");
                Console.WriteLine($"Producto: {productId} - SubType: {subType} - Export: {itemToExport}");

                if (!string.IsNullOrEmpty(itemToExport))
                {
                    string command =
                        $"\r\nplmxml_export -u=lacuna -p=lacuna -g=Proceso -item={itemToExport} " +
                        "-rev_rule=\"Latest Working\" -export_bom=yes -transfermode=ConfiguredDataExportDefault " +
                        $" -xml_file=\"{allXmlsPath}{itemToExport}.xml\"";

                    exportCommands.Add(command);
                }
            }

            Utilidades.EscribirEnLog($"Se generaron {exportCommands.Count} comandos de exportación a partir del diccionario.");


            //Utilidades.EscribirEnLog($"Se encontraron {uniqueProductsToExport.Count} ítems únicos (incluyendo el padre) para exportar.");

            // ====================================================================================================================
            // === FIN DEL CÓDIGO CORREGIDO
            // ====================================================================================================================


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

        static Dictionary<string, string> ConstruirDiccionarioProductos(XDocument xdoc)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var productElements = xdoc
                .Descendants()
                .Where(e => e.Name.LocalName == "Product");

            foreach (var productElement in productElements)
            {
                string? productId = productElement.Attribute("productId")?.Value?.Trim();
                string? subType = productElement.Attribute("subType")?.Value?.Trim();

                if (string.IsNullOrEmpty(productId))
                    continue;

                if (dict.TryGetValue(productId, out var subTypeExistente))
                {
                    if (!string.Equals(subTypeExistente, subType, StringComparison.OrdinalIgnoreCase))
                    {
                        Utilidades.EscribirEnLog(
                            $"WARNING: productId repetido con subType distinto. productId={productId}, existente={subTypeExistente}, nuevo={subType}"
                        );
                    }
                    continue;
                }

                dict[productId] = subType ?? string.Empty;
            }

            return dict;
        }

        static string SanitizarItemAExportar(string productId, string subType)
        {
            string itemToExport = "";

            if (string.Equals(subType, "Agm4_Pieza", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(productId))
                    return "";

                itemToExport = "P-" + productId;
            }
            else if (string.Equals(subType, "Agm4_SubCon", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(productId))
                    return "";

                if (productId.StartsWith("E", StringComparison.OrdinalIgnoreCase))
                    itemToExport = "P-" + productId.Substring(1);
                else
                    itemToExport = "P-" + productId;
            }
            else if (string.Equals(subType, "Agm4_sub_mBOM_E", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(subType, "Agm4_sub_mBOM_S", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(subType, "Agm4_conj_mBOM_F", StringComparison.OrdinalIgnoreCase)
                     )
            {
                if (string.IsNullOrWhiteSpace(productId))
                    return "";

                string id = productId;

                if (id.StartsWith("M-", StringComparison.OrdinalIgnoreCase))
                    id = id.Substring(2);
                else if (id.StartsWith("M", StringComparison.OrdinalIgnoreCase))
                    id = id.Substring(1);

                itemToExport = "P-" + id;
            }

            return itemToExport;
        }

    }
}