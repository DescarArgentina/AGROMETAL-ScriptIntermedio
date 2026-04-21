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
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;

namespace crucia
{
    public static class Utilidades
    {
        private static readonly object _lock = new object();

        // Base fija (misma que hoy)
        private static readonly string _logDir = @"E:\\DescarConector";

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

            // 5) Recortar largo
            if (safe.Length > maxLen)
                safe = safe.Substring(0, maxLen);

            if (string.IsNullOrWhiteSpace(safe))
                safe = "SIN_NOMBRE";

            return safe;
        }

        public static string SanitizarUltimaCarpetaDeRuta(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath))
                return fullPath;

            // Quitar separador final si viene
            string trimmed = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            string parent = Path.GetDirectoryName(trimmed);
            string last = Path.GetFileName(trimmed);

            // Sanitiza SOLO la última carpeta
            string lastSafe = SanitizarNombreParaRuta(last);

            if (string.IsNullOrEmpty(parent))
                return lastSafe;

            return Path.Combine(parent, lastSafe);
        }

        private static string ObtenerRutaLogDelDia()
        {
            string fechaHoy = DateTime.Now.ToString("dd_MM_yy");
            if (_rutaLogActual == null || _fechaCache != fechaHoy)
            {
                Directory.CreateDirectory(_logDir);
                _fechaCache = fechaHoy;
                _rutaLogActual = Path.Combine(_logDir, $"Intermedio_log_{fechaHoy}.txt");
            }
            return _rutaLogActual;
        }

        public static void EscribirEnLog(string mensaje)
        {
            lock (_lock)
            {
                try
                {
                    string rutaLog = ObtenerRutaLogDelDia();
                    File.AppendAllText(rutaLog, $"{DateTime.Now:dd-MM-yy HH:mm:ss} | {mensaje}{Environment.NewLine}");
                }
                catch
                {
                    // No frenamos el proceso si el log falla
                }
            }
        }

        public static void LimpiarCarpeta(string rutaCarpeta)
        {
            try
            {
                if (Directory.Exists(rutaCarpeta))
                {
                    DirectoryInfo di = new DirectoryInfo(rutaCarpeta);
                    foreach (FileInfo file in di.GetFiles())
                    {
                        file.Delete();
                    }
                    foreach (DirectoryInfo dir in di.GetDirectories())
                    {
                        dir.Delete(true);
                    }
                }
                else
                {
                    Directory.CreateDirectory(rutaCarpeta);
                }
            }
            catch (Exception ex)
            {
                Utilidades.EscribirEnLog($"Error al limpiar o crear la carpeta de salida {rutaCarpeta}: {ex.Message}");
            }
        }
    }

    class Program
    {

        // Límites por consola (proceso .bat + hijos)
        private static readonly TimeSpan TIMEOUT_MAX_EJECUCION = TimeSpan.FromHours(3);
        private static readonly TimeSpan TIMEOUT_SIN_SALIDA = TimeSpan.FromMinutes(90); // 1.5 horas
        private static readonly long LIMITE_MEMORIA_BYTES = CalcularLimiteMemoriaBytes();

        private static long CalcularLimiteMemoriaBytes()
        {
            const long eightGb = 8L * 1024 * 1024 * 1024;

            // En x86, el espacio de direcciones es de 4GB máximo, no tiene sentido poner 8GB.
            if (!Environment.Is64BitProcess)
                return uint.MaxValue; // 4GB - 1

            return eightGb;
        }
        private static readonly TimeSpan INTERVALO_MONITOREO = TimeSpan.FromSeconds(5);

        private sealed class ActivityTracker
        {
            public long StartTicksUtc;
            public long LastOutputTicksUtc;
        }

        private sealed class ProcessContext
        {
            public string BatFileName;
            public Process Proceso;
            public JobObject Job;
            public ActivityTracker Activity;

            public bool CancelSolicitado;
            public string MotivoCancelacion;
        }

        private static void MonitorearYAplicarLimites(List<ProcessContext> contexts)
        {
            if (contexts == null || contexts.Count == 0)
                return;

            int activos = contexts.Count;
            Utilidades.EscribirEnLog($"Monitoreo iniciado. Procesos activos: {activos}. Límites: Total={TIMEOUT_MAX_EJECUCION}, SinSalida={TIMEOUT_SIN_SALIDA}, Mem={FormatearBytes(LIMITE_MEMORIA_BYTES)}");

            while (activos > 0)
            {
                for (int i = contexts.Count - 1; i >= 0; i--)
                {
                    var ctx = contexts[i];

                    bool exited = false;
                    try { exited = ctx.Proceso.HasExited; } catch { exited = true; }

                    if (exited)
                    {
                        int exitCode = 0;
                        try { exitCode = ctx.Proceso.ExitCode; } catch { }

                        long peak = 0;
                        try { peak = ctx.Job.GetPeakJobMemoryUsed(); } catch { }

                        Utilidades.EscribirEnLog($"[{ctx.BatFileName}] Finalizó. ExitCode={exitCode}. PeakJobMem={FormatearBytes(peak)}");

                        try { ctx.Job.Dispose(); } catch { }
                        contexts.RemoveAt(i);
                        activos--;
                        continue;
                    }

                    long nowTicks = DateTime.UtcNow.Ticks;

                    // 1) Límite total (3 horas)
                    long runTicks = nowTicks - Interlocked.Read(ref ctx.Activity.StartTicksUtc);
                    if (!ctx.CancelSolicitado && runTicks > TIMEOUT_MAX_EJECUCION.Ticks)
                    {
                        ctx.CancelSolicitado = true;
                        ctx.MotivoCancelacion = "TIMEOUT_MAX_EJECUCUCION";
                        Utilidades.EscribirEnLog($"[{ctx.BatFileName}] ERROR: Tiempo máximo excedido (> {TIMEOUT_MAX_EJECUCION}). Se cancela.");
                        TryTerminateJob(ctx);
                        continue;
                    }

                    // 2) Límite por “sin salida” (1.5 horas)
                    long idleTicks = nowTicks - Interlocked.Read(ref ctx.Activity.LastOutputTicksUtc);
                    if (!ctx.CancelSolicitado && idleTicks > TIMEOUT_SIN_SALIDA.Ticks)
                    {
                        ctx.CancelSolicitado = true;
                        ctx.MotivoCancelacion = "TIMEOUT_SIN_SALIDA";
                        Utilidades.EscribirEnLog($"[{ctx.BatFileName}] ERROR: Sin salida por más de {TIMEOUT_SIN_SALIDA}. Se cancela.");
                        TryTerminateJob(ctx);
                        continue;
                    }

                    // 3) Límite de memoria (4 GB) por job (árbol completo)
                    if (!ctx.CancelSolicitado)
                    {
                        long mem = 0;
                        try { mem = ctx.Job.SumPrivateBytes(); } catch { }

                        if (mem >= LIMITE_MEMORIA_BYTES)
                        {
                            ctx.CancelSolicitado = true;
                            ctx.MotivoCancelacion = "MEMORIA_EXCEDIDA";
                            Utilidades.EscribirEnLog($"[{ctx.BatFileName}] ERROR: Memoria excedida. Actual={FormatearBytes(mem)} Límite={FormatearBytes(LIMITE_MEMORIA_BYTES)}. Se cancela.");
                            TryTerminateJob(ctx);
                            continue;
                        }
                    }
                }

                Thread.Sleep(INTERVALO_MONITOREO);
                activos = contexts.Count;
            }

            Utilidades.EscribirEnLog("Monitoreo finalizado. No quedan procesos activos.");
        }

        private static void TryTerminateJob(ProcessContext ctx)
        {
            try
            {
                ctx.Job.Terminate(1);
                Utilidades.EscribirEnLog($"[{ctx.BatFileName}] Cancelación enviada. Motivo={ctx.MotivoCancelacion}.");
            }
            catch (Exception ex)
            {
                Utilidades.EscribirEnLog($"[{ctx.BatFileName}] ERROR al cancelar. Motivo={ctx.MotivoCancelacion}. Ex={ex.Message}");
            }
        }

        private static string FormatearBytes(long bytes)
        {
            if (bytes <= 0) return "0 B";
            double b = bytes;
            string[] u = new[] { "B", "KB", "MB", "GB", "TB" };
            int i = 0;
            while (b >= 1024 && i < u.Length - 1)
            {
                b /= 1024;
                i++;
            }
            return $"{b:0.##} {u[i]}";
        }

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

            // ====================================================================================================================
            // === FIN DEL CÓDIGO CORREGIDO
            // ====================================================================================================================

            // 2. DISTRIBUCIÓN EQUITATIVA Y GENERACIÓN DE MÚLTIPLES BATCH FILES
            Utilidades.EscribirEnLog($"Se generaron {exportCommands.Count} comandos de exportación para distribuir en {NUM_BATCH_FILES} archivos.");

            List<ProcessContext> runningProcesses = new List<ProcessContext>();

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

                // 3. EJECUCIÓN PARALELA DE CADA BATCH FILE (con límites de tiempo y memoria)
                var ProcesosStarInfo = new ProcessStartInfo();
                ProcesosStarInfo.FileName = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
                ProcesosStarInfo.Arguments = $"/d /s /c \"\"{rutaBat}\"\"";
                ProcesosStarInfo.WorkingDirectory = directoryBat;
                ProcesosStarInfo.UseShellExecute = false;
                ProcesosStarInfo.RedirectStandardOutput = true;
                ProcesosStarInfo.RedirectStandardError = true;
                ProcesosStarInfo.CreateNoWindow = true;

                var proceso = new Process();
                proceso.StartInfo = ProcesosStarInfo;
                proceso.EnableRaisingEvents = true;

                var activity = new ActivityTracker();
                activity.StartTicksUtc = DateTime.UtcNow.Ticks;
                activity.LastOutputTicksUtc = activity.StartTicksUtc;

                // Manejo de la salida (cada proceso tendrá su log)
                proceso.OutputDataReceived += (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        Interlocked.Exchange(ref activity.LastOutputTicksUtc, DateTime.UtcNow.Ticks);
                        Utilidades.EscribirEnLog($"[{batFileName}] Salida del proceso: {e.Data}");
                    }
                };

                proceso.ErrorDataReceived += (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        Interlocked.Exchange(ref activity.LastOutputTicksUtc, DateTime.UtcNow.Ticks);
                        Utilidades.EscribirEnLog($"[{batFileName}] ERROR(stderr): {e.Data}");
                    }
                };

                try
                {
                    proceso.Start();
                    proceso.BeginOutputReadLine();
                    proceso.BeginErrorReadLine();

                    var job = new JobObject($"Intermedio_{batFileName}_{proceso.Id}");
                    job.SetJobMemoryLimit(LIMITE_MEMORIA_BYTES);
                    job.Assign(proceso);

                    runningProcesses.Add(new ProcessContext
                    {
                        BatFileName = batFileName,
                        Proceso = proceso,
                        Job = job,
                        Activity = activity
                    });
                }
                catch (Win32Exception ex)
                {
                    Utilidades.EscribirEnLog($"[{batFileName}] ERROR al iniciar/configurar el proceso: {ex.Message} | NativeErrorCode={ex.NativeErrorCode}");
                    try { proceso.Kill(true); } catch { }
                }
                catch (Exception ex)
                {
                    Utilidades.EscribirEnLog($"[{batFileName}] ERROR al iniciar/configurar el proceso: {ex.Message}");
                    try { proceso.Kill(true); } catch { }
                }
            }

            // 4. MONITOREAR HASTA QUE TODOS LOS PROCESOS TERMINEN (aplicando límites)
            MonitorearYAplicarLimites(runningProcesses);

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
                        // si aparecen distintos subType, no pisamos; solo registramos
                    }
                }
                else
                {
                    dict[productId] = subType ?? "";
                }
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

    internal sealed class JobObject : IDisposable
    {
        private IntPtr _hJob;
        private bool _disposed;

        private const int ERROR_MORE_DATA = 234;

        private const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000;
        private const uint JOB_OBJECT_LIMIT_JOB_MEMORY = 0x00000200;

        public JobObject(string name)
        {
            _hJob = CreateJobObject(IntPtr.Zero, name);
            if (_hJob == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error());

            SetLimits(killOnClose: true, jobMemoryLimitBytes: null);
        }

        public void SetJobMemoryLimit(long bytes)
        {
            if (bytes <= 0) throw new ArgumentOutOfRangeException(nameof(bytes));
            SetLimits(killOnClose: true, jobMemoryLimitBytes: bytes);
        }

        public void Assign(Process process)
        {
            if (process == null) throw new ArgumentNullException(nameof(process));
            if (!AssignProcessToJobObject(_hJob, process.Handle))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        public void Terminate(uint exitCode = 1)
        {
            if (_hJob != IntPtr.Zero)
                TerminateJobObject(_hJob, exitCode);
        }

        // Suma de Private Bytes de todos los procesos del job (árbol completo)
        public long SumPrivateBytes()
        {
            long sum = 0;
            foreach (var pid in GetProcessIds())
            {
                try
                {
                    using var p = Process.GetProcessById((int)pid);
                    sum += p.PrivateMemorySize64;
                }
                catch { }
            }
            return sum;
        }

        public long GetPeakJobMemoryUsed()
        {
            var info = QueryExtendedLimitInfo();
            return (long)info.PeakJobMemoryUsed.ToUInt64();
        }

        private void SetLimits(bool killOnClose, long? jobMemoryLimitBytes)
        {
            var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
            uint flags = 0;

            if (killOnClose) flags |= JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;

            if (jobMemoryLimitBytes.HasValue)
            {
                flags |= JOB_OBJECT_LIMIT_JOB_MEMORY;
                if (IntPtr.Size == 4)
                    info.JobMemoryLimit = new UIntPtr((uint)jobMemoryLimitBytes.Value);
                else
                    info.JobMemoryLimit = new UIntPtr((ulong)jobMemoryLimitBytes.Value);
            }

            info.BasicLimitInformation.LimitFlags = flags;

            int length = Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
            IntPtr ptr = Marshal.AllocHGlobal(length);
            try
            {
                Marshal.StructureToPtr(info, ptr, false);
                if (!SetInformationJobObject(_hJob, JobObjectInfoType.ExtendedLimitInformation, ptr, (uint)length))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }

        private JOBOBJECT_EXTENDED_LIMIT_INFORMATION QueryExtendedLimitInfo()
        {
            int length = Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
            IntPtr ptr = Marshal.AllocHGlobal(length);
            try
            {
                if (!QueryInformationJobObject(_hJob, JobObjectInfoType.ExtendedLimitInformation, ptr, (uint)length, out _))
                    throw new Win32Exception(Marshal.GetLastWin32Error());

                return Marshal.PtrToStructure<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>(ptr);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }

        private ulong[] GetProcessIds()
        {
            int capacity = 64;
            while (true)
            {
                int length = 8 + (IntPtr.Size * capacity);
                IntPtr buffer = Marshal.AllocHGlobal(length);
                try
                {
                    bool ok = QueryInformationJobObject(_hJob, JobObjectInfoType.BasicProcessIdList, buffer, (uint)length, out _);
                    if (!ok)
                    {
                        int err = Marshal.GetLastWin32Error();
                        if (err == ERROR_MORE_DATA)
                        {
                            capacity *= 2;
                            continue;
                        }
                        throw new Win32Exception(err);
                    }

                    uint count = (uint)Marshal.ReadInt32(buffer, 4); // NumberOfProcessIdsInList
                    var pids = new ulong[count];

                    int offset = 8;
                    for (int i = 0; i < count; i++)
                    {
                        if (IntPtr.Size == 8)
                            pids[i] = (ulong)Marshal.ReadInt64(buffer, offset + (i * 8));
                        else
                            pids[i] = (uint)Marshal.ReadInt32(buffer, offset + (i * 4));
                    }
                    return pids;
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_hJob != IntPtr.Zero)
            {
                CloseHandle(_hJob);
                _hJob = IntPtr.Zero;
            }
        }

        private enum JobObjectInfoType
        {
            BasicProcessIdList = 3,
            ExtendedLimitInformation = 9
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IO_COUNTERS
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JOBOBJECT_BASIC_LIMIT_INFORMATION
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;      // <-- FIX (antes long)
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
            public IO_COUNTERS IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string? lpName);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(IntPtr hJob, JobObjectInfoType infoType, IntPtr lpJobObjectInfo, uint cbJobObjectInfoLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool QueryInformationJobObject(IntPtr hJob, JobObjectInfoType infoType, IntPtr lpJobObjectInfo, uint cbJobObjectInfoLength, out uint lpReturnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateJobObject(IntPtr hJob, uint uExitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);
    }
}