﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using Newtonsoft.Json;

namespace Viramate {
    static partial class Program {
        [DllImport(
            "kernel32.dll", EntryPoint = "GetStdHandle", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern IntPtr GetStdHandle (int nStdHandle);

        [DllImport(
            "kernel32.dll", EntryPoint = "AllocConsole", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int AllocConsole ();

        [DllImport(
            "kernel32.dll", EntryPoint = "AttachConsole", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int AttachConsole (int processId);

        [DllImport(
            "kernel32.dll", EntryPoint = "SetConsoleTitle", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int SetConsoleTitle (string title);

        [DllImport(
            "kernel32.dll", EntryPoint = "TerminateProcess", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int TerminateProcess (int processId, int exitCode);

        private const int ATTACH_PARENT_PROCESS = -1;
        private const int STD_INPUT_HANDLE = -10;
        private const int STD_OUTPUT_HANDLE = -11;
        private const int STD_ERROR_HANDLE = -12;
        private const int MY_CODE_PAGE = 437;

        public static Assembly MyAssembly;
        static bool IsRunningInsideCmd = false;

        public const string ExtensionSourceUrl = "https://viramate.luminance.org/ext.zip";
        public const string InstallerSourceUrl = "https://viramate.luminance.org/installer.zip";

        static Program () {
            MyAssembly = Assembly.GetExecutingAssembly();
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
        }

        private static Assembly CurrentDomain_AssemblyResolve (object sender, ResolveEventArgs args) {
            const int maxLength = 1024 * 1024 * 2;
            var dllName = new AssemblyName(args.Name).Name + ".dll.gz";
            var resourceName = MyAssembly.GetManifestResourceNames().FirstOrDefault(s => s.EndsWith(dllName));
            if (resourceName != null) {
                using (var stream = MyAssembly.GetManifestResourceStream(resourceName))
                using (var gzstream = new GZipStream(stream, CompressionMode.Decompress, true)) {
                    var asmbuf = new byte[maxLength];
                    var bytesRead = gzstream.Read(asmbuf, 0, maxLength);
                    return Assembly.Load(asmbuf);
                }
            }
            Console.Error.WriteLine($"No resource named {dllName}");
            return null;
        }

        static void Main (string[] args) {
            try {
                if (
                    args.Contains("/?") ||
                    args.Contains("-?")
                ) {
                    InitConsole();
                    PrintHelp();
                } else if (
                    (args.Length <= 1) || 
                    !args.Any(a => a.StartsWith("chrome-extension://"))
                ) {
                    InitConsole();
                    InstallExtension().Wait();
                } else
                    MessagingHostMainLoop();
            } catch (Exception exc) {
                Console.Error.WriteLine("Uncaught: {0}", exc);
                Environment.ExitCode = 1;

                // HACK: Assume the crash might be an installer bug, so try to install an update.
                AutoUpdateInstaller().Wait();
            }
        }

        static void InitConsole () {
            if (Debugger.IsAttached) {
                // No work necessary I think?
            } else {
                if (AttachConsole(ATTACH_PARENT_PROCESS) == 0) {
                    AllocConsole();
                    SetConsoleTitle("Viramate Installer");
                } else {
                    IsRunningInsideCmd = true;
                }

                IntPtr
                    stdoutHandle = GetStdHandle(STD_OUTPUT_HANDLE), 
                    stderrHandle = GetStdHandle(STD_ERROR_HANDLE);
                var stdoutStream = new FileStream(new SafeFileHandle(stdoutHandle, true), FileAccess.Write);
                var stderrStream = new FileStream(new SafeFileHandle(stderrHandle, true), FileAccess.Write);
                var enc = new UTF8Encoding(false, false);
                var stdout = new StreamWriter(stdoutStream, enc);
                var stderr = new StreamWriter(stderrStream, enc);
                stdout.AutoFlush = stderr.AutoFlush = true;
                Console.SetOut(stdout);
                Console.SetError(stderr);
            }
        }

        static string ExtensionId {
            get {
                /*
                var identity = WindowsIdentity.GetCurrent();
                var binLength = identity.User.BinaryLength;
                var bin = new byte[binLength];
                identity.User.GetBinaryForm(bin, 0);
                Array.Reverse(bin);

                var result = new char[32];
                Array.Copy("viramate".ToCharArray(), result, 8);
                for (int i = 0; i < binLength; i++) {
                    int j = i + 8;
                    if (j >= result.Length)
                        break;
                    result[j] = (char)((int)'a' + (bin[i] % 26));
                }
                return new string(result);
                */

                // FIXME: Generating a new extension ID requires manufacturing a public signing key... ugh
                return "fgpokpknehglcioijejfeebigdnbnokj";
            }
        }

        static bool IsRunningDirectlyFromBuild {
            get {
                return (Debugger.IsAttached || ExecutablePath.EndsWith("ViramateInstaller\\bin\\Viramate.exe"));
            }
        }

        static string DiskSourcePath {
            get {
                return Path.GetFullPath(Path.Combine(ExecutableDirectory, "..", "..", "ext"));
            }
        }

        static bool InstallFromDisk {
            get {
                bool defaultDebug = false;

                try {
                    if (!Directory.Exists(DiskSourcePath))
                        return false;

                    defaultDebug = IsRunningDirectlyFromBuild;
                } catch (Exception exc) {
                    Console.Error.WriteLine(exc);
                }

                var args = Environment.GetCommandLineArgs();
                if (args.Length <= 1)
                    return defaultDebug;
                return (defaultDebug || args.Contains("--disk")) && !args.Contains("--network");
            }
        }

        static string DataPath {
            get {
                var folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(folder, "Viramate");
            }
        }

        static string InstallerInstallPath {
            get {
                return Path.Combine(DataPath, "Installer");
            }
        }

        static string InstallerExecutablePath {
            get {
                return Path.GetFullPath(Path.Combine(DataPath, "Installer", "Viramate.exe"));
            }
        }

        static string ExtensionInstallPath {
            get {
                return Path.Combine(DataPath, "Viramate");
            }
        }

        static string ExecutablePath {
            get {
                var cb = MyAssembly.CodeBase;
                var uri = new Uri(cb);
                return Path.GetFullPath(uri.LocalPath);
            }
        }
        
        static string MiscPath {
            get {
                return Path.Combine(DataPath, "Misc");
            }
        }

        static string ExecutableDirectory {
            get {
                return Path.GetDirectoryName(ExecutablePath);
            }
        }

        static void MessagingHostMainLoop () {
            var stdin = Console.OpenStandardInput();
            var stdout = Console.OpenStandardOutput();
            var logFilePath = Path.Combine(DataPath, "installer.log");

            using (var log = new StreamWriter(logFilePath, true, Encoding.UTF8)) {
                Console.SetOut(log);
                Console.SetError(log);
                log.WriteLine($"{DateTime.UtcNow.ToLongTimeString()} > Installer started as native messaging host. Command line: {Environment.CommandLine}");
                WriteMessage(log, stdout, new { type = "serverStarting" });
                log.Flush();
                log.AutoFlush = true;

                try {
                    var wss = new WebSocketServer();
                    var t = wss.Run();
                    WriteMessage(log, stdout, new { type = "serverStarted", url = wss.Url });
                    t.Wait();
                } catch (Exception exc) {
                    log.WriteLine(exc);
                } finally {
                    log.WriteLine($"{DateTime.UtcNow.ToLongTimeString()} > Exiting.");
                }

                log.Flush();
            }

            stdout.Flush();
            stdout.Close();
            stdin.Close();
        }

        public static string ReadManifestVersion (string zipFilePath) {
            try {
                string json;
                if (zipFilePath == null) {
                    var filename = Path.Combine(ExtensionInstallPath, "manifest.json");
                    json = File.ReadAllText(filename, Encoding.UTF8);
                } else {
                    using (var zf = new ZipArchive(File.OpenRead(zipFilePath), ZipArchiveMode.Read, false))
                    using (var fileStream = zf.Entries.FirstOrDefault(e => e.FullName.EndsWith("manifest.json")).Open())
                    using (var sr = new StreamReader(fileStream, Encoding.UTF8)) {
                        json = sr.ReadToEnd();
                    }
                }
                return JsonConvert.DeserializeObject<ManifestFragment>(json).version;
            } catch (Exception exc) {
                Console.Error.WriteLine(exc);
                return null;
            }
        }

        static T ReadMessage<T> (Stream stream)
            where T : class 
        {
            var inBuf = new byte[4096000];
            if (stream.Read(inBuf, 0, inBuf.Length) == 0)
                return null;

            var lengthBytes = (int)BitConverter.ToUInt32(inBuf, 0);
            var json = Encoding.UTF8.GetString(inBuf, 4, lengthBytes);
            var result = JsonConvert.DeserializeObject<T>(json);
            return result;
        }

        static void WriteMessage<T> (StreamWriter log, Stream stream, T message)
            where T : class
        {
            var json = JsonConvert.SerializeObject(message);
            var messageByteLength = Encoding.UTF8.GetByteCount(json);
            var messageBuf = new byte[messageByteLength + 4];
            Array.Copy(BitConverter.GetBytes((uint)messageByteLength), messageBuf, 4);
            Encoding.UTF8.GetBytes(json, 0, json.Length, messageBuf, 4);
            stream.Write(messageBuf, 0, messageBuf.Length);
            stream.Flush();
            log.WriteLine($"Wrote {messageByteLength} byte(s) of JSON: {json}");
        }

        class Msg {
            public string type;
        }

        class ManifestFragment {
            public string name;
            public string version;
        }
    }
}
