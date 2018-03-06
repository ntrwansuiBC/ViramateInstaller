﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
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
    class Program {
        [DllImport(
            "kernel32.dll", EntryPoint = "GetStdHandle", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport(
            "kernel32.dll", EntryPoint = "AllocConsole", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int AllocConsole();

        private const int STD_INPUT_HANDLE = -10;
        private const int STD_OUTPUT_HANDLE = -11;
        private const int STD_ERROR_HANDLE = -12;
        private const int MY_CODE_PAGE = 437;

        static Assembly MyAssembly;

        static void Main (string[] args) {
            MyAssembly = Assembly.GetExecutingAssembly();

            try {
                if ((args.Length == 0) || !args[0].StartsWith("chrome-extension://")) {
                    InitConsole();
                    InstallExtension().Wait();
                } else
                    MessagingHostMainLoop();
            } catch (Exception exc) {
                Console.Error.WriteLine("Uncaught: {0}", exc);
                Environment.ExitCode = 1;

                if (Debugger.IsAttached) {
                    Console.Error.WriteLine("Press enter to exit.");
                    Console.ReadLine();
                }
            }
        }

        static void InitConsole () {
            AllocConsole();
            IntPtr
                stdinHandle = GetStdHandle(STD_INPUT_HANDLE),
                stdoutHandle = GetStdHandle(STD_OUTPUT_HANDLE), 
                stderrHandle = GetStdHandle(STD_ERROR_HANDLE);
            var stdinStream = new FileStream(new SafeFileHandle(stdinHandle, true), FileAccess.Read);
            var stdoutStream = new FileStream(new SafeFileHandle(stdoutHandle, true), FileAccess.Write);
            var stderrStream = new FileStream(new SafeFileHandle(stderrHandle, true), FileAccess.Write);
            var stdin = new StreamReader(stdinStream, Encoding.UTF8);
            var stdout = new StreamWriter(stdoutStream, Encoding.UTF8);
            var stderr = new StreamWriter(stderrStream, Encoding.UTF8);
            stdout.AutoFlush = stderr.AutoFlush = true;
            Console.SetIn(stdin);
            Console.SetOut(stdout);
            Console.SetError(stderr);
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

        static bool InstallFromDisk {
            get {
                var defaultDebug = (Debugger.IsAttached || ExecutablePath.EndsWith("ViramateInstaller\\bin\\Viramate.exe"));
                var args = Environment.GetCommandLineArgs();
                if (args.Length <= 1)
                    return defaultDebug;

                return (defaultDebug || (args[1] == "--disk")) && (args[1] != "--network");
            }
        }

        const string SourceUrl = "http://luminance.org/vm/ext.zip";

        static string InstallPath {
            get {
                var folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(folder, "Viramate");
            }
        }

        static string ExecutablePath {
            get {
                var cb = MyAssembly.CodeBase;
                var uri = new Uri(cb);
                return uri.LocalPath;
            }
        }

        static string ExecutableDirectory {
            get {
                return Path.GetDirectoryName(ExecutablePath);
            }
        }

        public static Task UpdateFromFolder (string sourcePath) {
            return Task.Run(() => {
                var allFiles = Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\ext\\ts\\"));

                foreach (var f in allFiles) {
                    var localPath = Path.GetFullPath(f).Replace(sourcePath, "").Substring(1);
                    var destinationPath = Path.Combine(InstallPath, localPath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
                    File.Copy(f, destinationPath, true);
                    Console.Error.WriteLine(destinationPath);
                }
            });
        }

        public static async Task UpdateFromZipFile (string sourcePath) {
            using (var s = File.OpenRead(sourcePath))
            using (var zf = new ZipArchive(s, ZipArchiveMode.Read, true, Encoding.UTF8))
            foreach (var entry in zf.Entries) {
                if (entry.FullName.EndsWith("\\") || entry.FullName.EndsWith("/"))
                    continue;

                var destFilename = Path.Combine(InstallPath, entry.FullName);
                Directory.CreateDirectory(Path.GetDirectoryName(destFilename));

                using (var src = entry.Open())
                using (var dst = File.OpenWrite(destFilename))
                    await src.CopyToAsync(dst);

                File.SetLastWriteTimeUtc(destFilename, entry.LastWriteTime.UtcDateTime);
                Console.Error.WriteLine(destFilename);
            }
        }

        public static async Task<string> DownloadLatest (string sourceUrl) {
            var wc = new WebClient();
            var zipPath = Path.Combine(InstallPath, "latest.zip");
            await wc.DownloadFileTaskAsync(sourceUrl, zipPath + ".tmp");
            if (File.Exists(zipPath))
                File.Delete(zipPath);
            File.Move(zipPath + ".tmp", zipPath);
            return zipPath;
        }

        public static async Task<bool> InstallExtensionFiles (bool? installFromDisk = null) {
            Directory.CreateDirectory(InstallPath);

            if (installFromDisk.GetValueOrDefault(InstallFromDisk)) {
                var sourcePath = Path.GetFullPath(Path.Combine(ExecutableDirectory, "..", "..", "ext"));
                Console.Error.WriteLine($"Copying from {sourcePath} to {InstallPath} ...");
                await UpdateFromFolder(sourcePath);
                Console.Error.WriteLine("done.");
                return true;
            } else {
                string zipPath = null;

                Console.Error.Write($"Downloading {SourceUrl}... ");
                try {
                    zipPath = await DownloadLatest(SourceUrl);
                    Console.Error.WriteLine("done.");
                } catch (WebException exc) {
                    Console.Error.WriteLine(exc.Message);
                    return false;
                }

                Console.Error.WriteLine($"Extracting {zipPath} to {InstallPath} ...");
                await UpdateFromZipFile(zipPath);
                Console.Error.WriteLine($"done.");

                return true;
            }
        }

        static Stream OpenResource (string name) {
            return MyAssembly.GetManifestResourceStream("Viramate." + name.Replace("/", ".").Replace("\\", "."));
        }

        public static async Task InstallExtension () {
            Console.Error.WriteLine("Installing extension. This'll take a moment...");

            if (await InstallExtensionFiles()) {
                Console.Error.WriteLine($"Extension id: {ExtensionId}");

                string manifestText;
                using (var s = new StreamReader(OpenResource("nmh.json"), Encoding.UTF8))
                    manifestText = s.ReadToEnd();

                manifestText = manifestText
                    .Replace(
                        "$executable_path$", 
                        ExecutablePath.Replace("\\", "\\\\").Replace("\"", "\\\"")
                    ).Replace(
                        "$extension_id$", ExtensionId
                    );
                var manifestPath = Path.Combine(InstallPath, "nmh.json");
                File.WriteAllText(manifestPath, manifestText);

                const string keyName = @"Software\Google\Chrome\NativeMessagingHosts\com.viramate.installer";
                using (var key = Registry.CurrentUser.CreateSubKey(keyName, true)) {
                    Console.Error.WriteLine($"{keyName}\\@ = {manifestPath}");
                    key.SetValue(null, manifestPath);
                }

                try {
                    WebSocketServer.SetupFirewallRule();
                } catch (Exception exc) {
                    Console.Error.WriteLine("Failed to install firewall rule: {0}", exc);
                }

                Directory.CreateDirectory(Path.Combine(InstallPath, "Help"));
                foreach (var n in MyAssembly.GetManifestResourceNames()) {
                    if (!n.EndsWith(".png"))
                        continue;

                    var destinationPath = Path.Combine(InstallPath, n.Replace("Viramate.", "").Replace("Help.", "Help\\"));
                    using (var src = MyAssembly.GetManifestResourceStream(n))
                    using (var dst = File.OpenWrite(destinationPath))
                        await src.CopyToAsync(dst);
                }

                string helpFileText;
                using (var s = new StreamReader(OpenResource("Help/index.html"), Encoding.UTF8))
                    helpFileText = s.ReadToEnd();

                helpFileText = Regex.Replace(
                    helpFileText, 
                    @"\<pre\ id='extension_path'>[^<]*\</pre\>", 
                    @"<pre id='extension_path'>" + InstallPath + "</pre>"
                );

                var helpFilePath = Path.Combine(InstallPath, "Help", "index.html");
                File.WriteAllText(helpFilePath, helpFileText);

                Console.Error.WriteLine("Viramate has been downloaded. Opening install instructions...");
                Process.Start(helpFilePath);
            } else {
                Console.Error.WriteLine("Failed to install extension.");
            }
            Console.Error.WriteLine("Press enter to exit.");
            Console.ReadLine();
        }

        static void MessagingHostMainLoop () {
            var stdin = Console.OpenStandardInput();
            var stdout = Console.OpenStandardOutput();
            var logFilePath = Path.Combine(InstallPath, "installer.log");

            using (var log = new StreamWriter(logFilePath, true, Encoding.UTF8)) {
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
                    log.WriteLine($"Exiting.");
                }

                log.Flush();
            }

            stdout.Flush();
            stdout.Close();
            stdin.Close();
        }

        static string ReadManifestVersion () {
            var filename = Path.Combine(InstallPath, "manifest.json");
            try {
                var json = File.ReadAllText(filename, Encoding.UTF8);
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