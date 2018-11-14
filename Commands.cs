using System;
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
        public static Task UpdateFromFolder (string sourcePath, string destinationPath) {
            sourcePath = Path.GetFullPath(sourcePath);
            destinationPath = Path.GetFullPath(destinationPath);

            return Task.Run(() => {
                int i = 1;

                var allFiles = Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\ext\\ts\\"));

                foreach (var f in allFiles) {
                    var localPath = Path.GetFullPath(f).Replace(sourcePath, "").Substring(1);
                    var destFile = Path.Combine(destinationPath, localPath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destFile));

                    // HACK
                    if (File.Exists(destFile))
                        File.SetAttributes(destFile, FileAttributes.Normal);

                    File.Copy(f, destFile, true);
                    if (i++ % 3 == 0)
                        Console.WriteLine(localPath);
                    else
                        Console.Write(localPath + " ");
                }

                SetupDesktopIni(destinationPath);
            });
        }

        public static async Task ExtractZipFile (string zipFile, string destinationPath) {
            destinationPath = Path.GetFullPath(destinationPath);

            int i = 1;

            using (var s = File.OpenRead(zipFile))
            using (var zf = new ZipArchive(s, ZipArchiveMode.Read, true, Encoding.UTF8))
            foreach (var entry in zf.Entries) {
                if (entry.FullName.EndsWith("\\") || entry.FullName.EndsWith("/"))
                    continue;

                var destFilename = Path.Combine(destinationPath, entry.FullName);
                Directory.CreateDirectory(Path.GetDirectoryName(destFilename));

                // HACK
                if (File.Exists(destFilename))
                    File.SetAttributes(destFilename, FileAttributes.Normal);

                using (var src = entry.Open())
                using (var dst = File.Open(destFilename, FileMode.Create))
                    await src.CopyToAsync(dst);

                File.SetLastWriteTimeUtc(destFilename, entry.LastWriteTime.UtcDateTime);
                if (i++ % 3 == 0)
                    Console.WriteLine(entry.FullName);
                else
                    Console.Write(entry.FullName + " ");
            }

            SetupDesktopIni(destinationPath);
        }

        private static void SetupDesktopIni (string directory) {
            var desktopIni = Path.Combine(directory, "desktop.ini");
            if (File.Exists(desktopIni))
            try {
                (new DirectoryInfo(directory)).Attributes = FileAttributes.System;
                File.SetAttributes(desktopIni, FileAttributes.System | FileAttributes.Hidden);
            } catch (Exception exc) {
                Console.Error.WriteLine(exc);
            }
        }

        private static HttpWebRequest MakeUpdateWebRequest (string url) {
            var wr = WebRequest.CreateHttp(url);
            wr.CachePolicy = new RequestCachePolicy(
                Environment.GetCommandLineArgs().Contains("--force")
                    ? RequestCacheLevel.Reload
                    : RequestCacheLevel.Revalidate
            );
            return wr;
        }

        public struct DownloadResult {
            public string ZipPath;
            public bool   WasCached;
        }

        public static async Task<DownloadResult> DownloadLatest (string sourceUrl) {
            var wr = MakeUpdateWebRequest(sourceUrl);
            var resp = await wr.GetResponseAsync();
            if (resp.IsFromCache)
                Console.WriteLine("no new update. Using cached update.");

            Directory.CreateDirectory(MiscPath);

            var zipPath = Path.Combine(MiscPath, Path.GetFileName(resp.ResponseUri.LocalPath));
            using (var src = resp.GetResponseStream())
            using (var dst = File.Open(zipPath + ".tmp", FileMode.Create))
                await src.CopyToAsync(dst);

            if (File.Exists(zipPath))
                File.Delete(zipPath);
            File.Move(zipPath + ".tmp", zipPath);

            if (!resp.IsFromCache)
                Console.WriteLine(" done.");

            return new DownloadResult {
                ZipPath = zipPath,
                WasCached = resp.IsFromCache
            };
        }

        public static async Task<bool> AutoUpdateInstaller () {
            try {
                Console.Write($"Checking for installer update ... ");

                var result = await DownloadLatest(InstallerSourceUrl);
                if (result.WasCached)
                    return false;

                var newVersionDirectory = Path.Combine(DataPath, "Installer Update");
                Directory.CreateDirectory(newVersionDirectory);

                try {
                    if (Directory.Exists(newVersionDirectory))
                    foreach (var filename in Directory.GetFiles(newVersionDirectory))
                        File.Delete(filename);
                } catch (Exception exc) {
                    Console.Error.WriteLine($"Error while emptying new version directory: {exc.Message}");
                }

                Console.WriteLine($"Extracting {result.ZipPath} to {newVersionDirectory} ...");
                await ExtractZipFile(result.ZipPath, newVersionDirectory);
                Console.WriteLine($"done.");

                Console.WriteLine("Performing in-place update.");

                try {
                    foreach (var filename in Directory.GetFiles(InstallerInstallPath, "*.old"))
                        File.Delete(filename);
                } catch (Exception exc) {
                    Console.Error.WriteLine($"Error while deleting old update files: {exc.Message}");
                }

                foreach (var filename in Directory.GetFiles(InstallerInstallPath)) {
                    try {
                        File.Move(filename, filename + ".old");
                    } catch (Exception exc) {
                        Console.Error.WriteLine($"File operation failed on {filename}: {exc.Message}");
                    }
                }

                bool failed = false;

                try {
                    var refasm = Assembly.ReflectionOnlyLoadFrom(
                        Path.Combine(newVersionDirectory, "Viramate.exe")
                    );
                    if (refasm == null) {
                        Console.Error.WriteLine("Could not load updated installer");
                        failed = true;
                    } else {
                        Console.WriteLine($"Installing updater v{refasm.GetName().Version}");
                    }
                } catch (Exception exc) {
                    Console.Error.WriteLine($"Failed to load updated installer: {exc.Message}");
                    failed = true;
                }

                foreach (var filename in Directory.GetFiles(newVersionDirectory)) {
                    var destFilename = Path.Combine(InstallerInstallPath, Path.GetFileName(filename));
                    try {
                        File.Move(filename, destFilename);
                        Console.WriteLine($"{Path.GetFileName(filename)}");
                    } catch (Exception exc) {
                        Console.Error.WriteLine($"Copy failed for {filename}: {exc.Message}");
                        failed = true;
                    }
                }

                if (failed) {
                    Console.Error.WriteLine("Attempting to roll back failed update");
                    foreach (var filename in Directory.GetFiles(InstallerInstallPath, "*.old")) {
                        var destFilename = filename.Replace(".old", "");
                        try {
                            if (File.Exists(destFilename))
                                File.Delete(destFilename);
                        } catch (Exception exc) {
                            Console.Error.WriteLine(exc.Message);
                        }
                        File.Move(filename, destFilename);
                        Console.WriteLine($"{Path.GetFileName(filename)} -> {Path.GetFileName(destFilename)}");
                    }
                    return false;
                } else {
                    Console.Error.WriteLine("Update installed (probably)");
                    return true;
                }
            } catch (Exception exc) {
                Console.Error.WriteLine(exc.Message);
                return false;
            }
        }

        public static async Task<InstallResult> InstallExtensionFiles (bool onlyIfModified, bool? installFromDisk) {
            Directory.CreateDirectory(DataPath);

            if (File.Exists(Path.Combine(DataPath, "manifest.json"))) {
                Console.WriteLine("Detected old installation. Removing manifest. You'll need to re-install!");
                onlyIfModified = false;
                File.Delete(Path.Combine(DataPath, "manifest.json"));
            }

            if (!Directory.Exists(InstallerInstallPath))
                Directory.CreateDirectory(InstallerInstallPath);

            if (ExecutablePath != InstallerExecutablePath) {
                Console.WriteLine($"First run. Copying {ExecutablePath} to {InstallerExecutablePath}");
                try {
                    File.Copy(ExecutablePath, InstallerExecutablePath, true);
                } catch (Exception exc) {
                    Console.Error.WriteLine(exc.Message);
                }
            }

            if (installFromDisk.GetValueOrDefault(InstallFromDisk)) {
                Console.WriteLine($"Copying from {DiskSourcePath} to {ExtensionInstallPath} ...");
                await UpdateFromFolder(DiskSourcePath, ExtensionInstallPath);
                Console.WriteLine("done.");
            } else {
                DownloadResult result;

                Console.Write($"Downloading {ExtensionSourceUrl}... ");
                try {
                    result = await DownloadLatest(ExtensionSourceUrl);
                    if (result.WasCached && onlyIfModified)
                        return InstallResult.NotUpdated;

                } catch (Exception exc) {
                    Console.Error.WriteLine(exc.Message);
                    return InstallResult.Failed;
                }

                Console.WriteLine($"Current version {Program.ReadManifestVersion(null)}, new version {Program.ReadManifestVersion(result.ZipPath)}");
                Console.WriteLine($"Extracting {result.ZipPath} to {ExtensionInstallPath} ...");
                await ExtractZipFile(result.ZipPath, ExtensionInstallPath);
                Console.WriteLine($"done.");
            }

            return InstallResult.Updated;
        }

        static Stream OpenResource (string name) {
            return MyAssembly.GetManifestResourceStream("Viramate." + name.Replace("/", ".").Replace("\\", "."));
        }

        static void LogWrite (StreamWriter log, string text = "") {
            if (log != null)
                log.WriteLine(text);
            Console.WriteLine(text);
        }

        public static async Task InstallExtension () {
            var allowAutoClose = true;

            StreamWriter log = null;
            try {
                log = new StreamWriter(LogFilePath, false, Encoding.UTF8);
                log.AutoFlush = true;
            } catch (Exception exc) {
                Console.WriteLine($"Failed creating log file: {exc}");
                try {
                    log = new StreamWriter(LogFilePath, true, Encoding.UTF8);
                    log.AutoFlush = true;
                } catch (Exception exc2) {
                    Console.WriteLine($"Failed appending to log file: {exc2}");
                }
            }

            try {
                Console.WriteLine();
                Console.WriteLine($"Viramate Installer v{MyAssembly.GetName().Version}");
                if (Environment.GetCommandLineArgs().Contains("--version"))
                    return;

                Console.WriteLine($"Use viramate -? for info on command line switches");

                if (Environment.GetCommandLineArgs().Contains("--update"))
                    await AutoUpdateInstaller();

                LogWrite(log, "Installing extension. This'll take a moment...");

                if (await InstallExtensionFiles(false, null) != InstallResult.Failed) {
                    LogWrite(log, $"Extension id: {ExtensionId}");

                    string manifestText;
                    using (var s = new StreamReader(OpenResource("nmh.json"), Encoding.UTF8))
                        manifestText = s.ReadToEnd();

                    manifestText = manifestText
                        .Replace(
                            "$executable_path$", 
                            InstallerExecutablePath.Replace("\\", "\\\\").Replace("\"", "\\\"")
                        ).Replace(
                            "$extension_id$", ExtensionId
                        );

                    var manifestPath = Path.Combine(MiscPath, "nmh.json");
                    Directory.CreateDirectory(MiscPath);
                    File.WriteAllText(manifestPath, manifestText);

                    Directory.CreateDirectory(Path.Combine(DataPath, "Help"));
                    foreach (var n in MyAssembly.GetManifestResourceNames()) {
                        if (!n.EndsWith(".gif") && !n.EndsWith(".png"))
                            continue;

                        var destinationPath = Path.Combine(DataPath, n.Replace("Viramate.", "").Replace("Help.", "Help\\"));
                        using (var src = MyAssembly.GetManifestResourceStream(n))
                        using (var dst = File.Open(destinationPath, FileMode.Create))
                            await src.CopyToAsync(dst);
                    }

                    const string keyName = @"Software\Google\Chrome\NativeMessagingHosts\com.viramate.installer";
                    using (var key = Registry.CurrentUser.CreateSubKey(keyName)) {
                        LogWrite(log, $"{keyName}\\@ = {manifestPath}");
                        key.SetValue(null, manifestPath);
                    }

                    try {
                        WebSocketServer.SetupFirewallRule();
                    } catch (Exception exc) {
                        LogWrite(log, $"Failed to install firewall rule: {exc}");
                        allowAutoClose = false;
                    }

                    string helpFileText;
                    using (var s = new StreamReader(OpenResource("Help/index.html"), Encoding.UTF8))
                        helpFileText = s.ReadToEnd();

                    helpFileText = Regex.Replace(
                        helpFileText, 
                        @"\<pre\ id='install_path'>[^<]*\</pre\>", 
                        @"<pre id='install_path'>" + DataPath + "</pre>"
                    );

                    var helpFilePath = Path.Combine(DataPath, "Help", "index.html");
                    File.WriteAllText(helpFilePath, helpFileText);

                    LogWrite(log, $"Viramate v{ReadManifestVersion(null)} has been installed.");
                    if (!Environment.GetCommandLineArgs().Contains("--nohelp")) {
                        LogWrite(log, "Opening install instructions...");
                        Process.Start(helpFilePath);
                    } else if (!Debugger.IsAttached && !IsRunningInsideCmd) {
                        return;
                    }

                    if (!Environment.GetCommandLineArgs().Contains("--nodir")) {
                        LogWrite(log, "Waiting, then opening install directory...");
                        await Task.Delay(2000);
                        Process.Start(DataPath);
                    }

                    if (!Debugger.IsAttached)
                        Thread.Sleep(1000 * 30);

                    Environment.Exit(0);
                } else {
                    await AutoUpdateInstaller();

                    LogWrite(log, "Failed to install extension.");

                    if (!Debugger.IsAttached)
                        Thread.Sleep(1000 * 30);

                    Environment.Exit(1);
                }
            } finally {
                if (log != null)
                    log.Dispose();
            }
        }

        public static void PrintHelp () {
            Console.WriteLine();
            Console.WriteLine($"Viramate Installer v{MyAssembly.GetName().Version}");
            Console.WriteLine(
@"-? /?
    Print help
--version
    Print version number and quit
--update
    Update the installer even if install succeeds
--nodir
    Don't open the install directory
--nohelp
    Don't open the help webpage after install
--disk
    Only install from a local directory instead of the internet (for debugging purposes)
--network
    Only install from the internet, not a local directory
--force
    Force install/update even if nothing is changed"
            );
        }
    }

    public enum InstallResult {
        Failed = 0,
        NotUpdated,
        Updated
    }
}
