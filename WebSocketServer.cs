using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Security.Cryptography;
using NetFwTypeLib;

namespace Viramate {
    public class WebSocketServer : IDisposable {
        public TimeSpan IdleTimeout = TimeSpan.FromMinutes(1);

        public HttpListener Listener { get; private set; }
        public bool IsDisposed { get; private set; }

        private readonly object Lock = new object();

        public string Url = "ws://127.0.0.1:8678/";

        private DateTime LastActivity = default(DateTime);

        public WebSocketServer () {
            LastActivity = DateTime.UtcNow;
            try {
                Init();
            } catch (Exception exc) {
                Console.Error.WriteLine("Failed to init http server: {0}", exc);
                throw;
            }
        }

        public async Task<bool> Run () {
            try {
                Listener.Start();
            } catch (Exception exc) {
                Console.Error.WriteLine("Failed to start http server: {0}", exc);
                return false;
            }

            Console.Error.WriteLine();

            try {
                SetupFirewallRule();
            } catch (Exception exc) {
                Console.Error.WriteLine("Failed to set-up firewall rule: {0}", exc);
            }

            IdleTimeoutTask();

            HttpListener listener;
            while (!IsDisposed) {
                lock (Lock)
                    listener = Listener;

                var context = await listener.GetContextAsync();
                try {
                    await TryHandleRequest(context);
                } catch (Exception exc) {
                    Console.Error.WriteLine("Unhandled exception: {0}", exc);
                }
            }

            return true;
        }

        public static void SetupFirewallRule () {
            var firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(
                Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));

            try {
                var existingRule = firewallPolicy.Rules.Item("Viramate Installer");

                if (existingRule != null) {
                    if (!existingRule.Enabled)
                        Console.Error.WriteLine("WARNING: Viramate windows firewall rule is currently disabled");

                    return;
                }
            } catch (Exception) {
            }

            var firewallRule = (INetFwRule)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
            firewallRule.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
            firewallRule.Description = "Allows Viramate chrome extension to communicate with installer";
            firewallRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
            firewallRule.Enabled = true;
            firewallRule.InterfaceTypes = "All";
            firewallRule.Protocol = 6; // TCP
            firewallRule.LocalPorts = "8678";
            firewallRule.Name = "Viramate Installer";

            firewallPolicy.Rules.Add(firewallRule);
            Console.Error.WriteLine("Added exception to Windows firewall");
        }

        private bool IsAuthenticated (HttpListenerContext context) {
            if (context == null)
                return false;

            switch (context.Request.HttpMethod.ToLower()) {
                case "options":
                case "post":
                    if (context.Request.Url.LocalPath == "/vm")
                        return true;

                    break;
            }

            if (context.Request.IsWebSocketRequest)
                return true;

            return false;
        }

        private async Task TryHandleRequest (HttpListenerContext context) {
            context.Response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate");

            if (!IsAuthenticated(context))
                return;

            Console.Error.WriteLine("ws connect {0}", context.Request.RawUrl);

            RequestTask(context);
        }

        private void Init () {
            Listener = new HttpListener();
            Listener.Prefixes.Add(Url.Replace("ws:", "http:"));
        }

        private async Task IdleTimeoutTask () {
            while (!this.IsDisposed) {
                await Task.Delay(IdleTimeout);
                var timeIdle = DateTime.UtcNow - LastActivity;
                if (timeIdle > IdleTimeout) {
                    Console.WriteLine("Shutting down due to idle timeout.");
                    this.Dispose();
                }
            }
        }

        private static Task Send (WebSocket ws, string text) {
            var buffer = Encoding.UTF8.GetBytes(text);
            return ws.SendAsync(new ArraySegment<byte>(buffer), WebSocketMessageType.Text, true, CancellationToken.None);
        }

        private async Task HandleSyncRequest (HttpListenerContext context) {
            try {
                context.Response.AddHeader("installerversion", Program.MyAssembly.GetName().Version.ToString());
                context.Response.ContentType = "application/json";

                var readBuf = new byte[102400];
                var count = context.Request.InputStream.Read(readBuf, 0, readBuf.Length);
                var result = await ProcessCommand(readBuf, count);

                if (result != null) {
                    var json = JsonConvert.SerializeObject(result);
                    var bytes = Encoding.UTF8.GetBytes(json);
                    context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                }
            } catch (Exception exc) {
                context.Response.StatusCode = 500;
            } finally {
                context.Response.Close();
            }
        }

        private async Task RequestTask (HttpListenerContext context) {
            var method = context.Request.HttpMethod.ToLower();
            context.Response.AddHeader("Access-Control-Allow-Origin", "chrome-extension://fgpokpknehglcioijejfeebigdnbnokj");

            if (method == "post") {
                Console.Error.WriteLine("Handling sync request");
                try {
                    await HandleSyncRequest(context);
                } catch (Exception exc) {
                    Console.Error.WriteLine(exc);
                }
            } else if (method == "options") {
                var r = context.Response;
                r.AddHeader("Access-Control-Allow-Headers", "Content-Type, Accept, X-Requested-With");
                r.AddHeader("Access-Control-Allow-Methods", "GET, POST");
                r.AddHeader("Access-Control-Max-Age", "1728000");
            } else {
                HttpListenerWebSocketContext wsc = null;
                try {
                    wsc = await context.AcceptWebSocketAsync(null);
                } catch (Exception _) {
                    Console.Error.WriteLine("Accept failed {0}", _);
                    return;
                }

                try {
                    var ws = wsc.WebSocket;
                    await SocketMainLoop(ws);
                } catch (Exception exc) {
                    Console.Error.WriteLine(exc);
                } finally {
                        if (wsc != null)
                            wsc.WebSocket.Dispose();
                }

                Console.Error.WriteLine("ws disconnect");
            }
        }

        private async Task SocketMainLoop (WebSocket ws) {
            LastActivity = DateTime.UtcNow;

            var readBuffer = new ArraySegment<byte>(new byte[10240]);

            WebSocketReceiveResult wsrr;

            while (ws.State == WebSocketState.Open) {
                int readOffset = 0, count = 0;
                try {
                    do {
                        wsrr = await ws.ReceiveAsync(new ArraySegment<byte>(readBuffer.Array, readOffset, readBuffer.Count - readOffset), CancellationToken.None);
                        readOffset += wsrr.Count;
                        count += wsrr.Count;
                    } while (!wsrr.EndOfMessage);
                } catch (Exception exc) {
                    Console.Error.WriteLine(exc);
                    break;
                }

                var result = await ProcessCommand(readBuffer.Array, count);
                if (result != null)
                    await Send(ws, result);
            }

            Dispose();
        }

        private async Task<object> ProcessCommand (byte[] buffer, int count) {
            string json = "";
            JObject obj = null;
            try {
                json = Encoding.UTF8.GetString(buffer, 0, count);
                obj = (JObject)JsonConvert.DeserializeObject(json);
            } catch (Exception exc) {
                Console.Error.WriteLine($"JSON parse error: {exc}");
                Console.Error.WriteLine("// Failed json blob below //");
                Console.Error.WriteLine(json);
                Console.Error.WriteLine("// Failed json blob above //");
            }

            if (obj == null)
                return null;

            Console.Error.WriteLine($"Running command: {json}");

            LastActivity = DateTime.UtcNow;

            var msgType = (string)obj["type"];
            int? id = null;
            if (obj.ContainsKey("id"))
                id = obj["id"].ToObject<int>();

            switch (msgType) {
                case "getVersion": {
                    return new {
                        result = new {
                            extension = Program.ReadManifestVersion(null),
                            installer = Program.MyAssembly.GetName().Version.ToString()
                        },
                        id = id
                    };
                }

                case "forceInstallUpdate":
                case "installUpdate": {
                    try {
                        var result = await Program.InstallExtensionFiles((msgType == "installUpdate"), null);
                        return new {
                            result = new {
                                result = result.ToString(),
                                installedVersion = Program.ReadManifestVersion(null)
                            },
                            id = id,
                        };
                    } catch (Exception exc) {
                        Console.Error.WriteLine(exc);
                        return new {
                            result = new {
                                result = "Failed",
                                error = exc.Message
                            },
                            id = id,
                        };
                    }
                }

                case "downloadNewUpdate": {
                    try {
                        var result = await Program.DownloadLatest(Program.ExtensionSourceUrl);
                        return new {
                            result = new {
                                ok = true,
                                installedVersion = Program.ReadManifestVersion(null),
                                downloadedVersion = Program.ReadManifestVersion(result.ZipPath)
                            },
                            id = id,
                        };
                    } catch (Exception exc) {
                        Console.Error.WriteLine(exc);
                        return new {
                            result = new {
                                ok = false,
                                error = exc.Message
                            },
                            id = id,
                        };
                    }
                }

                case "autoUpdateInstaller": {
                    try {
                        var result = await Program.AutoUpdateInstaller();
                        if (result) {
                            Console.WriteLine("Exiting in response to successful installer update request.");
                            ThreadPool.QueueUserWorkItem((_) => {
                                Thread.Sleep(1000);
                                Environment.Exit(0);
                            });
                        }

                        return new {
                            result = result,
                            id = id
                        };
                    } catch (Exception exc) {
                        Console.Error.WriteLine(exc);
                        return new {
                            result = false,
                            id = id
                        };
                    }
                }

                default:
                    Console.Error.WriteLine("Unhandled message {0}", msgType);
                    if (id != null)
                        return new {
                            id = id
                        };
                    else
                        return null;
            }
        }

        public Task Send<T> (WebSocket ws, T value) {
            var json = JsonConvert.SerializeObject(value);
            var bytes = Encoding.UTF8.GetBytes(json);

            var tcs = new TaskCompletionSource<int>();

            lock (ws) {
                if (IsDisposed) {
                    tcs.SetCanceled();
                    return tcs.Task;
                }

                return ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None);
            }
        }

        public void Dispose () {
            if (IsDisposed)
                return;

            IsDisposed = true;
            Listener.Stop();
            Listener.Close();
        }
    }
}
