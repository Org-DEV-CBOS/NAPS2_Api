using System.Net;
using System.Text;
using Autofac;
using Eto.Forms;
using NAPS2.EtoForms;
using NAPS2.EtoForms.Desktop;
using NAPS2.EtoForms.Ui;
using NAPS2.Images;
using NAPS2.Pdf;
using NAPS2.Scan;
using NAPS2.Platform;
using NAPS2.Platform.Windows;
using NAPS2.Config;
using NAPS2.ImportExport;
using NAPS2.Scan.Batch;
using System.Threading;

namespace NAPS2.LocalServer;

public class LocalScanServer : IDisposable
{
    private readonly DesktopFormProvider _desktopFormProvider;
    private readonly IDesktopScanController _desktopScanController;
    private readonly UiImageList _imageList;
    private readonly ScanningContext _scanningContext;
    private readonly HttpListener _listener = new();
    private readonly object _scanLock = new();
    private bool _scanInProgress;
    private CancellationTokenSource _cts = new();
    private readonly IProfileManager _profileManager;
    private readonly IDesktopSubFormController _desktopSubFormController;
    private readonly Naps2Config _config;

    public LocalScanServer(DesktopFormProvider desktopFormProvider,
        IDesktopScanController desktopScanController,
        UiImageList imageList,
        ScanningContext scanningContext,
        IProfileManager profileManager,
        IDesktopSubFormController desktopSubFormController,
        Naps2Config config)
    {
        _desktopFormProvider = desktopFormProvider;
        _desktopScanController = desktopScanController;
        _imageList = imageList;
        _scanningContext = scanningContext;
        _profileManager = profileManager;
        _desktopSubFormController = desktopSubFormController;
        _config = config;

        // Bind only to loopback for local-only access
        _listener.Prefixes.Add("http://127.0.0.1:8765/");
        _listener.Start();
        _ = Task.Run(ListenLoop);
    }

    private async Task ListenLoop()
    {
        try
        {
            while (!_cts.IsCancellationRequested)
            {
                var ctx = await _listener.GetContextAsync().ConfigureAwait(false);
                _ = Task.Run(() => HandleRequest(ctx));
            }
        }
        catch (ObjectDisposedException)
        {
        }
        catch (HttpListenerException)
        {
        }
    }

    private async Task HandleRequest(HttpListenerContext context)
    {
        try
        {
            var req = context.Request;
            var res = context.Response;

            if (req.HttpMethod != "GET")
            {
                res.StatusCode = 405;
                res.Close();
                return;
            }

            // Normalize path
            var path = req.Url?.AbsolutePath?.TrimEnd('/') ?? string.Empty;
            if (!string.Equals(path, "/scan", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(path, "/batch-scan", StringComparison.OrdinalIgnoreCase))
            {
                res.StatusCode = 404;
                res.Close();
                return;
            }

            lock (_scanLock)
            {
                if (_scanInProgress)
                {
                    res.StatusCode = 409; // Conflict - already scanning
                    res.Close();
                    return;
                }
                _scanInProgress = true;
            }

            try
            {
                // Bring the main window to foreground on UI thread
                Invoker.Current.Invoke(() =>
                {
                    var formOnTop = Application.Instance.Windows.Last();
                    if (PlatformCompat.System.CanUseWin32)
                    {
                        if (formOnTop.WindowState == WindowState.Minimized)
                        {
                            Win32.ShowWindow(formOnTop.NativeHandle, Win32.ShowWindowCommands.Restore);
                        }
                        Win32.SetForegroundWindow(formOnTop.NativeHandle);
                    }
                    else
                    {
                        formOnTop.BringToFront();
                    }
                });

                List<(string fileName, MemoryStream stream)> pdfStreams;
                if (string.Equals(path, "/scan", StringComparison.OrdinalIgnoreCase))
                {
                    // Single scan: Run via ScanDefault and capture newly added images
                    var scanTcs = new TaskCompletionSource<List<ProcessedImage>>();
                    Invoker.Current.InvokeDispatch(async () =>
                    {
                        try
                        {
                            int startCount = _imageList.Images.Count;
                            await _desktopScanController.ScanDefault();
                            var newUiImages = _imageList.Images.Skip(startCount).ToList();
                            var clones = newUiImages.Select(i => i.GetClonedImage()).ToList();
                            scanTcs.TrySetResult(clones);
                        }
                        catch (Exception ex)
                        {
                            scanTcs.TrySetException(ex);
                        }
                    });

                    var clonedImages = await scanTcs.Task.ConfigureAwait(false);

                if (clonedImages.Count == 0)
                {
                    res.StatusCode = 204; // No Content (user canceled or no images)
                    res.Close();
                    return;
                }

                    // Determine grouping based on selected profile's AutoSave separator (per page/scan), default single file
                var sep = SaveSeparator.None;
                var selProfile = _profileManager.DefaultProfile;
                if (selProfile?.EnableAutoSave == true && selProfile.AutoSaveSettings != null)
                {
                    sep = selProfile.AutoSaveSettings.Separator;
                }

                    var groups = SaveSeparatorHelper.SeparateScans(new[] { clonedImages }, sep).ToList();

                    // Export groups to PDFs
                var exporter = new PdfExporter(_scanningContext);
                    pdfStreams = new List<(string fileName, MemoryStream stream)>();
                try
                {
                        for (int i = 0; i < groups.Count; i++)
                    {
                        var group = groups[i];
                        var ms = new MemoryStream();
                        var ok = await exporter.Export(ms, group);
                        if (!ok)
                        {
                            res.StatusCode = 500;
                            res.Close();
                            return;
                        }
                        ms.Position = 0;

                            string baseName = "scan";
                        string? configuredName = null;
                        if (selProfile?.AutoSaveSettings != null)
                        {
                            configuredName = Path.GetFileName(selProfile.AutoSaveSettings.FilePath);
                        }
                        if (!string.IsNullOrWhiteSpace(configuredName))
                        {
                            baseName = Path.GetFileNameWithoutExtension(configuredName);
                        }
                            var suffix = groups.Count > 1 ? $"_{i + 1}" : "";
                        var fileName = baseName + suffix + ".pdf";
                            pdfStreams.Add((fileName, ms));
                    }
                }
                finally
                {
                        foreach (var img in clonedImages)
                    {
                        img.Dispose();
                    }
                }
                }
                else
                {
                    // Batch scan: Show BatchScanForm and, after it completes, return PDFs based on chosen output
                    var startTimeUtc = DateTime.UtcNow.AddSeconds(-1);
                    int initialCount = _imageList.Images.Count;
                    Invoker.Current.Invoke(() => _desktopSubFormController.ShowBatchScanForm());

                    var bs = _config.Get(c => c.BatchSettings);

                    if (bs.OutputType == BatchOutputType.Load)
                    {
                        // Images were loaded into the app; collect delta and export
                        var newUiImages = _imageList.Images.Skip(initialCount).ToList();
                        var clones = newUiImages.Select(i => i.GetClonedImage()).ToList();
                        if (clones.Count == 0)
                        {
                            res.StatusCode = 204;
                            res.Close();
                            return;
                        }
                        var groups = bs.SaveSeparator == SaveSeparator.None
                            ? new List<List<ProcessedImage>> { clones }
                            : SaveSeparatorHelper.SeparateScans(new[] { clones }, bs.SaveSeparator).ToList();
                        var exporter = new PdfExporter(_scanningContext);
                        pdfStreams = new List<(string fileName, MemoryStream stream)>();
                        for (int i = 0; i < groups.Count; i++)
                        {
                            var ms = new MemoryStream();
                            var ok = await exporter.Export(ms, groups[i]);
                            if (!ok)
                            {
                                res.StatusCode = 500;
                                res.Close();
                                return;
                            }
                            ms.Position = 0;
                            var baseName = Path.GetFileNameWithoutExtension(bs.SavePath) ?? "batch";
                            var suffix = groups.Count > 1 ? $"_{i + 1}" : "";
                            pdfStreams.Add((baseName + suffix + ".pdf", ms));
                        }
                        // Dispose clones after export
                        foreach (var img in clones) img.Dispose();
                    }
                    else
                    {
                        // Files were saved to disk; collect the files created since start time and return them
                        string expanded = Placeholders.NonNumeric.WithDate(DateTime.Now)
                            .Substitute(bs.SavePath, incrementIfExists: false) ?? bs.SavePath!;
                        var dir = Path.GetDirectoryName(expanded) ?? Environment.CurrentDirectory;
                        var baseName = Path.GetFileNameWithoutExtension(expanded);
                        var ext = Path.GetExtension(expanded);
                        var candidates = Directory.Exists(dir)
                            ? new DirectoryInfo(dir)
                                .GetFiles("*" + ext)
                                .Where(f => f.LastWriteTimeUtc >= startTimeUtc &&
                                            f.Name.StartsWith(baseName, StringComparison.InvariantCultureIgnoreCase))
                                .OrderBy(f => f.Name)
                                .ToList()
                            : new List<FileInfo>();
                        if (candidates.Count == 0)
                        {
                            res.StatusCode = 204;
                            res.Close();
                            return;
                        }
                        pdfStreams = new List<(string fileName, MemoryStream stream)>();
                        foreach (var fi in candidates)
                        {
                            var ms = new MemoryStream(await File.ReadAllBytesAsync(fi.FullName));
                            ms.Position = 0;
                            pdfStreams.Add((fi.Name, ms));
                        }
                    }
                }

                // Build multipart/form-data with one part per PDF
                var boundary = "--------------------------" + DateTime.UtcNow.Ticks.ToString("x");
                res.StatusCode = 200;
                res.ContentType = $"multipart/form-data; boundary={boundary}";

                long totalLength = 0;
                var partHeaders = new List<byte[]>();
                foreach (var part in pdfStreams)
                {
                    var h = $"--{boundary}\r\n" +
                            $"Content-Disposition: form-data; name=\"files\"; filename=\"{part.fileName}\"\r\n" +
                            "Content-Type: application/pdf\r\n\r\n";
                    var hb = Encoding.UTF8.GetBytes(h);
                    partHeaders.Add(hb);
                    totalLength += hb.Length + part.stream.Length + 2; // +2 for CRLF between parts
                }
                var endBoundaryBytes = Encoding.UTF8.GetBytes($"--{boundary}--\r\n");
                totalLength += endBoundaryBytes.Length;
                res.ContentLength64 = totalLength;

                for (int idx = 0; idx < pdfStreams.Count; idx++)
                {
                    var (fileName, stream) = pdfStreams[idx];
                    var hb = partHeaders[idx];
                    await res.OutputStream.WriteAsync(hb, 0, hb.Length);
                    await stream.CopyToAsync(res.OutputStream);
                    await res.OutputStream.WriteAsync(Encoding.UTF8.GetBytes("\r\n"), 0, 2);
                }
                await res.OutputStream.WriteAsync(endBoundaryBytes, 0, endBoundaryBytes.Length);
                res.OutputStream.Flush();
                res.Close();
            }
            finally
            {
                lock (_scanLock)
                {
                    _scanInProgress = false;
                }
            }
        }
        catch
        {
            try { context.Response.StatusCode = 500; context.Response.Close(); } catch { }
        }
    }

    public void Dispose()
    {
        _cts.Cancel();
        _listener.Close();
    }
}


