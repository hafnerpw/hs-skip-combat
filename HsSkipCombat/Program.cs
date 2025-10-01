using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

#nullable enable

namespace HearthstoneBattlegroundsSkip
{
    internal static partial class Program
    {
        private const string AppName = "HearthstoneBattlegroundsSkip";
        private static readonly string BasePath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppName);
        private static readonly string ConfigPath = Path.Combine(BasePath, "config.json");

        private static Config _cfg = Config.Default();
        private static readonly object ConsoleLock = new();
        private static int _statusBlockStartRow = -1;
        private static int _statusBlockLineCount;
        private static CancellationTokenSource? _ocrCts;
        private static Thread? _ocrThread;
        private static string? _lastOcrText;

        private static bool IsOcrCapturing => _ocrThread is { IsAlive: true };

        private static void WithColor(ConsoleColor color, Action action)
        {
            var previous = Console.ForegroundColor;
            Console.ForegroundColor = color;
            try
            {
                action();
            }
            finally
            {
                Console.ForegroundColor = previous;
            }
        }

        private static void WriteColoredLine(string text, ConsoleColor color)
            => WithColor(color, () => Console.WriteLine(text));

        private static void WriteColored(string text, ConsoleColor color)
            => WithColor(color, () => Console.Write(text));

        private static void WriteMenuOption(string key, string description)
        {
            WriteColored($"  {key}) ", ConsoleColor.Green);
            Console.WriteLine(description);
        }

        private static void WriteStatusValue(string label, string value, ConsoleColor color)
        {
            Console.Write($" {label,-19}: ");
            WriteColoredLine(value, color);
        }

        [STAThread]
        private static async Task<int> Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.Title = "Hearthstone Battlegrounds Skip (Console)";

            EnsureAppFolder();
            LoadConfig();

            AppDomain.CurrentDomain.ProcessExit += (_, __) =>
            {
                StopOcrCapture();
                SafeReconnect();
            };
            Console.CancelKeyPress += (_, e) =>
            {
                e.Cancel = true;
                StopOcrCapture();
                SafeReconnect();
                Environment.Exit(0);
            };

            // Optional quick commands
            if (args.Length > 0)
            {
                switch (args[0].ToLowerInvariant())
                {
                    case "skip":
                        await EnsureAdminOrRelaunchAsync();
                        await EnsureRuleExistsAsync();
                        await DoSkipAsync(_cfg.DisconnectedSeconds);
                        return 0;
                    case "status":
                        ShowStatus();
                        return 0;
                }
            }

            await MenuLoopAsync();
            return 0;
        }

        private static void EnsureAppFolder()
        {
            if (!Directory.Exists(BasePath))
                Directory.CreateDirectory(BasePath);
        }

        private static void LoadConfig()
        {
            if (!File.Exists(ConfigPath))
            {
                SaveConfig();
                return;
            }

            try
            {
                var json = File.ReadAllText(ConfigPath);
                var loaded = JsonSerializer.Deserialize(json, ConfigJsonContext.Default.Config);
                if (loaded != null) _cfg = loaded;
            }
            catch
            {
                // Fallback on parse errors
                _cfg = Config.Default();
                SaveConfig();
            }
        }

        private static void SaveConfig()
        {
            var json = JsonSerializer.Serialize(_cfg, ConfigJsonContext.Default.Config);
            File.WriteAllText(ConfigPath, json);
        }

        // === MENU ===
        private static async Task MenuLoopAsync()
        {
            while (true)
            {
                Console.Clear();
                WriteHeader("Hearthstone Battlegrounds Skip");
                ShowStatus();
                Console.WriteLine();
                WriteColoredLine("Choose an option:", ConsoleColor.Cyan);
                WriteMenuOption("1", "Skip now");
                WriteMenuOption("2", "Set Hearthstone.exe path");
                WriteMenuOption("3", "Set disconnected seconds");
                WriteMenuOption("4", "Select OCR monitor");
                WriteMenuOption("5", "Select OCR capture area");
                WriteMenuOption("6", IsOcrCapturing ? "Stop OCR capture" : "Start OCR capture");
                WriteMenuOption("7", "Exit");
                Console.WriteLine();
                WriteColored("Enter choice: ", ConsoleColor.Cyan);

                var key = Console.ReadKey(intercept: true).KeyChar;
                Console.WriteLine();

                switch (key)
                {
                    case '1':
                        await EnsureAdminOrRelaunchAsync();
                        if (!await EnsureHearthstonePathAsync()) break;
                        await EnsureRuleExistsAsync();
                        await DoSkipAsync(_cfg.DisconnectedSeconds);
                        break;

                    case '2':
                        await PromptForHearthstonePathAsync();
                        break;

                    case '3':
                        PromptForSeconds();
                        break;

                    case '4':
                        PromptForOcrMonitor();
                        break;

                    case '5':
                        PromptForOcrArea();
                        break;

                    case '6':
                        ToggleOcrCapture();
                        break;

                    case '7':
                        StopOcrCapture();
                        SafeReconnect();
                        return;

                    default:
                        Pause("Invalid option.");
                        break;
                }
            }
        }

        // === STATUS / UX ===
        private static void ShowStatus()
        {
            var startRow = Console.CursorTop;
            var linesWritten = WriteStatusBlockContents();
            _statusBlockStartRow = startRow;
            _statusBlockLineCount = linesWritten;
        }

        private static int WriteStatusBlockContents()
        {
            var start = Console.CursorTop;

            Console.WriteLine();
            WriteSection("Config");
            WriteStatusValue("Rule Name", _cfg.OutboundRuleName, ConsoleColor.White);
            var path = _cfg.HearthstonePath ?? "(not set)";
            var pathColor = IsValidHearthstonePath(_cfg.HearthstonePath) ? ConsoleColor.Green : ConsoleColor.Red;
            WriteStatusValue("Hearthstone Path", path, pathColor);
            WriteStatusValue("Disconnected Seconds", _cfg.DisconnectedSeconds.ToString(), ConsoleColor.White);

            WriteSection("Firewall Rule");
            var exists = RuleExists();
            var enabled = exists && RuleEnabled();
            WriteStatusValue("Exists", exists ? "Yes" : "No", exists ? ConsoleColor.Green : ConsoleColor.Red);
            WriteStatusValue("Enabled", enabled ? "Yes" : "No", enabled ? ConsoleColor.Green : ConsoleColor.DarkGray);
            var isAdmin = IsAdmin();
            WriteStatusValue("Admin?", isAdmin ? "Yes" : "No", isAdmin ? ConsoleColor.Green : ConsoleColor.Red);

            WriteSection("OCR");
            var screens = Screen.AllScreens;
            var monitorColor = ConsoleColor.Red;
            string monitorText;
            if (screens.Length == 0)
            {
                monitorText = "(no monitors detected)";
            }
            else if (_cfg.OcrMonitorIndex is int monitorIndex && monitorIndex >= 0 && monitorIndex < screens.Length)
            {
                var screen = screens[monitorIndex];
                var primaryMark = screen.Primary ? " (Primary)" : string.Empty;
                monitorText = $"{monitorIndex + 1}: {screen.DeviceName} {screen.Bounds.Width}x{screen.Bounds.Height}{primaryMark}";
                monitorColor = ConsoleColor.Green;
            }
            else if (_cfg.OcrMonitorIndex.HasValue)
            {
                monitorText = "(invalid selection, please reconfigure)";
            }
            else
            {
                monitorText = "(not set)";
            }

            WriteStatusValue("Monitor", monitorText, monitorColor);

            ConsoleColor areaColor;
            string areaText;
            if (_cfg.OcrArea is { Width: > 0, Height: > 0 } area)
            {
                areaText = $"{area.Width}x{area.Height} @ ({area.X},{area.Y})";
                areaColor = ConsoleColor.Green;
            }
            else if (_cfg.OcrArea != null)
            {
                areaText = "(invalid area, please reselect)";
                areaColor = ConsoleColor.Red;
            }
            else
            {
                areaText = "(not set)";
                areaColor = ConsoleColor.Red;
            }

            WriteStatusValue("Area", areaText, areaColor);
            WriteStatusValue("Capturing", IsOcrCapturing ? "Yes" : "No", IsOcrCapturing ? ConsoleColor.Green : ConsoleColor.DarkGray);

            return Console.CursorTop - start;
        }

        private static void RefreshStatusBlock()
        {
            if (_statusBlockStartRow < 0 || _statusBlockLineCount <= 0)
                return;

            var savedLeft = Console.CursorLeft;
            var savedTop = Console.CursorTop;
            var width = Math.Max(1, Console.BufferWidth - 1);
            var linesToClear = _statusBlockLineCount;

            for (var i = 0; i < linesToClear; i++)
            {
                var targetRow = _statusBlockStartRow + i;
                if (targetRow < 0 || targetRow >= Console.BufferHeight)
                    continue;
                Console.SetCursorPosition(0, targetRow);
                Console.Write(new string(' ', width));
            }

            Console.SetCursorPosition(0, Math.Max(0, _statusBlockStartRow));
            var linesWritten = WriteStatusBlockContents();
            _statusBlockLineCount = linesWritten;

            Console.SetCursorPosition(Math.Max(0, savedLeft), Math.Max(0, savedTop));
        }

        private static void WriteHeader(string text)
        {
            WithColor(ConsoleColor.Cyan, () =>
            {
                Console.WriteLine(text);
                Console.WriteLine(new string('=', text.Length));
            });
        }

        private static void WriteSection(string text)
        {
            Console.WriteLine();
            WriteColoredLine($"[{text}]", ConsoleColor.Yellow);
        }

        private static void Pause(string? msg = null)
        {
            if (!string.IsNullOrWhiteSpace(msg))
                WriteColoredLine(msg, ConsoleColor.Yellow);
            WriteColored("Press any key to continue...", ConsoleColor.DarkGray);
            Console.ReadKey(true);
            Console.WriteLine();
        }

        // === SKIP FLOW ===
        private static async Task DoSkipAsync(int seconds)
        {
            if (seconds <= 0) seconds = 4;

            Console.WriteLine();
            WriteColoredLine("Disconnecting (enabling firewall rule) …", ConsoleColor.Yellow);
            await SetRuleEnabledAsync(true);
            RefreshStatusBlock();

            try
            {
                await CountdownAsync(seconds);
            }
            finally
            {
                ClearCountdownLine();
                await SetRuleEnabledAsync(false);
                RefreshStatusBlock();
            }
        }

        private static async Task CountdownAsync(int totalSeconds)
        {
            var start = DateTime.UtcNow;
            var end = start.AddSeconds(totalSeconds);

            while (DateTime.UtcNow < end)
            {
                var remaining = (int)Math.Ceiling((end - DateTime.UtcNow).TotalSeconds);
                DrawCountdownBar(totalSeconds, remaining);
                await Task.Delay(200);
            }
            DrawCountdownBar(totalSeconds, 0);
        }

        private static void DrawCountdownBar(int total, int remaining)
        {
            lock (ConsoleLock)
            {
                var width = Math.Max(20, Math.Min(Console.WindowWidth - 15, 60));
                var elapsed = Math.Clamp(total - remaining, 0, total);
                var filled = total == 0 ? width : (int)Math.Round(width * (elapsed / (double)total));
                var bar = new string('█', Math.Clamp(filled, 0, width)) + new string('░', Math.Clamp(width - filled, 0, width));
                var previousColor = Console.ForegroundColor;
                Console.CursorVisible = false;
                Console.ForegroundColor = ConsoleColor.Cyan;
                var line = $"[{bar}] ";
                Console.Write("\r" + line);
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write($"{remaining,2}s   ");
                Console.ForegroundColor = previousColor;
                Console.CursorVisible = true;
            }
        }

        private static void ClearCountdownLine()
        {
            lock (ConsoleLock)
            {
                var width = Math.Max(1, Console.WindowWidth - 1);
                Console.CursorVisible = false;
                Console.Write("\r" + new string(' ', width));
                Console.Write("\r");
                Console.CursorVisible = true;
            }
        }

        private static void SafeReconnect()
        {
            // Best-effort: disable the rule so connections resume if app exits during countdown
            try
            {
                if (RuleExists() && RuleEnabled())
                {
                    RunNetsh($"advfirewall firewall set rule name=\"{_cfg.OutboundRuleName}\" new enable=no", out _, out _);
                }
            }
            catch { /* ignore */ }
        }

        // === PATH SETUP ===
        private static async Task<bool> EnsureHearthstonePathAsync()
        {
            if (IsValidHearthstonePath(_cfg.HearthstonePath))
                return true;

            // try auto-detect
            var guessed = GuessHearthstonePath();
            if (IsValidHearthstonePath(guessed))
            {
                _cfg.HearthstonePath = guessed!;
                SaveConfig();
                return true;
            }

            WriteColoredLine("Hearthstone.exe not set or invalid.", ConsoleColor.Red);
            await PromptForHearthstonePathAsync();
            return IsValidHearthstonePath(_cfg.HearthstonePath);
        }

        private static async Task PromptForHearthstonePathAsync()
        {
            Console.WriteLine();
            WriteColoredLine("Enter full path to Hearthstone.exe (Tip: you can drag & drop the file into this window, leave empty to cancel):", ConsoleColor.Cyan);
            WriteColored("> ", ConsoleColor.Cyan);
            var rawInput = Console.ReadLine();
            var input = rawInput?.Trim('"', ' ');

            if (string.IsNullOrWhiteSpace(input))
            {
                // Offer auto-detect
                var guessed = GuessHearthstonePath();
                if (IsValidHearthstonePath(guessed))
                {
                    WriteStatusValue("Auto-detected", guessed!, ConsoleColor.Green);
                    _cfg.HearthstonePath = guessed!;
                    SaveConfig();
                    await Task.CompletedTask;
                    return;
                }

                WriteColoredLine("Auto-detect failed. Keeping current path.", ConsoleColor.Yellow);
                return;
            }

            if (input.Equals("cancel", StringComparison.OrdinalIgnoreCase))
            {
                WriteColoredLine("Cancelled. Keeping current path.", ConsoleColor.Yellow);
                await Task.CompletedTask;
                return;
            }

            if (!IsValidHearthstonePath(input))
            {
                WriteColoredLine("Invalid path. Make sure it points to Hearthstone.exe.", ConsoleColor.Red);
                return;
            }

            _cfg.HearthstonePath = Path.GetFullPath(input!);
            SaveConfig();
            WriteColoredLine("Saved.", ConsoleColor.Green);
            await Task.CompletedTask;
        }

        private static bool IsValidHearthstonePath(string? p)
            => !string.IsNullOrWhiteSpace(p)
               && File.Exists(p)
               && string.Equals(Path.GetFileName(p), "Hearthstone.exe", StringComparison.OrdinalIgnoreCase);

        private static string? GuessHearthstonePath()
        {
            var candidates = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Hearthstone", "Hearthstone.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Hearthstone", "Hearthstone.exe"),
                Path.Combine("C:\\Program Files (x86)\\Hearthstone", "Hearthstone.exe"),
                Path.Combine("C:\\Program Files\\Hearthstone", "Hearthstone.exe"),
            };

            foreach (var c in candidates)
                if (File.Exists(c)) return c;

            return null;
        }

        private static void PromptForSeconds()
        {
            Console.WriteLine();
            WriteColored($"Enter disconnect seconds (current: {_cfg.DisconnectedSeconds}, default 4, blank to cancel): ", ConsoleColor.Cyan);
            var s = Console.ReadLine()?.Trim();

            if (string.IsNullOrEmpty(s))
            {
                WriteColoredLine("Cancelled. Keeping current seconds.", ConsoleColor.Yellow);
                return;
            }

            if (s.Equals("cancel", StringComparison.OrdinalIgnoreCase))
            {
                WriteColoredLine("Cancelled. Keeping current seconds.", ConsoleColor.Yellow);
                return;
            }

            if (int.TryParse(s, out var val) && val >= 1 && val <= 60)
            {
                _cfg.DisconnectedSeconds = val;
                SaveConfig();
                WriteColoredLine("Saved.", ConsoleColor.Green);
            }
            else
            {
                WriteColoredLine("Not a valid number between 1 and 60.", ConsoleColor.Red);
            }
        }

        // === OCR CAPTURE ===
        private static void PromptForOcrMonitor()
        {
            Console.WriteLine();
            var screens = Screen.AllScreens;
            if (screens.Length == 0)
            {
                WriteColoredLine("No monitors detected for OCR capture.", ConsoleColor.Red);
                Pause();
                return;
            }

            WriteColoredLine("Available monitors:", ConsoleColor.Cyan);
            for (var i = 0; i < screens.Length; i++)
            {
                var screen = screens[i];
                var primaryMark = screen.Primary ? " (Primary)" : string.Empty;
                WriteColoredLine($"  {i + 1}) {screen.DeviceName} {screen.Bounds.Width}x{screen.Bounds.Height}{primaryMark}", ConsoleColor.White);
            }

            WriteColored("Enter monitor number (blank to cancel): ", ConsoleColor.Cyan);
            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(input))
            {
                WriteColoredLine("Cancelled. Keeping current monitor.", ConsoleColor.Yellow);
                Pause();
                return;
            }

            if (!int.TryParse(input, out var choice) || choice < 1 || choice > screens.Length)
            {
                WriteColoredLine("Invalid monitor selection.", ConsoleColor.Red);
                Pause();
                return;
            }

            _cfg.OcrMonitorIndex = choice - 1;
            SaveConfig();
            RefreshStatusBlock();
            WriteColoredLine("Saved monitor selection.", ConsoleColor.Green);
            Pause();
        }

        private static void PromptForOcrArea()
        {
            Console.WriteLine();
            if (!TryGetConfiguredScreen(out var screen, showError: true))
            {
                Pause();
                return;
            }

            WriteColoredLine("An overlay will appear on the selected monitor.", ConsoleColor.Cyan);
            WriteColoredLine("Click and drag to choose the area. Press Esc or right-click to cancel.", ConsoleColor.Cyan);

            try
            {
                var selection = CaptureAreaSelector.SelectArea(screen);
                if (selection is null)
                {
                    WriteColoredLine("Selection cancelled.", ConsoleColor.Yellow);
                    Pause();
                    return;
                }

                _cfg.OcrArea = CaptureAreaConfig.FromRectangle(selection.Value);
                SaveConfig();
                RefreshStatusBlock();
                WriteColoredLine($"Saved area {selection.Value.Width}x{selection.Value.Height} @ ({selection.Value.X},{selection.Value.Y}).", ConsoleColor.Green);
            }
            catch (Exception ex)
            {
                WriteColoredLine($"Failed to capture area: {ex.Message}", ConsoleColor.Red);
            }

            Pause();
        }

        private static void ToggleOcrCapture()
        {
            Console.WriteLine();
            if (IsOcrCapturing)
            {
                StopOcrCapture();
                RefreshStatusBlock();
                WriteColoredLine("OCR capture stopped.", ConsoleColor.Yellow);
                Pause();
                return;
            }

            if (!TryGetConfiguredScreen(out var screen, showError: true) || !TryGetOcrArea(screen, out var area, showError: true))
            {
                Pause();
                return;
            }

            StartOcrCapture(area);
            RefreshStatusBlock();
            WriteColoredLine("OCR capture started. Use the menu to stop.", ConsoleColor.Green);
            Pause();
        }

        private static bool TryGetConfiguredScreen(out Screen screen, bool showError)
        {
            var screens = Screen.AllScreens;
            if (screens.Length == 0)
            {
                screen = Screen.PrimaryScreen!;
                if (showError)
                    WriteColoredLine("No monitors detected.", ConsoleColor.Red);
                return false;
            }

            if (_cfg.OcrMonitorIndex is int index && index >= 0 && index < screens.Length)
            {
                screen = screens[index];
                return true;
            }

            screen = Screen.PrimaryScreen ?? screens[0];
            if (showError)
                WriteColoredLine("OCR monitor not configured. Please select a monitor first.", ConsoleColor.Red);
            return false;
        }

        private static bool TryGetOcrArea(Screen screen, out Rectangle area, bool showError)
        {
            if (_cfg.OcrArea is { Width: > 0, Height: > 0 } stored)
            {
                var rect = stored.ToRectangle();
                if (screen.Bounds.Contains(rect))
                {
                    area = rect;
                    return true;
                }

                if (showError)
                    WriteColoredLine("The configured area no longer fits on the selected monitor. Please reselect it.", ConsoleColor.Red);
            }
            else if (showError)
            {
                WriteColoredLine("OCR capture area not configured.", ConsoleColor.Red);
            }

            area = default;
            return false;
        }

        private static void StartOcrCapture(Rectangle area)
        {
            StopOcrCapture();
            _ocrCts = new CancellationTokenSource();
            _lastOcrText = null;

            var token = _ocrCts.Token;
            var thread = new Thread(() => RunOcrCaptureThread(area, token))
            {
                IsBackground = true,
                Name = "OCR Capture"
            };
            thread.SetApartmentState(ApartmentState.STA);
            _ocrThread = thread;
            thread.Start();
        }

        private static void StopOcrCapture()
        {
            var cts = Interlocked.Exchange(ref _ocrCts, null);
            if (cts == null)
                return;

            try
            {
                cts.Cancel();
                if (_ocrThread is { IsAlive: true } thread)
                {
                    if (!thread.Join(TimeSpan.FromSeconds(5)))
                    {
                        thread.Interrupt();
                        thread.Join(TimeSpan.FromSeconds(2));
                    }
                }
            }
            catch
            {
                // ignore cleanup errors
            }
            finally
            {
                _ocrThread = null;
                cts.Dispose();
            }
        }

        private static void RunOcrCaptureThread(Rectangle area, CancellationToken token)
        {
            CaptureIndicatorOverlay? overlay = null;
            try
            {
                Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                overlay = new CaptureIndicatorOverlay(area);
                overlay.Show();
                Application.DoEvents();

                var engine = CreateOcrEngine();
                if (engine == null)
                {
                    return;
                }

                while (!token.IsCancellationRequested)
                {
                    Application.DoEvents();
                    string? text = null;
                    try
                    {
                        text = CaptureAndRecognize(area, engine);
                    }
                    catch (Exception ex)
                    {
                        lock (ConsoleLock)
                        {
                            WriteColoredLine($"OCR capture error: {ex.Message}", ConsoleColor.Red);
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(text) && !string.Equals(text, _lastOcrText, StringComparison.Ordinal))
                    {
                        _lastOcrText = text;
                        lock (ConsoleLock)
                        {
                            Console.WriteLine();
                            WriteColoredLine($"[OCR {DateTime.Now:HH:mm:ss}]", ConsoleColor.Magenta);
                            Console.WriteLine(text);
                            Console.WriteLine();
                        }
                    }
                    for (var i = 0; i < 10 && !token.IsCancellationRequested; i++)
                    {
                        Application.DoEvents();
                        Thread.Sleep(100);
                    }
                }
            }
            catch (ThreadInterruptedException)
            {
                // graceful shutdown
            }
            catch (Exception ex)
            {
                lock (ConsoleLock)
                {
                    WriteColoredLine($"OCR capture stopped: {ex.Message}", ConsoleColor.Red);
                }
            }
            finally
            {
                overlay?.Close();
                overlay?.Dispose();
                if (ReferenceEquals(_ocrThread, Thread.CurrentThread))
                {
                    _ocrThread = null;
                    var cts = Interlocked.Exchange(ref _ocrCts, null);
                    cts?.Dispose();
                }
            }
        }

        private static OcrEngine? CreateOcrEngine()
        {
            try
            {
                var engine = OcrEngine.TryCreateFromUserProfileLanguages();
                if (engine != null)
                    return engine;

                var fallback = new Language("en");
                return OcrEngine.TryCreateFromLanguage(fallback);
            }
            catch (Exception ex)
            {
                lock (ConsoleLock)
                {
                    WriteColoredLine($"Failed to initialize OCR engine: {ex.Message}", ConsoleColor.Red);
                }
                return null;
            }
        }

        private static string? CaptureAndRecognize(Rectangle area, OcrEngine engine)
        {
            using var bitmap = new Bitmap(area.Width, area.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(area.Location, Point.Empty, area.Size, CopyPixelOperation.SourceCopy);
            }

            using var memory = new MemoryStream();
            bitmap.Save(memory, ImageFormat.Png);
            memory.Position = 0;

            using var stream = new InMemoryRandomAccessStream();
            using (var writer = new DataWriter(stream))
            {
                writer.WriteBytes(memory.ToArray());
                writer.StoreAsync().AsTask().GetAwaiter().GetResult();
                writer.DetachStream();
            }

            stream.Seek(0);
            var decoder = BitmapDecoder.CreateAsync(stream).AsTask().GetAwaiter().GetResult();
            using var softwareBitmap = decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied).AsTask().GetAwaiter().GetResult();
            var result = engine.RecognizeAsync(softwareBitmap).AsTask().GetAwaiter().GetResult();
            var text = result.Text?.Trim();
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private static class CaptureAreaSelector
        {
            public static Rectangle? SelectArea(Screen screen)
            {
                var tcs = new TaskCompletionSource<Rectangle?>();
                var thread = new Thread(() => RunSelection(screen, tcs))
                {
                    IsBackground = true,
                    Name = "OCR Area Selector"
                };
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();

                try
                {
                    var result = tcs.Task.GetAwaiter().GetResult();
                    thread.Join();
                    return result;
                }
                catch
                {
                    thread.Join();
                    throw;
                }
            }

            private static void RunSelection(Screen screen, TaskCompletionSource<Rectangle?> completion)
            {
                try
                {
                    Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    using var overlay = new SelectionOverlay(screen.Bounds);
                    overlay.FormClosed += (_, __) => completion.TrySetResult(overlay.SelectedArea);
                    Application.Run(overlay);
                }
                catch (Exception ex)
                {
                    completion.TrySetException(ex);
                }
            }
        }

        private sealed class SelectionOverlay : Form
        {
            private readonly Rectangle _screenBounds;
            private Point _startPoint;
            private Rectangle? _currentSelection;
            private bool _dragging;

            public SelectionOverlay(Rectangle bounds)
            {
                _screenBounds = bounds;
                DoubleBuffered = true;
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                ShowInTaskbar = false;
                TopMost = true;
                Bounds = bounds;
                Location = bounds.Location;
                Cursor = Cursors.Cross;
                BackColor = Color.Black;
                Opacity = 0.2;
                KeyPreview = true;
            }

            public Rectangle? SelectedArea { get; private set; }

            protected override void OnKeyDown(KeyEventArgs e)
            {
                base.OnKeyDown(e);
                if (e.KeyCode == Keys.Escape)
                {
                    SelectedArea = null;
                    Close();
                }
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                base.OnMouseDown(e);
                if (e.Button != MouseButtons.Left)
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        SelectedArea = null;
                        Close();
                    }
                    return;
                }

                _dragging = true;
                _startPoint = e.Location;
                _currentSelection = new Rectangle(e.Location, Size.Empty);
                Invalidate();
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);
                if (!_dragging)
                    return;

                _currentSelection = Normalize(_startPoint, e.Location);
                Invalidate();
            }

            protected override void OnMouseUp(MouseEventArgs e)
            {
                base.OnMouseUp(e);
                if (!_dragging || e.Button != MouseButtons.Left)
                    return;

                _dragging = false;
                _currentSelection = Normalize(_startPoint, e.Location);
                if (_currentSelection is { Width: > 0, Height: > 0 } selection)
                {
                    SelectedArea = new Rectangle(selection.X + _screenBounds.X, selection.Y + _screenBounds.Y, selection.Width, selection.Height);
                    Close();
                }
                else
                {
                    SelectedArea = null;
                    Close();
                }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                if (_currentSelection is not { Width: > 0, Height: > 0 } selection)
                    return;

                using var fillBrush = new SolidBrush(Color.FromArgb(80, Color.DeepSkyBlue));
                using var pen = new Pen(Color.Cyan, 2);
                e.Graphics.FillRectangle(fillBrush, selection);
                e.Graphics.DrawRectangle(pen, selection);
            }

            private static Rectangle Normalize(Point start, Point end)
            {
                var x = Math.Min(start.X, end.X);
                var y = Math.Min(start.Y, end.Y);
                var width = Math.Abs(start.X - end.X);
                var height = Math.Abs(start.Y - end.Y);
                return new Rectangle(x, y, width, height);
            }
        }

        private sealed class CaptureIndicatorOverlay : Form
        {
            public CaptureIndicatorOverlay(Rectangle bounds)
            {
                DoubleBuffered = true;
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                ShowInTaskbar = false;
                TopMost = true;
                Bounds = bounds;
                Location = bounds.Location;
                BackColor = Color.Lime;
                TransparencyKey = Color.Lime;
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    var cp = base.CreateParams;
                    cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW
                    cp.ExStyle |= 0x00080000; // WS_EX_LAYERED
                    cp.ExStyle |= 0x00000020; // WS_EX_TRANSPARENT
                    return cp;
                }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using var pen = new Pen(Color.Red, 3);
                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.DrawRectangle(pen, rect);
            }
        }

        // === FIREWALL RULES ===
        private static async Task EnsureRuleExistsAsync()
        {
            if (RuleExists())
                return;

            await CreateRuleAsync();
        }

        private static async Task CreateRuleAsync()
        {
            // Create disabled rule that blocks outbound for Hearthstone.exe
            var cmd = $"advfirewall firewall add rule name=\"{_cfg.OutboundRuleName}\" dir=out action=block program=\"{_cfg.HearthstonePath}\" enable=no";
            var ok = RunNetsh(cmd, out var so, out var se);
            if (!ok)
                throw new InvalidOperationException($"Failed to create rule.\n{se}\n{so}");
            await Task.CompletedTask;
        }

        private static bool RuleExists()
        {
            var cmd = $"advfirewall firewall show rule name=\"{_cfg.OutboundRuleName}\"";
            var ok = RunNetsh(cmd, out var so, out _);
            if (!ok) return false;

            var output = so ?? string.Empty;
            var lower = output.ToLowerInvariant();

            if (lower.Contains("no rules match") || lower.Contains("keine regeln"))
                return false;

            if (lower.Contains("rule name") || lower.Contains("regelname"))
                return true;

            // Fallback: any non-empty output without an explicit negative usually means the rule exists
            return !string.IsNullOrWhiteSpace(output);
        }

        private static bool RuleEnabled()
        {
            var cmd = $"advfirewall firewall show rule name=\"{_cfg.OutboundRuleName}\"";
            var ok = RunNetsh(cmd, out var so, out _);
            if (!ok) return false;

            foreach (var rawLine in so.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var line = rawLine.Trim();
                var lower = line.ToLowerInvariant();

                if (lower.StartsWith("enabled") || lower.StartsWith("enabled:") || lower.StartsWith("aktiviert"))
                {
                    if (lower.Contains("yes") || lower.Contains("ja")) return true;
                    if (lower.Contains("no") || lower.Contains("nein")) return false;
                }
            }

            return false;
        }

        private static async Task SetRuleEnabledAsync(bool enabled)
        {
            var state = enabled ? "yes" : "no";
            var cmd = $"advfirewall firewall set rule name=\"{_cfg.OutboundRuleName}\" new enable={state}";
            var ok = RunNetsh(cmd, out _, out var se);
            if (!ok)
            {
                WriteColoredLine("Failed to set rule state:", ConsoleColor.Red);
                WriteColoredLine(se, ConsoleColor.Red);
            }
            await Task.CompletedTask;
        }

        private static bool RunNetsh(string args, out string stdOut, out string stdErr)
        {
            using var p = new Process();
            p.StartInfo = new ProcessStartInfo
            {
                FileName = "netsh",
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            try
            {
                p.Start();
                stdOut = p.StandardOutput.ReadToEnd();
                stdErr = p.StandardError.ReadToEnd();
                p.WaitForExit();
                return p.ExitCode == 0;
            }
            catch (Exception ex)
            {
                stdOut = "";
                stdErr = ex.Message;
                return false;
            }
        }

        // === ELEVATION ===
        private static bool IsAdmin()
        {
            try
            {
                using var id = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(id);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private static async Task EnsureAdminOrRelaunchAsync()
        {
            if (IsAdmin()) return;

            Console.WriteLine();
            WriteColoredLine("⚠ This action requires Administrator privileges (for netsh).", ConsoleColor.Yellow);
            WriteColored("Relaunch elevated now? [Y/n]: ", ConsoleColor.Cyan);
            var key = Console.ReadKey(true).Key;
            Console.WriteLine();

            if (key == ConsoleKey.N)
            {
                WriteColoredLine("Continuing without elevation may fail.", ConsoleColor.Red);
                return;
            }

            RelaunchElevated();
            // If relaunch succeeds, current process exits.
            await Task.Delay(100);
        }

        private static void RelaunchElevated()
        {
            var exe = Environment.ProcessPath!;
            var psi = new ProcessStartInfo(exe)
            {
                UseShellExecute = true,
                Verb = "runas" // triggers UAC
            };

            try
            {
                Process.Start(psi);
                Environment.Exit(0);
            }
            catch
            {
                WriteColoredLine("Elevation was canceled.", ConsoleColor.Red);
            }
        }

        [JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
        [JsonSerializable(typeof(Config))]
        [JsonSerializable(typeof(CaptureAreaConfig))]
        private partial class ConfigJsonContext : JsonSerializerContext
        {
        }

        private sealed class CaptureAreaConfig
        {
            public int X { get; set; }
            public int Y { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }

            public Rectangle ToRectangle() => new(X, Y, Width, Height);

            public static CaptureAreaConfig FromRectangle(Rectangle rect) => new()
            {
                X = rect.X,
                Y = rect.Y,
                Width = rect.Width,
                Height = rect.Height
            };
        }

        // === CONFIG MODEL ===
        private sealed class Config
        {
            public string OutboundRuleName { get; set; } = "HS Connection Blocker";
            public string? HearthstonePath { get; set; } = null;
            public int DisconnectedSeconds { get; set; } = 4;
            public int? OcrMonitorIndex { get; set; }
            public CaptureAreaConfig? OcrArea { get; set; }

            public static Config Default() => new();
        }
    }
}
