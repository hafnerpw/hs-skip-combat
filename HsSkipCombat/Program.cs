using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

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

        private static async Task<int> Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.Title = "Hearthstone Battlegrounds Skip (Console)";

            EnsureAppFolder();
            LoadConfig();

            AppDomain.CurrentDomain.ProcessExit += (_, __) => SafeReconnect();
            Console.CancelKeyPress += (_, e) =>
            {
                e.Cancel = true;
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
                WriteMenuOption("4", "Exit");
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
        private partial class ConfigJsonContext : JsonSerializerContext
        {
        }

        // === CONFIG MODEL ===
        private sealed class Config
        {
            public string OutboundRuleName { get; set; } = "HS Connection Blocker";
            public string? HearthstonePath { get; set; } = null;
            public int DisconnectedSeconds { get; set; } = 4;

            public static Config Default() => new();
        }
    }
}
