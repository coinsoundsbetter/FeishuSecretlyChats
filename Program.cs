using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;

namespace FeishuStealthChat
{
    // ============== 轻量日志 ==============
    internal static class Log
    {
        private static readonly object _lock = new object();
        private static string _logDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FeishuStealthChat");
        private static string _logPath = Path.Combine(_logDir, "app.log");

        public static string LogPath => _logPath;

        public static void Init(string? customDir = null)
        {
            try
            {
                if (!string.IsNullOrEmpty(customDir))
                {
                    _logDir = customDir!;
                    _logPath = Path.Combine(_logDir, "app.log");
                }
                Directory.CreateDirectory(_logDir);
                Write("==== App start ====");
            }
            catch { /* ignore */ }
        }

        public static void Write(string msg)
        {
            try
            {
                lock (_lock)
                {
                    File.AppendAllText(_logPath,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}\r\n",
                        Encoding.UTF8);
                }
            }
            catch { /* ignore */ }
        }

        public static void Error(string msg, Exception ex)
        {
            Write($"{msg} EX={ex.GetType().Name}: {ex.Message}");
        }

        public static void Clear()
        {
            try
            {
                lock (_lock)
                {
                    if (File.Exists(_logPath)) File.Delete(_logPath);
                }
                Write("==== Log cleared ====");
            }
            catch { /* ignore */ }
        }
    }

    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Log.Init(); // 日志初始化
            Application.Run(new TrayApp());
        }
    }

    // ================= 配置模型 =================
    public class AppConfig
    {
        // 主功能热键（执行：文字->图片->粘贴->延迟回车）
        public HotkeySetting Hotkey { get; set; } = new HotkeySetting { Key = "F9" };

        // 启用/禁用 切换快捷键（组合键；再次按下就反向）
        public HotkeySetting ToggleHotkey { get; set; } = new HotkeySetting { Ctrl = true, Shift = true, Key = "E" }; // Ctrl+Shift+E

        public int SendDelayMs { get; set; } = 3000; // 粘贴->自动回车的延迟（毫秒）
        public bool Enabled { get; set; } = true;    // 偷聊模式开关
    }

    public class HotkeySetting
    {
        public string Key { get; set; } = "None"; // "None"/空 = 不注册
        public bool Ctrl { get; set; } = false;
        public bool Alt { get; set; } = false;
        public bool Shift { get; set; } = false;
        public bool Win { get; set; } = false;

        public bool IsConfigured()
        {
            return !string.IsNullOrWhiteSpace(Key) && !Key.Equals("None", StringComparison.OrdinalIgnoreCase);
        }

        public (uint mod, Keys key) ToRegisterArgs()
        {
            uint mod = 0;
            if (Ctrl) mod |= HotkeyWindow.MOD_CONTROL;
            if (Alt) mod |= HotkeyWindow.MOD_ALT;
            if (Shift) mod |= HotkeyWindow.MOD_SHIFT;
            if (Win) mod |= HotkeyWindow.MOD_WIN;

            if (!Enum.TryParse(Key, true, out Keys parsed) || parsed == Keys.None)
                parsed = Keys.None;

            parsed &= Keys.KeyCode;
            return (mod, parsed);
        }

        public override string ToString()
        {
            string mods = "";
            if (Ctrl) mods += "Ctrl+";
            if (Alt) mods += "Alt+";
            if (Shift) mods += "Shift+";
            if (Win) mods += "Win+";
            return string.IsNullOrWhiteSpace(Key) ? "(未配置)" : $"{mods}{Key}";
        }
    }

    // ================= 托盘程序 =================
    internal class TrayApp : ApplicationContext
    {
        private NotifyIcon trayIcon;
        private HotkeyWindow hotkeyWnd;
        private AppConfig config;
        private readonly string configPath;

        // 重入保护
        private volatile bool _inProgress = false;

        // 多热键ID
        private const int HOTKEY_ID_ACTION = 0xA001; // 主功能
        private const int HOTKEY_ID_TOGGLE = 0xA004; // 启用/禁用 切换

        public TrayApp()
        {
            configPath = EnsureConfigFile();
            config = LoadConfig(configPath);
            Log.Write($"Config loaded: Enabled={config.Enabled}, Hotkey={config.Hotkey}, Toggle={config.ToggleHotkey}, DelayMs={config.SendDelayMs}, Path={configPath}");

            trayIcon = new NotifyIcon()
            {
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
                Visible = true,
                Text = config.Enabled ? "飞书偷聊王（已启用）" : "飞书偷聊王（已禁用）"
            };

            trayIcon.ContextMenuStrip = BuildContextMenu();

            hotkeyWnd = new HotkeyWindow();
            hotkeyWnd.HotkeyPressed += id =>
            {
                Log.Write($"Hotkey pressed: id=0x{id:X}");
                if (!IsFeishuForeground(out string? fgInfo))
                {
                    Log.Write($"Ignored: foreground is not Feishu. FG={fgInfo}");
                    return;
                }
                Log.Write($"Foreground OK: {fgInfo}");

                if (id == HOTKEY_ID_TOGGLE)
                {
                    SetEnabled(!config.Enabled);
                    return;
                }

                if (id == HOTKEY_ID_ACTION && config.Enabled && IsFeishuForeground(out fgInfo))
                {
                    ConvertTextToImageAndSend();
                }
            };

            RegisterAllHotkeys();

            trayIcon.ShowBalloonTip(
                2500,
                "开始偷聊吧！🤭",
                $"已启动\n模式：{(config.Enabled ? "启用" : "禁用")}\n主热键：{config.Hotkey}\n切换快捷键：{config.ToggleHotkey}\n延迟：{config.SendDelayMs} ms\n配置：{configPath}",
                ToolTipIcon.None
            );
        }

        private ContextMenuStrip BuildContextMenu()
        {
            var menu = new ContextMenuStrip();

            // 启用/禁用
            var enableItem = new ToolStripMenuItem("启用偷聊模式")
            {
                Checked = config.Enabled,
                CheckOnClick = true
            };
            enableItem.CheckedChanged += (_, __) => SetEnabled(enableItem.Checked);
            menu.Items.Add(enableItem);

            // 发送延迟
            var delayMenu = new ToolStripMenuItem("发送延迟");
            void AddDelayItem(string text, int ms)
            {
                var item = new ToolStripMenuItem(text) { Checked = (config.SendDelayMs == ms) };
                item.Click += (_, __) =>
                {
                    config.SendDelayMs = ms;
                    SaveConfig();
                    UncheckAllMenuItems(delayMenu.DropDownItems);
                    item.Checked = true;
                    Log.Write($"Delay changed to {ms}ms");
                    trayIcon.ShowBalloonTip(1200, "延迟已更新", $"{ms} ms", ToolTipIcon.None);
                };
                delayMenu.DropDownItems.Add(item);
            }
            AddDelayItem("0.5 秒", 500);
            AddDelayItem("1 秒", 1000);
            AddDelayItem("2 秒", 2000);
            AddDelayItem("3 秒", 3000);
            AddDelayItem("5 秒", 5000);

            var customDelay = new ToolStripMenuItem("自定义…");
            customDelay.Click += (_, __) =>
            {
                if (PromptDelayMs(config.SendDelayMs, out int ms))
                {
                    config.SendDelayMs = Math.Max(0, ms);
                    SaveConfig();
                    UncheckAllMenuItems(delayMenu.DropDownItems);
                    Log.Write($"Delay changed to {ms}ms (custom)");
                    trayIcon.ShowBalloonTip(1200, "延迟已更新", $"{ms} ms", ToolTipIcon.None);
                }
            };
            delayMenu.DropDownItems.Add(new ToolStripSeparator());
            delayMenu.DropDownItems.Add(customDelay);
            menu.Items.Add(delayMenu);

            // 主快捷键
            var hotkeyMenu = new ToolStripMenuItem("主快捷键");
            void AddHotkeyItem(string keyName)
            {
                var item = new ToolStripMenuItem(keyName)
                {
                    Checked = config.Hotkey.Key.Equals(keyName, StringComparison.OrdinalIgnoreCase)
                };
                item.Click += (_, __) =>
                {
                    config.Hotkey.Key = keyName;
                    config.Hotkey.Ctrl = config.Hotkey.Alt = config.Hotkey.Shift = config.Hotkey.Win = false;
                    SaveConfig();
                    RegisterAllHotkeys();

                    UncheckAllMenuItems(hotkeyMenu.DropDownItems);
                    item.Checked = true;

                    Log.Write($"Main hotkey changed to {keyName}");
                    trayIcon.ShowBalloonTip(1200, "主快捷键已更新", keyName, ToolTipIcon.None);
                };
                hotkeyMenu.DropDownItems.Add(item);
            }
            AddHotkeyItem("F9");
            AddHotkeyItem("F10");
            AddHotkeyItem("F11");
            AddHotkeyItem("F12");

            var customHotkey = new ToolStripMenuItem("自定义…");
            customHotkey.Click += (_, __) =>
            {
                if (PromptHotkey(config.Hotkey.Key, out string keyName))
                {
                    config.Hotkey.Key = keyName.Trim();
                    config.Hotkey.Ctrl = config.Hotkey.Alt = config.Hotkey.Shift = config.Hotkey.Win = false;
                    SaveConfig();
                    RegisterAllHotkeys();
                    UncheckAllMenuItems(hotkeyMenu.DropDownItems);

                    Log.Write($"Main hotkey changed to {keyName} (custom)");
                    trayIcon.ShowBalloonTip(1200, "主快捷键已更新", keyName, ToolTipIcon.None);
                }
            };
            hotkeyMenu.DropDownItems.Add(new ToolStripSeparator());
            hotkeyMenu.DropDownItems.Add(customHotkey);
            menu.Items.Add(hotkeyMenu);

            // 切换快捷键
            var toggleHKMenu = new ToolStripMenuItem("切换快捷键（启用↔禁用）");
            void AddTogglePreset(string text, bool ctrl, bool alt, bool shift, bool win, string key)
            {
                var item = new ToolStripMenuItem(text)
                {
                    Checked = (config.ToggleHotkey.Ctrl == ctrl &&
                               config.ToggleHotkey.Alt == alt &&
                               config.ToggleHotkey.Shift == shift &&
                               config.ToggleHotkey.Win == win &&
                               config.ToggleHotkey.Key.Equals(key, StringComparison.OrdinalIgnoreCase))
                };
                item.Click += (_, __) =>
                {
                    config.ToggleHotkey.Ctrl = ctrl;
                    config.ToggleHotkey.Alt = alt;
                    config.ToggleHotkey.Shift = shift;
                    config.ToggleHotkey.Win = win;
                    config.ToggleHotkey.Key = key;
                    SaveConfig();
                    RegisterAllHotkeys();

                    UncheckAllMenuItems(toggleHKMenu.DropDownItems);
                    item.Checked = true;

                    Log.Write($"Toggle hotkey changed to {text}");
                    trayIcon.ShowBalloonTip(1200, "切换快捷键已更新", $"{text}", ToolTipIcon.Info);
                };
                toggleHKMenu.DropDownItems.Add(item);
            }
            AddTogglePreset("Ctrl+Shift+E", true, false, true, false, "E");
            AddTogglePreset("Ctrl+Alt+E", true, true, false, false, "E");
            AddTogglePreset("Alt+Shift+E", false, true, true, false, "E");

            var customToggleHK = new ToolStripMenuItem("自定义…");
            customToggleHK.Click += (_, __) =>
            {
                if (PromptComboHotkey(config.ToggleHotkey, out var hk))
                {
                    config.ToggleHotkey = hk;
                    SaveConfig();
                    RegisterAllHotkeys();
                    UncheckAllMenuItems(toggleHKMenu.DropDownItems);

                    Log.Write($"Toggle hotkey changed to {hk}");
                    trayIcon.ShowBalloonTip(1200, "切换快捷键已更新", hk.ToString(), ToolTipIcon.Info);
                }
            };
            toggleHKMenu.DropDownItems.Add(new ToolStripSeparator());
            toggleHKMenu.DropDownItems.Add(customToggleHK);
            menu.Items.Add(toggleHKMenu);

            menu.Items.Add(new ToolStripSeparator());

            // 日志菜单
            var miOpenLog = new ToolStripMenuItem("打开日志文件", null, (_, __) =>
            {
                try { Process.Start("notepad.exe", Log.LogPath); } catch { }
            });
            var miClearLog = new ToolStripMenuItem("清空日志", null, (_, __) =>
            {
                Log.Clear();
                trayIcon.ShowBalloonTip(1200, "日志", "日志已清空", ToolTipIcon.Info);
            });

            var miReload = new ToolStripMenuItem("重载配置", null, (_, __) =>
            {
                config = LoadConfig(configPath);
                RegisterAllHotkeys();
                trayIcon.Text = config.Enabled ? "Feishu Stealth Chat（已启用）" : "Feishu Stealth Chat（已禁用）";
                Log.Write("Config reloaded");
                trayIcon.ShowBalloonTip(1500, "配置已重载",
                    $"模式：{(config.Enabled ? "启用" : "禁用")}\n主热键：{config.Hotkey}\n切换快捷键：{config.ToggleHotkey}\n延迟：{config.SendDelayMs} ms",
                    ToolTipIcon.Info);
            });

            var miExit = new ToolStripMenuItem("退出", null, (_, __) =>
            {
                Log.Write("App exit");
                trayIcon.Visible = false;
                hotkeyWnd.Dispose();
                Application.Exit();
            });

            menu.Items.Add(miOpenLog);
            menu.Items.Add(miClearLog);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(miReload);
            menu.Items.Add(miExit);

            return menu;
        }

        private void RegisterAllHotkeys()
        {
            // 先注销，避免冲突
            hotkeyWnd.TryUnregister(HOTKEY_ID_ACTION);
            hotkeyWnd.TryUnregister(HOTKEY_ID_TOGGLE);

            // 主功能
            if (config.Hotkey.IsConfigured())
            {
                var (mod, key) = config.Hotkey.ToRegisterArgs();
                if (key != Keys.None)
                {
                    hotkeyWnd.Register(HOTKEY_ID_ACTION, mod, key);
                    Log.Write($"Registered main hotkey: id=0x{HOTKEY_ID_ACTION:X}, mod={mod}, key={key}");
                }
            }

            // 切换启用/禁用（组合键）
            if (config.ToggleHotkey.IsConfigured())
            {
                var (mod, key) = config.ToggleHotkey.ToRegisterArgs();
                if (key != Keys.None)
                {
                    hotkeyWnd.Register(HOTKEY_ID_TOGGLE, mod, key);
                    Log.Write($"Registered toggle hotkey: id=0x{HOTKEY_ID_TOGGLE:X}, mod={mod}, key={key}");
                }
            }
        }

        private void SetEnabled(bool enabled)
        {
            config.Enabled = enabled;
            SaveConfig();
            trayIcon.Text = enabled ? "Feishu Stealth Chat（已启用）" : "Feishu Stealth Chat（已禁用）";
            Log.Write($"Mode changed: Enabled={enabled}");
            trayIcon.ShowBalloonTip(2000, "模式切换", enabled ? "偷聊模式已启用" : "偷聊模式已禁用", ToolTipIcon.Info);
        }

        private static void UncheckAllMenuItems(ToolStripItemCollection items)
        {
            foreach (ToolStripItem it in items)
                if (it is ToolStripMenuItem mi)
                    mi.Checked = false;
        }

        private string EnsureConfigFile()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "FeishuStealthChat");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "config.json");
            if (!File.Exists(path))
            {
                var def = new AppConfig();
                var json = JsonSerializer.Serialize(def, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(path, json);
                Log.Write($"Config created at {path}");
            }
            return path;
        }

        private AppConfig LoadConfig(string path)
        {
            try
            {
                var json = File.ReadAllText(path);
                var cfg = JsonSerializer.Deserialize<AppConfig>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new AppConfig();

                if (cfg.SendDelayMs < 0) cfg.SendDelayMs = 0;
                return cfg;
            }
            catch (Exception ex)
            {
                Log.Error("LoadConfig failed, write default", ex);
                var def = new AppConfig();
                var json = JsonSerializer.Serialize(def, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(path, json);
                return def;
            }
        }

        private void SaveConfig()
        {
            try
            {
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(configPath, json);
                Log.Write("Config saved");
            }
            catch (Exception ex)
            {
                Log.Error("SaveConfig failed", ex);
                trayIcon.ShowBalloonTip(2500, "保存配置失败", ex.Message, ToolTipIcon.Error);
            }
        }

        private bool IsFeishuForeground()
        {
            return IsFeishuForeground(out _);
        }

        private bool IsFeishuForeground(out string? info)
        {
            info = null;
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) { info = "NoFGWindow"; return false; }

            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            try
            {
                var p = Process.GetProcessById((int)pid);
                info = $"{p.ProcessName} (PID {p.Id})";
                return p.ProcessName.Equals("Lark", StringComparison.OrdinalIgnoreCase)
                    || p.ProcessName.Equals("Feishu", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Log.Error("IsFeishuForeground get process failed", ex);
                return false;
            }
        }

        // ====== 新增：剪贴板序列号 API 与等待 ======
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();

        private bool WaitClipboardSeqChanged(uint seqBefore, int timeoutMs = 600, int pollMs = 20)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    if (GetClipboardSequenceNumber() != seqBefore) return true;
                }
                catch { /* ignore */ }
                Thread.Sleep(pollMs);
                Application.DoEvents();
            }
            return false;
        }

        // ============ 性能优化 + 序列号校验的 ConvertTextToImageAndSend ============
        private void ConvertTextToImageAndSend()
        {
            // 重入保护（第二次 F9 时如果上一次未释放，则直接忽略）
            if (_inProgress)
            {
                Log.Write("Skip: previous run still in progress");
                return;
            }
            _inProgress = true;

            try
            {
                var swTotal = Stopwatch.StartNew();

                // === 复制阶段（基于序列号校验，避免读到历史剪贴板）===
                uint seq0 = GetClipboardSequenceNumber();

                // 快路径：只尝试 Ctrl+C
                SendKeys.SendWait("^c");
                Application.DoEvents();
                Log.Write($"Copy fast-path: ^C sent, seq0={seq0}");

                // 等待剪贴板真的变化
                if (!WaitClipboardSeqChanged(seq0, 500))
                {
                    // 序列号没变，兜底：^A + ^C
                    SendKeys.SendWait("^a");
                    Application.DoEvents();
                    Thread.Sleep(30); // 更短
                    SendKeys.SendWait("^c");
                    Application.DoEvents();
                    Log.Write("Copy fallback: ^A then ^C sent, wait seq change...");

                    if (!WaitClipboardSeqChanged(seq0, 800))
                    {
                        Log.Write("Copy failed: clipboard sequence did not change");
                        MessageBox.Show("没有检测到文本，请先选中文本再试。");
                        _inProgress = false;
                        return;
                    }
                }

                // 到这里，序列号已变化，说明本次复制生效，再去拿文本
                if (!TryGetTextAdaptive(out var text) || string.IsNullOrWhiteSpace(text))
                {
                    Log.Write("GetText after seq-change failed or empty");
                    MessageBox.Show("读取剪贴板文本失败或为空。");
                    _inProgress = false;
                    return;
                }
                Log.Write($"Read text ok, Len={text.Length}");

                // 更快渲染（TextRenderer + Segoe UI Light；无则回退到 Segoe UI） + 略小宽度减少像素量
                using var bmp = RenderTextToImage(text, 380);
                Log.Write($"Render ok, {bmp.Width}x{bmp.Height}");

                // 自适应写入图片
                if (!TrySetImageAdaptive(bmp))
                {
                    Log.Write("SetImage adaptive failed");
                    MessageBox.Show("写入剪贴板图片失败，可能被其它程序占用。");
                    _inProgress = false;
                    return;
                }
                Application.DoEvents();
                Log.Write("SetImage ok");

                // 粘贴
                try
                {
                    SendKeys.SendWait("^v");
                    Application.DoEvents();
                    Log.Write("Paste ok");
                }
                catch (Exception ex)
                {
                    Log.Error("Paste failed", ex);
                }

                // 延迟回车 + 释放重入锁（稍做缓冲，避免飞书还在取剪贴板）
                var t = new System.Windows.Forms.Timer();
                t.Interval = Math.Max(0, config.SendDelayMs);
                t.Tick += (s, e) =>
                {
                    t.Stop();
                    t.Dispose();
                    try
                    {
                        SendKeys.SendWait("{ENTER}");
                        Log.Write("ENTER sent");
                    }
                    catch (Exception ex)
                    {
                        Log.Error("ENTER failed", ex);
                    }

                    // 释放重入（再缓冲 500ms 防止飞书未完成取图）
                    var release = new System.Windows.Forms.Timer();
                    release.Interval = 500; // 可调 300~800
                    release.Tick += (_, __) =>
                    {
                        release.Stop();
                        release.Dispose();
                        _inProgress = false;
                        Log.Write("Done: unlock for next run");
                    };
                    release.Start();
                };
                t.Start();

                Log.Write($"All scheduled in {swTotal.ElapsedMilliseconds}ms");
                Log.Write("End ConvertTextToImageAndSend()");
            }
            catch (Exception ex)
            {
                Log.Error("ConvertTextToImageAndSend exception", ex);
                _inProgress = false; // 异常时别卡死
            }
        }

        // 高质量渲染（GDI+）：ClearType + AntiAlias，画质优先
        private Bitmap RenderTextToImage(string text, int maxWidth)
        {
            // 优先用 Segoe UI Light；若不可用则回退到 Segoe UI
            Font font;
            try
            {
                var f = new Font("Segoe UI Light", 14f, GraphicsUnit.Pixel);
                if (!string.Equals(f.Name, "Segoe UI Light", StringComparison.OrdinalIgnoreCase))
                {
                    f.Dispose();
                    font = new Font("Segoe UI", 14f, GraphicsUnit.Pixel);
                    Log.Write("Font fallback: Segoe UI Light -> Segoe UI (GDI+ mode)");
                }
                else font = f;
            }
            catch
            {
                font = new Font("Segoe UI", 14f, GraphicsUnit.Pixel);
                Log.Write("Font create failed for Light, using Segoe UI (GDI+ mode)");
            }

            int padding = 8;
            int layoutWidth = Math.Max(1, maxWidth - padding * 2);

            // 先用暂存画布测量文本尺寸（配合 StringFormat 控制换行）
            using var measBmp = new Bitmap(1, 1);
            using var measG = Graphics.FromImage(measBmp);
            measG.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            measG.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            measG.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            measG.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

            using var sf = new StringFormat(StringFormatFlags.LineLimit)
            {
                Trimming = StringTrimming.Word // 以单词为单位换行/裁剪
            };

            // 测量高度（注意：MeasureString 返回值略大，稍微向上取整更稳）
            var measured = measG.MeasureString(text, font, layoutWidth, sf);
            int w = Math.Max(1, Math.Min(maxWidth, (int)Math.Ceiling(measured.Width) + padding * 2));
            int h = Math.Max(1, (int)Math.Ceiling(measured.Height) + padding * 2);

            // 用 32bppPArgb 获得更平滑的文本边缘
            var bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);

                // 质量优先的绘制参数
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                var rect = new RectangleF(padding, padding, w - padding * 2, h - padding * 2);

                // 真正绘制
                g.DrawString(text, font, Brushes.Black, rect, sf);
            }

            font.Dispose();
            return bmp;
        }


        // ============ 剪贴板重试工具（自适应） ============
        private bool TryGetTextAdaptive(out string text)
        {
            // 三档退避：300ms -> 700ms -> 1200ms
            if (ClipboardRetryGetText(out text, 300, 20)) return true;
            if (ClipboardRetryGetText(out text, 700, 20)) return true;
            return ClipboardRetryGetText(out text, 1200, 20);
        }

        private bool TrySetImageAdaptive(Image img)
        {
            // 三档退避：400ms -> 900ms -> 1500ms
            if (ClipboardRetrySetImage(img, 400, 20)) return true;
            if (ClipboardRetrySetImage(img, 900, 20)) return true;
            return ClipboardRetrySetImage(img, 1500, 20);
        }

        private bool ClipboardRetryContainsText(int timeoutMs, int pollMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    if (Clipboard.ContainsText()) return true;
                }
                catch { /* 占用，等一下再试 */ }
                Thread.Sleep(pollMs);
                Application.DoEvents();
            }
            return false;
        }

        private bool ClipboardRetryGetText(out string text, int timeoutMs, int pollMs)
        {
            text = "";
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    if (Clipboard.ContainsText())
                    {
                        var t = Clipboard.GetText();
                        if (!string.IsNullOrWhiteSpace(t))
                        {
                            text = t;
                            return true;
                        }
                    }
                }
                catch { /* 占用，等一下再试 */ }
                Thread.Sleep(pollMs);
                Application.DoEvents();
            }
            return false;
        }

        private bool ClipboardRetrySetImage(Image img, int timeoutMs, int pollMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    Clipboard.SetImage(img);
                    return true;
                }
                catch { /* 占用，等一下再试 */ }
                Thread.Sleep(pollMs);
                Application.DoEvents();
            }
            return false;
        }

        // ========= Win32 =========
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private bool PromptDelayMs(int defaultMs, out int ms)
        {
            ms = defaultMs;
            using var form = new Form
            {
                Text = "设置发送延迟（毫秒）",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false,
                TopMost = true,
                ClientSize = new Size(300, 120)
            };
            var lbl = new Label { AutoSize = true, Text = "请输入延迟（毫秒）:", Location = new Point(12, 15) };
            var tb = new TextBox { Location = new Point(15, 40), Width = 260, Text = defaultMs.ToString() };
            var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(120, 80), Width = 70 };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(205, 80), Width = 70 };
            form.Controls.AddRange(new Control[] { lbl, tb, ok, cancel });
            form.AcceptButton = ok;
            form.CancelButton = cancel;

            if (form.ShowDialog() == DialogResult.OK && int.TryParse(tb.Text, out int val))
            {
                ms = val;
                return true;
            }
            return false;
        }

        private bool PromptHotkey(string defaultKey, out string keyName)
        {
            keyName = defaultKey;
            using var form = new Form
            {
                Text = "设置主快捷键（单键）",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false,
                TopMost = true,
                ClientSize = new Size(320, 140)
            };
            var lbl = new Label { AutoSize = true, Text = "请输入按键名（如 F9、F10、A、Q、Space）:", Location = new Point(12, 15) };
            var tb = new TextBox { Location = new Point(15, 40), Width = 280, Text = defaultKey };
            var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(140, 90), Width = 70 };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(225, 90), Width = 70 };
            form.Controls.AddRange(new Control[] { lbl, tb, ok, cancel });
            form.AcceptButton = ok;
            form.CancelButton = cancel;

            // 支持按下键自动填充
            tb.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true;
                tb.Text = (e.KeyCode & Keys.KeyCode).ToString();
            };

            if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(tb.Text))
            {
                keyName = tb.Text.Trim();
                return true;
            }
            return false;
        }

        private bool PromptComboHotkey(HotkeySetting current, out HotkeySetting result)
        {
            result = new HotkeySetting
            {
                Ctrl = current.Ctrl,
                Alt = current.Alt,
                Shift = current.Shift,
                Win = current.Win,
                Key = current.Key
            };

            using var form = new Form
            {
                Text = "设置组合快捷键",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false,
                TopMost = true,
                ClientSize = new Size(360, 180)
            };

            var cbCtrl = new CheckBox { Text = "Ctrl", Left = 15, Top = 15, Checked = current.Ctrl };
            var cbAlt = new CheckBox { Text = "Alt", Left = 80, Top = 15, Checked = current.Alt };
            var cbShift = new CheckBox { Text = "Shift", Left = 145, Top = 15, Checked = current.Shift };
            var cbWin = new CheckBox { Text = "Win", Left = 220, Top = 15, Checked = current.Win };

            var lbl = new Label { Left = 15, Top = 55, AutoSize = true, Text = "主键（按下即可捕获，也可手输）:" };
            var tb = new TextBox { Left = 15, Top = 80, Width = 320, Text = current.Key };

            // 捕获按键
            tb.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true;
                var keyOnly = (e.KeyCode & Keys.KeyCode);
                tb.Text = keyOnly.ToString();
                // 同步修饰按下状态
                cbCtrl.Checked = e.Control;
                cbAlt.Checked = e.Alt;
                cbShift.Checked = e.Shift;
                cbWin.Checked = (e.KeyData & Keys.LWin) == Keys.LWin || (e.KeyData & Keys.RWin) == Keys.RWin;
            };

            var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Left = 185, Width = 70, Top = 120 };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Left = 265, Width = 70, Top = 120 };

            form.Controls.AddRange(new Control[] { cbCtrl, cbAlt, cbShift, cbWin, lbl, tb, ok, cancel });
            form.AcceptButton = ok;
            form.CancelButton = cancel;

            if (form.ShowDialog() == DialogResult.OK)
            {
                var k = tb.Text?.Trim();
                if (!string.IsNullOrWhiteSpace(k) && Enum.TryParse(k, true, out Keys parsed) && parsed != Keys.None)
                {
                    result.Ctrl = cbCtrl.Checked;
                    result.Alt = cbAlt.Checked;
                    result.Shift = cbShift.Checked;
                    result.Win = cbWin.Checked;
                    result.Key = (parsed & Keys.KeyCode).ToString();
                    return true;
                }
                MessageBox.Show("无效主键，请输入或按下一个有效按键。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return false;
        }
    }

    // ================= 热键窗口 =================
    internal class HotkeyWindow : NativeWindow, IDisposable
    {
        public const int WM_HOTKEY = 0x0312;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;
        public const uint MOD_SHIFT = 0x0004;
        public const uint MOD_WIN = 0x0008;

        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public event Action<int>? HotkeyPressed;

        public HotkeyWindow()
        {
            CreateHandle(new CreateParams());
        }

        public void Register(int id, uint mod, Keys key)
        {
            // 注销再注册，避免冲突
            TryUnregister(id);
            if (key == Keys.None) return;

            if (!RegisterHotKey(Handle, id, mod, (uint)key))
            {
                Log.Write($"RegisterHotKey failed: id=0x{id:X}, mod={mod}, key={key}");
                throw new InvalidOperationException($"注册热键失败：id={id}, mod={mod}, key={key}");
            }
            else
            {
                Log.Write($"RegisterHotKey ok: id=0x{id:X}, mod={mod}, key={key}");
            }
        }

        public void TryUnregister(int id)
        {
            try
            {
                UnregisterHotKey(Handle, id);
                //Log.Write($"UnregisterHotKey id=0x{id:X}");
            }
            catch (Exception ex)
            {
                Log.Error($"UnregisterHotKey failed id=0x{id:X}", ex);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                HotkeyPressed?.Invoke(id);
            }
            base.WndProc(ref m);
        }

        public void Dispose()
        {
            TryUnregister(0xA001);
            TryUnregister(0xA004);
            DestroyHandle();
        }
    }
}
