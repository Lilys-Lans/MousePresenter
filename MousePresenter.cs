// Mouse Presenter для Windows 8.1 / 10 / 11
//
// Пока на переднем плане идёт показ слайдов PowerPoint:
//   Правая кнопка   -> следующий слайд
//   Левая кнопка    -> предыдущий слайд
//   Колесо вверх/вниз -> громкость системы
//   Нажатие колеса  -> выход из показа (Esc), PowerPoint остаётся открытым
//
// Вне показа слайдов мышь работает как обычно.
// Программа живёт в трее (значок мышки). Выход: правый клик по значку -> "Выход".
//
// Код совместим с C# 5, чтобы собираться штатным csc.exe из Windows 8.1.

using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("Mouse Presenter")]
[assembly: AssemblyProduct("Mouse Presenter")]
[assembly: AssemblyDescription("Управление слайдами PowerPoint кнопками мыши")]
[assembly: AssemblyVersion("1.0.0.0")]

internal static class Program
{
    // ---- Константы Windows ----
    private const int WH_MOUSE_LL = 14;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_LBUTTONUP = 0x0202;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_RBUTTONUP = 0x0205;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_MBUTTONUP = 0x0208;
    private const int WM_MOUSEWHEEL = 0x020A;

    private const uint LLMHF_INJECTED = 0x01;

    private const byte VK_PRIOR = 0x21;        // Page Up
    private const byte VK_NEXT = 0x22;         // Page Down
    private const byte VK_ESCAPE = 0x1B;
    private const byte VK_VOLUME_DOWN = 0xAE;
    private const byte VK_VOLUME_UP = 0xAF;

    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
    private const uint KEYEVENTF_KEYUP = 0x0002;

    // Класс окна показа слайдов PowerPoint (не зависит от языка интерфейса)
    private const string SlideShowClass = "screenclass";

    // Запасной вариант: фразы в заголовке окна показа
    private static readonly string[] SlideShowTitleHints =
    {
        "powerpoint slide show",
        "powerpoint slideshow",
        "показ слайдов",
        "demostración de powerpoint",
        "diapositivas de powerpoint"
    };

    // Процессы, для которых включаем управление
    private static readonly string[] AllowedProcesses = { "powerpnt", "pptview" };

    // ---- Структуры и импорт ----
    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    // ---- Состояние ----
    // Делегат хранится в статическом поле, чтобы сборщик мусора его не удалил.
    private static readonly LowLevelMouseProc HookProc = HookCallback;
    private static IntPtr hookHandle = IntPtr.Zero;

    // Кнопки, нажатие которых мы подавили: их отпускание тоже нужно подавить,
    // иначе PowerPoint может показать контекстное меню.
    private static bool suppressLeftUp;
    private static bool suppressRightUp;
    private static bool suppressMiddleUp;

    [STAThread]
    private static void Main()
    {
        bool createdNew;
        using (Mutex mutex = new Mutex(true, "MousePresenter_SingleInstance", out createdNew))
        {
            if (!createdNew)
            {
                MessageBox.Show("Mouse Presenter уже запущен (значок мышки в трее).",
                    "Mouse Presenter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Application.EnableVisualStyles();

            string moduleName = Process.GetCurrentProcess().MainModule.ModuleName;
            hookHandle = SetWindowsHookEx(WH_MOUSE_LL, HookProc, GetModuleHandle(moduleName), 0);
            if (hookHandle == IntPtr.Zero)
            {
                MessageBox.Show("Не удалось установить перехват мыши. Код ошибки: " +
                    Marshal.GetLastWin32Error(),
                    "Mouse Presenter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Icon icon;
            try { icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { icon = SystemIcons.Application; }

            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem info = new ToolStripMenuItem("Mouse Presenter работает");
            info.Enabled = false;
            ToolStripMenuItem exit = new ToolStripMenuItem("Выход");
            exit.Click += delegate { Application.Exit(); };
            menu.Items.Add(info);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exit);

            NotifyIcon tray = new NotifyIcon();
            tray.Icon = icon;
            tray.Text = "Mouse Presenter (ПКМ - выход)";
            tray.ContextMenuStrip = menu;
            tray.Visible = true;
            tray.ShowBalloonTip(4000, "Mouse Presenter запущен",
                "Правая кнопка - вперёд, левая - назад, колесо - громкость, " +
                "нажатие колеса - выход из показа.",
                ToolTipIcon.Info);

            Application.Run();

            tray.Visible = false;
            tray.Dispose();
            UnhookWindowsHookEx(hookHandle);
            hookHandle = IntPtr.Zero;
        }
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            int msg = (int)wParam;

            // Быстрый выход для остальных событий (движение мыши и т.д.)
            if (msg == WM_LBUTTONDOWN || msg == WM_LBUTTONUP ||
                msg == WM_RBUTTONDOWN || msg == WM_RBUTTONUP ||
                msg == WM_MBUTTONDOWN || msg == WM_MBUTTONUP ||
                msg == WM_MOUSEWHEEL)
            {
                MSLLHOOKSTRUCT data = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));

                if ((data.flags & LLMHF_INJECTED) == 0)
                {
                    switch (msg)
                    {
                        case WM_RBUTTONDOWN:
                            if (IsPowerPointSlideShow())
                            {
                                suppressRightUp = true;
                                PressKey(VK_NEXT, true);       // следующий слайд
                                return (IntPtr)1;
                            }
                            break;

                        case WM_LBUTTONDOWN:
                            if (IsPowerPointSlideShow())
                            {
                                suppressLeftUp = true;
                                PressKey(VK_PRIOR, true);      // предыдущий слайд
                                return (IntPtr)1;
                            }
                            break;

                        case WM_MBUTTONDOWN:
                            if (IsPowerPointSlideShow())
                            {
                                suppressMiddleUp = true;
                                PressKey(VK_ESCAPE, false);    // выход из показа
                                return (IntPtr)1;
                            }
                            break;

                        case WM_RBUTTONUP:
                            if (suppressRightUp) { suppressRightUp = false; return (IntPtr)1; }
                            break;

                        case WM_LBUTTONUP:
                            if (suppressLeftUp) { suppressLeftUp = false; return (IntPtr)1; }
                            break;

                        case WM_MBUTTONUP:
                            if (suppressMiddleUp) { suppressMiddleUp = false; return (IntPtr)1; }
                            break;

                        case WM_MOUSEWHEEL:
                            if (IsPowerPointSlideShow())
                            {
                                short delta = unchecked((short)((data.mouseData >> 16) & 0xFFFF));
                                if (delta > 0) PressKey(VK_VOLUME_UP, false);
                                else if (delta < 0) PressKey(VK_VOLUME_DOWN, false);
                                return (IntPtr)1;              // не листать слайды колесом
                            }
                            break;
                    }
                }
            }
        }

        return CallNextHookEx(hookHandle, nCode, wParam, lParam);
    }

    private static bool IsPowerPointSlideShow()
    {
        try
        {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return false;

            StringBuilder cls = new StringBuilder(256);
            GetClassName(hwnd, cls, cls.Capacity);
            bool looksLikeShow = cls.ToString().Equals(SlideShowClass, StringComparison.OrdinalIgnoreCase);

            if (!looksLikeShow)
            {
                StringBuilder title = new StringBuilder(512);
                GetWindowText(hwnd, title, title.Capacity);
                string t = title.ToString().ToLowerInvariant();
                foreach (string hint in SlideShowTitleHints)
                {
                    if (t.Contains(hint)) { looksLikeShow = true; break; }
                }
            }
            if (!looksLikeShow) return false;

            // Дополнительная проверка: окно должно принадлежать PowerPoint
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            string name = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant();
            foreach (string allowed in AllowedProcesses)
            {
                if (name == allowed) return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    private static void PressKey(byte vk, bool extended)
    {
        uint flags = extended ? KEYEVENTF_EXTENDEDKEY : 0;
        keybd_event(vk, 0, flags, UIntPtr.Zero);
        keybd_event(vk, 0, flags | KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
