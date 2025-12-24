using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace RpaReplace.DesktopAgent.Interop;

internal static class Win32
{
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    public static IReadOnlyList<WindowInfo> EnumerateTopLevelWindows(bool visibleOnly = true, bool includeEmptyTitles = false)
    {
        var windows = new List<WindowInfo>();

        _ = EnumWindows((hwnd, lParam) =>
        {
            if (hwnd == IntPtr.Zero)
            {
                return true;
            }

            if (visibleOnly && !IsWindowVisible(hwnd))
            {
                return true;
            }

            string title = GetWindowText(hwnd);
            if (!includeEmptyTitles && string.IsNullOrWhiteSpace(title))
            {
                return true;
            }

            _ = GetWindowRect(hwnd, out var rect);
            _ = GetWindowThreadProcessId(hwnd, out uint pid);

            string? processName = null;
            try
            {
                processName = Process.GetProcessById((int)pid).ProcessName;
            }
            catch
            {
                // ignore
            }

            windows.Add(new WindowInfo
            {
                Hwnd = hwnd,
                Title = title,
                ProcessId = (int)pid,
                ProcessName = processName,
                Rect = rect,
                IsVisible = IsWindowVisible(hwnd),
                ClassName = GetClassName(hwnd),
            });

            return true;
        }, IntPtr.Zero);

        return windows;
    }

    public static string GetWindowText(IntPtr hwnd)
    {
        int length = GetWindowTextLengthW(hwnd);
        if (length <= 0)
        {
            return string.Empty;
        }

        var buffer = new StringBuilder(length + 1);
        _ = GetWindowTextW(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    public static string GetClassName(IntPtr hwnd)
    {
        var buffer = new StringBuilder(256);
        _ = GetClassNameW(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    public static NativePoint GetCursorPos()
    {
        _ = GetCursorPos(out var point);
        return point;
    }

    public static void SetCursorPos(int x, int y) => _ = SetCursorPosNative(x, y);

    public static void SetForegroundWindow(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        _ = SetForegroundWindowNative(hwnd);
    }

    public static int GetSystemMetric(int index) => GetSystemMetrics(index);

    public static uint SendInput(IReadOnlyList<INPUT> inputs)
    {
        if (inputs.Count == 0)
        {
            return 0;
        }

        return SendInputNative((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
    }

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLengthW(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect lpRect);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
    private static extern bool SetCursorPosNative(int X, int Y);

    [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
    private static extern bool SetForegroundWindowNative(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll", EntryPoint = "SendInput")]
    private static extern uint SendInputNative(uint nInputs, INPUT[] pInputs, int cbSize);
}

internal sealed class WindowInfo
{
    public required IntPtr Hwnd { get; init; }
    public required string Title { get; init; }
    public required int ProcessId { get; init; }
    public string? ProcessName { get; init; }
    public required NativeRect Rect { get; init; }
    public required bool IsVisible { get; init; }
    public required string ClassName { get; init; }

    public string HwndHex => $"0x{Hwnd.ToInt64():X}";
}

[StructLayout(LayoutKind.Sequential)]
internal struct NativePoint
{
    public int X;
    public int Y;
}

[StructLayout(LayoutKind.Sequential)]
internal struct NativeRect
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;

    public int Width => Right - Left;
    public int Height => Bottom - Top;
    public int CenterX => Left + (Width / 2);
    public int CenterY => Top + (Height / 2);

    public override readonly string ToString() => $"({Left},{Top})-({Right},{Bottom})";
}

[StructLayout(LayoutKind.Sequential)]
internal struct INPUT
{
    public uint type;
    public InputUnion U;
}

[StructLayout(LayoutKind.Explicit)]
internal struct InputUnion
{
    [FieldOffset(0)]
    public MOUSEINPUT mi;

    [FieldOffset(0)]
    public KEYBDINPUT ki;
}

[StructLayout(LayoutKind.Sequential)]
internal struct MOUSEINPUT
{
    public int dx;
    public int dy;
    public uint mouseData;
    public uint dwFlags;
    public uint time;
    public IntPtr dwExtraInfo;
}

[StructLayout(LayoutKind.Sequential)]
internal struct KEYBDINPUT
{
    public ushort wVk;
    public ushort wScan;
    public uint dwFlags;
    public uint time;
    public IntPtr dwExtraInfo;
}
