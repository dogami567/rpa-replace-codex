namespace RpaReplace.DesktopAgent.Interop;

internal static class InputSimulator
{
    private const uint INPUT_MOUSE = 0;
    private const uint INPUT_KEYBOARD = 1;

    private const uint MOUSEEVENTF_MOVE = 0x0001;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    private const uint MOUSEEVENTF_XDOWN = 0x0080;
    private const uint MOUSEEVENTF_XUP = 0x0100;
    private const uint MOUSEEVENTF_WHEEL = 0x0800;
    private const uint MOUSEEVENTF_HWHEEL = 0x01000;
    private const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;
    private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;

    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_UNICODE = 0x0004;

    private const uint XBUTTON1 = 0x0001;
    private const uint XBUTTON2 = 0x0002;

    private const int WHEEL_DELTA = 120;

    private const int SM_CXVIRTUALSCREEN = 78;
    private const int SM_CYVIRTUALSCREEN = 79;
    private const int SM_XVIRTUALSCREEN = 76;
    private const int SM_YVIRTUALSCREEN = 77;

    public static void LeftClickAtScreenPoint(int x, int y)
    {
        var (absoluteX, absoluteY) = ToAbsoluteVirtualDesktopCoords(x, y);

        var inputs = new List<INPUT>
        {
            MouseMoveAbsolute(absoluteX, absoluteY),
            MouseLeftDown(),
            MouseLeftUp(),
        };

        _ = Win32.SendInput(inputs);
    }

    public static void MoveMouseToScreenPoint(int x, int y)
    {
        var (absoluteX, absoluteY) = ToAbsoluteVirtualDesktopCoords(x, y);
        _ = Win32.SendInput([MouseMoveAbsolute(absoluteX, absoluteY)]);
    }

    public static void MoveMouseToScreenPoint(int x, int y, int steps, int durationMs)
    {
        steps = Math.Max(1, steps);
        durationMs = Math.Max(0, durationMs);

        if (steps == 1)
        {
            MoveMouseToScreenPoint(x, y);
            return;
        }

        var start = Win32.GetCursorPos();
        int delayMs = durationMs > 0 ? (int)Math.Round(durationMs / (double)(steps - 1)) : 0;

        for (int i = 1; i <= steps; i++)
        {
            double t = i / (double)steps;
            int ix = (int)Math.Round(start.X + ((x - start.X) * t));
            int iy = (int)Math.Round(start.Y + ((y - start.Y) * t));
            MoveMouseToScreenPoint(ix, iy);

            if (delayMs > 0 && i != steps)
            {
                Thread.Sleep(delayMs);
            }
        }
    }

    public static void MouseDown(MouseButton button) =>
        _ = Win32.SendInput([MouseButtonDown(button)]);

    public static void MouseUp(MouseButton button) =>
        _ = Win32.SendInput([MouseButtonUp(button)]);

    public static void Click(MouseButton button) =>
        _ = Win32.SendInput([MouseButtonDown(button), MouseButtonUp(button)]);

    public static void DoubleClick(MouseButton button, int interClickDelayMs = 50)
    {
        interClickDelayMs = Math.Clamp(interClickDelayMs, 0, 2000);

        Click(button);
        if (interClickDelayMs > 0)
        {
            Thread.Sleep(interClickDelayMs);
        }

        Click(button);
    }

    public static void ScrollWheel(int deltaY, int deltaX = 0)
    {
        var inputs = new List<INPUT>(2);
        if (deltaY != 0)
        {
            inputs.Add(MouseWheel(deltaY * WHEEL_DELTA));
        }

        if (deltaX != 0)
        {
            inputs.Add(MouseHWheel(deltaX * WHEEL_DELTA));
        }

        _ = Win32.SendInput(inputs);
    }

    public static void DragMouse(
        int fromX,
        int fromY,
        int toX,
        int toY,
        MouseButton button,
        int steps,
        int durationMs,
        int afterDownDelayMs = 50,
        int beforeUpDelayMs = 50)
    {
        afterDownDelayMs = Math.Clamp(afterDownDelayMs, 0, 5000);
        beforeUpDelayMs = Math.Clamp(beforeUpDelayMs, 0, 5000);

        MoveMouseToScreenPoint(fromX, fromY);
        Thread.Sleep(50);
        MouseDown(button);
        if (afterDownDelayMs > 0)
        {
            Thread.Sleep(afterDownDelayMs);
        }

        MoveMouseToScreenPoint(toX, toY, steps: steps, durationMs: durationMs);

        if (beforeUpDelayMs > 0)
        {
            Thread.Sleep(beforeUpDelayMs);
        }

        MouseUp(button);
    }

    public static void SendText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var inputs = new List<INPUT>(text.Length * 2);
        foreach (char c in text)
        {
            inputs.Add(KeyboardUnicodeDown(c));
            inputs.Add(KeyboardUnicodeUp(c));
        }

        _ = Win32.SendInput(inputs);
    }

    public static void SendText(string text, int delayMsBetweenChars)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        delayMsBetweenChars = Math.Clamp(delayMsBetweenChars, 0, 5000);

        foreach (char c in text)
        {
            _ = Win32.SendInput([KeyboardUnicodeDown(c), KeyboardUnicodeUp(c)]);
            if (delayMsBetweenChars > 0)
            {
                Thread.Sleep(delayMsBetweenChars);
            }
        }
    }

    public static void KeyDown(ushort virtualKey) =>
        _ = Win32.SendInput([KeyboardVkDown(virtualKey)]);

    public static void KeyUp(ushort virtualKey) =>
        _ = Win32.SendInput([KeyboardVkUp(virtualKey)]);

    public static void KeyPress(ushort virtualKey, int repeat = 1, int interKeyDelayMs = 0)
    {
        repeat = Math.Clamp(repeat, 1, 100);
        interKeyDelayMs = Math.Clamp(interKeyDelayMs, 0, 5000);

        for (int i = 0; i < repeat; i++)
        {
            _ = Win32.SendInput([KeyboardVkDown(virtualKey), KeyboardVkUp(virtualKey)]);
            if (interKeyDelayMs > 0 && i != repeat - 1)
            {
                Thread.Sleep(interKeyDelayMs);
            }
        }
    }

    public static void Hotkey(IReadOnlyList<ushort> virtualKeys, int interKeyDelayMs = 0)
    {
        if (virtualKeys.Count == 0)
        {
            return;
        }

        interKeyDelayMs = Math.Clamp(interKeyDelayMs, 0, 5000);

        var inputs = new List<INPUT>(virtualKeys.Count * 2 + 2);

        for (int i = 0; i < virtualKeys.Count - 1; i++)
        {
            inputs.Add(KeyboardVkDown(virtualKeys[i]));
        }

        ushort last = virtualKeys[^1];
        inputs.Add(KeyboardVkDown(last));
        inputs.Add(KeyboardVkUp(last));

        for (int i = virtualKeys.Count - 2; i >= 0; i--)
        {
            inputs.Add(KeyboardVkUp(virtualKeys[i]));
        }

        _ = Win32.SendInput(inputs);

        if (interKeyDelayMs > 0)
        {
            Thread.Sleep(interKeyDelayMs);
        }
    }

    private static INPUT MouseMoveAbsolute(int absoluteX, int absoluteY) =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = absoluteX,
                    dy = absoluteY,
                    mouseData = 0,
                    dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseLeftDown() =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = 0,
                    dwFlags = MOUSEEVENTF_LEFTDOWN,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseLeftUp() =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = 0,
                    dwFlags = MOUSEEVENTF_LEFTUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseRightDown() =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = 0,
                    dwFlags = MOUSEEVENTF_RIGHTDOWN,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseRightUp() =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = 0,
                    dwFlags = MOUSEEVENTF_RIGHTUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseMiddleDown() =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = 0,
                    dwFlags = MOUSEEVENTF_MIDDLEDOWN,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseMiddleUp() =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = 0,
                    dwFlags = MOUSEEVENTF_MIDDLEUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseXDown(uint xButton) =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = xButton,
                    dwFlags = MOUSEEVENTF_XDOWN,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseXUp(uint xButton) =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = xButton,
                    dwFlags = MOUSEEVENTF_XUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseWheel(int wheelDelta) =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = unchecked((uint)wheelDelta),
                    dwFlags = MOUSEEVENTF_WHEEL,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseHWheel(int wheelDelta) =>
        new()
        {
            type = INPUT_MOUSE,
            U = new InputUnion
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = unchecked((uint)wheelDelta),
                    dwFlags = MOUSEEVENTF_HWHEEL,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT MouseButtonDown(MouseButton button) =>
        button switch
        {
            MouseButton.Left => MouseLeftDown(),
            MouseButton.Right => MouseRightDown(),
            MouseButton.Middle => MouseMiddleDown(),
            MouseButton.X1 => MouseXDown(XBUTTON1),
            MouseButton.X2 => MouseXDown(XBUTTON2),
            _ => MouseLeftDown(),
        };

    private static INPUT MouseButtonUp(MouseButton button) =>
        button switch
        {
            MouseButton.Left => MouseLeftUp(),
            MouseButton.Right => MouseRightUp(),
            MouseButton.Middle => MouseMiddleUp(),
            MouseButton.X1 => MouseXUp(XBUTTON1),
            MouseButton.X2 => MouseXUp(XBUTTON2),
            _ => MouseLeftUp(),
        };

    private static INPUT KeyboardUnicodeDown(char c) =>
        new()
        {
            type = INPUT_KEYBOARD,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = c,
                    dwFlags = KEYEVENTF_UNICODE,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT KeyboardUnicodeUp(char c) =>
        new()
        {
            type = INPUT_KEYBOARD,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = c,
                    dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT KeyboardVkDown(ushort virtualKey) =>
        new()
        {
            type = INPUT_KEYBOARD,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = virtualKey,
                    wScan = 0,
                    dwFlags = IsExtendedKey(virtualKey) ? KEYEVENTF_EXTENDEDKEY : 0,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static INPUT KeyboardVkUp(ushort virtualKey) =>
        new()
        {
            type = INPUT_KEYBOARD,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = virtualKey,
                    wScan = 0,
                    dwFlags = (IsExtendedKey(virtualKey) ? KEYEVENTF_EXTENDEDKEY : 0) | KEYEVENTF_KEYUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero,
                },
            },
        };

    private static bool IsExtendedKey(ushort virtualKey) =>
        virtualKey is 0x21 // VK_PRIOR (Page Up)
            or 0x22 // VK_NEXT (Page Down)
            or 0x23 // VK_END
            or 0x24 // VK_HOME
            or 0x25 // VK_LEFT
            or 0x26 // VK_UP
            or 0x27 // VK_RIGHT
            or 0x28 // VK_DOWN
            or 0x2D // VK_INSERT
            or 0x2E // VK_DELETE
            or 0x5B // VK_LWIN
            or 0x5C // VK_RWIN
            or 0x5D // VK_APPS
            or 0xA3 // VK_RCONTROL
            or 0xA5; // VK_RMENU (Right Alt)

    private static (int X, int Y) ToAbsoluteVirtualDesktopCoords(int screenX, int screenY)
    {
        int vx = Win32.GetSystemMetric(SM_XVIRTUALSCREEN);
        int vy = Win32.GetSystemMetric(SM_YVIRTUALSCREEN);
        int vw = Win32.GetSystemMetric(SM_CXVIRTUALSCREEN);
        int vh = Win32.GetSystemMetric(SM_CYVIRTUALSCREEN);

        if (vw <= 1 || vh <= 1)
        {
            return (0, 0);
        }

        int normalizedX = (int)Math.Round((screenX - vx) * 65535.0 / (vw - 1));
        int normalizedY = (int)Math.Round((screenY - vy) * 65535.0 / (vh - 1));

        normalizedX = Math.Clamp(normalizedX, 0, 65535);
        normalizedY = Math.Clamp(normalizedY, 0, 65535);

        return (normalizedX, normalizedY);
    }
}

internal enum MouseButton
{
    Left = 0,
    Right = 1,
    Middle = 2,
    X1 = 3,
    X2 = 4,
}
