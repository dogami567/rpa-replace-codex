using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using RpaReplace.DesktopAgent.Interop;
using RpaReplace.DesktopAgent.Uia;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.Json;
using System.Windows.Automation;

namespace RpaReplace.DesktopAgent.Tools;

[McpServerToolType]
public sealed class DesktopTools
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
    };

    [McpServerTool(Name = "list_windows", UseStructuredContent = true, ReadOnly = true)]
    [Description("List top-level windows on the desktop.")]
    public static CallToolResult ListWindows(
        [Description("Only include visible windows.")] bool visibleOnly = true,
        [Description("Include windows with empty titles.")] bool includeEmptyTitles = false)
    {
        var windows = Win32.EnumerateTopLevelWindows(visibleOnly: visibleOnly, includeEmptyTitles: includeEmptyTitles)
            .Select(w => new
            {
                hwnd = w.HwndHex,
                title = w.Title,
                processId = w.ProcessId,
                processName = w.ProcessName,
                className = w.ClassName,
                rect = new { left = w.Rect.Left, top = w.Rect.Top, right = w.Rect.Right, bottom = w.Rect.Bottom },
                isVisible = w.IsVisible,
            })
            .ToArray();

        return JsonResult(windows);
    }

    [McpServerTool(Name = "inspect", UseStructuredContent = true, ReadOnly = true)]
    [Description("Inspect the UI Automation element at a screen coordinate (defaults to current cursor position).")]
    public static CallToolResult Inspect(
        [Description("Screen X coordinate.")] int? x = null,
        [Description("Screen Y coordinate.")] int? y = null)
    {
        var point = x is not null && y is not null ? new NativePoint { X = x.Value, Y = y.Value } : Win32.GetCursorPos();
        var element = AutomationElement.FromPoint(new System.Windows.Point(point.X, point.Y));

        return JsonResult(DescribeElement(element));
    }

    [McpServerTool(Name = "query", UseStructuredContent = true, ReadOnly = true)]
    [Description("Find an element using UI Automation within an optional window scope.")]
    public static CallToolResult Query(
        [Description("Window title substring match.")] string? windowTitle = null,
        [Description("Window title regex match.")] string? windowTitleRegex = null,
        [Description("Window handle (HWND) as hex string, e.g. 0x1234.")] string? windowHwnd = null,
        [Description("Target element Name.")] string? name = null,
        [Description("Target element Name regex.")] string? nameRegex = null,
        [Description("Target element AutomationId.")] string? automationId = null,
        [Description("Target element ClassName.")] string? className = null,
        [Description("Target element ControlType (e.g., Button, Edit).")] string? controlType = null,
        [Description("Which match index to return (0-based).")] int index = 0)
    {
        var windowQuery = new UiaWindowQuery(
            Title: windowTitle,
            TitleRegex: windowTitleRegex,
            Hwnd: ParseHwnd(windowHwnd),
            Index: 0);

        var windowRoot = UiaQuery.ResolveWindowRoot(windowQuery);
        if (windowRoot is null && HasAnyWindowFilter(windowTitle, windowTitleRegex, windowHwnd))
        {
            return ErrorResult("Window not found.");
        }

        var root = windowRoot ?? AutomationElement.RootElement;

        var elementQuery = new UiaElementQuery(
            Name: name,
            NameRegex: nameRegex,
            AutomationId: automationId,
            ClassName: className,
            ControlType: controlType,
            Index: index);

        var element = UiaQuery.FindElement(root, elementQuery);
        if (element is null)
        {
            return ErrorResult($"Element not found (index={index}).");
        }

        return JsonResult(DescribeElement(element));
    }

    [McpServerTool(Name = "click", UseStructuredContent = true)]
    [Description("Click an element. Tries UIA patterns first, then falls back to mouse click at bounding rectangle center.")]
    public static CallToolResult Click(
        [Description("Window title substring match.")] string? windowTitle = null,
        [Description("Window title regex match.")] string? windowTitleRegex = null,
        [Description("Window handle (HWND) as hex string, e.g. 0x1234.")] string? windowHwnd = null,
        [Description("Target element Name.")] string? name = null,
        [Description("Target element AutomationId.")] string? automationId = null,
        [Description("Target element ClassName.")] string? className = null,
        [Description("Target element ControlType (e.g., Button, Edit).")] string? controlType = null,
        [Description("Which match index to click (0-based).")] int index = 0,
        [Description("Bring the target window to foreground before clicking.")] bool focusWindow = true)
    {
        var element = TryFindElement(windowTitle, windowTitleRegex, windowHwnd, name, null, automationId, className, controlType, index);
        if (element is null)
        {
            return ErrorResult("Element not found.");
        }

        if (focusWindow)
        {
            FocusTopLevelWindow(element);
        }

        if (TryInvoke(element, out string? invokedBy))
        {
            return JsonResult(new { ok = true, method = invokedBy });
        }

        var rect = element.Current.BoundingRectangle;
        if (rect.IsEmpty)
        {
            return ErrorResult("Element has empty BoundingRectangle; cannot click by coordinates.");
        }

        int cx = (int)Math.Round(rect.X + (rect.Width / 2));
        int cy = (int)Math.Round(rect.Y + (rect.Height / 2));
        InputSimulator.LeftClickAtScreenPoint(cx, cy);

        return JsonResult(new { ok = true, method = "mouse", x = cx, y = cy });
    }

    [McpServerTool(Name = "type", UseStructuredContent = true)]
    [Description("Type text into an element. Tries ValuePattern.SetValue first, then falls back to focusing element and sending keyboard input.")]
    public static CallToolResult TypeText(
        [Description("Text to input.")] string text,
        [Description("Window title substring match.")] string? windowTitle = null,
        [Description("Window title regex match.")] string? windowTitleRegex = null,
        [Description("Window handle (HWND) as hex string, e.g. 0x1234.")] string? windowHwnd = null,
        [Description("Target element Name.")] string? name = null,
        [Description("Target element AutomationId.")] string? automationId = null,
        [Description("Target element ClassName.")] string? className = null,
        [Description("Target element ControlType (e.g., Edit).")] string? controlType = null,
        [Description("Which match index to type into (0-based).")] int index = 0,
        [Description("Clear existing value if ValuePattern is used.")] bool clearFirst = true,
        [Description("Bring the target window to foreground before typing.")] bool focusWindow = true)
    {
        var element = TryFindElement(windowTitle, windowTitleRegex, windowHwnd, name, null, automationId, className, controlType, index);
        if (element is null)
        {
            return ErrorResult("Element not found.");
        }

        if (focusWindow)
        {
            FocusTopLevelWindow(element);
        }

        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var patternObj)
            && patternObj is ValuePattern valuePattern
            && !valuePattern.Current.IsReadOnly)
        {
            valuePattern.SetValue(clearFirst ? text : (valuePattern.Current.Value ?? string.Empty) + text);
            return JsonResult(new { ok = true, method = "valuePattern" });
        }

        TryFocusOrClick(element);
        InputSimulator.SendText(text);

        return JsonResult(new { ok = true, method = "keyboard" });
    }

    [McpServerTool(Name = "keyboard_type", UseStructuredContent = true)]
    [Description("Type text into the currently focused element (no UIA targeting).")]
    public static CallToolResult KeyboardType(
        [Description("Text to type.")] string text,
        [Description("Delay between characters in milliseconds (0 = no delay).")] int delayMsBetweenChars = 0,
        [Description("Press Enter after typing.")] bool submit = false)
    {
        if (delayMsBetweenChars > 0)
        {
            InputSimulator.SendText(text, delayMsBetweenChars: delayMsBetweenChars);
        }
        else
        {
            InputSimulator.SendText(text);
        }

        if (submit)
        {
            InputSimulator.KeyPress(0x0D); // VK_RETURN
        }

        return JsonResult(new
        {
            ok = true,
            delayMsBetweenChars = Math.Clamp(delayMsBetweenChars, 0, 5000),
            submit,
        });
    }

    [McpServerTool(Name = "key_down", UseStructuredContent = true)]
    [Description("Press (and hold) a key using a virtual-key code.")]
    public static CallToolResult KeyDown(
        [Description("Key name (e.g., Enter, Tab, Ctrl, Shift, Alt, A, F5) or virtual-key code (e.g., 0x41 or 65).")] string key)
    {
        if (!TryParseVirtualKey(key, out ushort vk, out string error))
        {
            return ErrorResult(error);
        }

        InputSimulator.KeyDown(vk);
        return JsonResult(new { ok = true, key = NormalizeKeyName(key), vk });
    }

    [McpServerTool(Name = "key_up", UseStructuredContent = true)]
    [Description("Release a key using a virtual-key code.")]
    public static CallToolResult KeyUp(
        [Description("Key name (e.g., Enter, Tab, Ctrl, Shift, Alt, A, F5) or virtual-key code (e.g., 0x41 or 65).")] string key)
    {
        if (!TryParseVirtualKey(key, out ushort vk, out string error))
        {
            return ErrorResult(error);
        }

        InputSimulator.KeyUp(vk);
        return JsonResult(new { ok = true, key = NormalizeKeyName(key), vk });
    }

    [McpServerTool(Name = "key_press", UseStructuredContent = true)]
    [Description("Press and release a key using a virtual-key code.")]
    public static CallToolResult KeyPress(
        [Description("Key name (e.g., Enter, Tab, Ctrl, Shift, Alt, A, F5) or virtual-key code (e.g., 0x41 or 65).")] string key,
        [Description("Repeat count.")] int repeat = 1,
        [Description("Delay between repeats in milliseconds.")] int interKeyDelayMs = 0)
    {
        if (!TryParseVirtualKey(key, out ushort vk, out string error))
        {
            return ErrorResult(error);
        }

        InputSimulator.KeyPress(vk, repeat: repeat, interKeyDelayMs: interKeyDelayMs);
        return JsonResult(new
        {
            ok = true,
            key = NormalizeKeyName(key),
            vk,
            repeat = Math.Clamp(repeat, 1, 100),
            interKeyDelayMs = Math.Clamp(interKeyDelayMs, 0, 5000),
        });
    }

    [McpServerTool(Name = "hotkey", UseStructuredContent = true)]
    [Description("Press a key combination like Ctrl+Shift+P or Alt+F4.")]
    public static CallToolResult Hotkey(
        [Description("Hotkey string like Ctrl+Shift+P or Alt+F4.")] string keys,
        [Description("Optional delay after sending in milliseconds.")] int delayMs = 0)
    {
        var parts = (keys ?? string.Empty)
            .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

        if (parts.Length == 0)
        {
            return ErrorResult("Missing keys.");
        }

        var vks = new List<ushort>(parts.Length);
        foreach (var part in parts)
        {
            if (!TryParseVirtualKey(part, out ushort vk, out string error))
            {
                return ErrorResult(error);
            }

            vks.Add(vk);
        }

        InputSimulator.Hotkey(vks, interKeyDelayMs: delayMs);
        return JsonResult(new { ok = true, keys = string.Join('+', parts.Select(NormalizeKeyName)), delayMs = Math.Clamp(delayMs, 0, 5000) });
    }

    [McpServerTool(Name = "shutdown", UseStructuredContent = true)]
    [Description("Shut down the desktop agent process.")]
    public static CallToolResult Shutdown(
        [Description("Delay before exit in milliseconds.")] int delayMs = 200)
    {
        delayMs = Math.Clamp(delayMs, 0, 5000);

        _ = Task.Run(async () =>
        {
            await Task.Delay(delayMs);
            Environment.Exit(0);
        });

        return JsonResult(new { ok = true, exitingInMs = delayMs });
    }

    [McpServerTool(Name = "mouse_move", UseStructuredContent = true)]
    [Description("Move the mouse cursor to a screen coordinate.")]
    public static CallToolResult MouseMove(
        [Description("Screen X coordinate.")] int x,
        [Description("Screen Y coordinate.")] int y,
        [Description("Number of steps for smooth movement (1 = instant).")] int steps = 1,
        [Description("Total duration in milliseconds for smooth movement.")] int durationMs = 0)
    {
        InputSimulator.MoveMouseToScreenPoint(x, y, steps: steps, durationMs: durationMs);
        return JsonResult(new { ok = true, x, y, steps = Math.Max(1, steps), durationMs = Math.Max(0, durationMs) });
    }

    [McpServerTool(Name = "mouse_down", UseStructuredContent = true)]
    [Description("Press (and hold) a mouse button.")]
    public static CallToolResult MouseDown(
        [Description("Mouse button: left|right|middle|x1|x2.")] string button = "left")
    {
        if (!TryParseMouseButton(button, out var parsed, out var error))
        {
            return ErrorResult(error);
        }

        InputSimulator.MouseDown(parsed);
        return JsonResult(new { ok = true, button = NormalizeMouseButton(parsed) });
    }

    [McpServerTool(Name = "mouse_up", UseStructuredContent = true)]
    [Description("Release a mouse button.")]
    public static CallToolResult MouseUp(
        [Description("Mouse button: left|right|middle|x1|x2.")] string button = "left")
    {
        if (!TryParseMouseButton(button, out var parsed, out var error))
        {
            return ErrorResult(error);
        }

        InputSimulator.MouseUp(parsed);
        return JsonResult(new { ok = true, button = NormalizeMouseButton(parsed) });
    }

    [McpServerTool(Name = "mouse_click", UseStructuredContent = true)]
    [Description("Click a mouse button at an optional screen coordinate.")]
    public static CallToolResult MouseClick(
        [Description("Mouse button: left|right|middle|x1|x2.")] string button = "left",
        [Description("Screen X coordinate (optional).")] int? x = null,
        [Description("Screen Y coordinate (optional).")] int? y = null)
    {
        if (!TryParseMouseButton(button, out var parsed, out var error))
        {
            return ErrorResult(error);
        }

        if ((x is null) != (y is null))
        {
            return ErrorResult("Provide both x and y, or neither.");
        }

        if (x is not null && y is not null)
        {
            InputSimulator.MoveMouseToScreenPoint(x.Value, y.Value);
        }

        InputSimulator.Click(parsed);
        return JsonResult(new { ok = true, button = NormalizeMouseButton(parsed), x, y });
    }

    [McpServerTool(Name = "mouse_double_click", UseStructuredContent = true)]
    [Description("Double click a mouse button at an optional screen coordinate.")]
    public static CallToolResult MouseDoubleClick(
        [Description("Mouse button: left|right|middle|x1|x2.")] string button = "left",
        [Description("Milliseconds between clicks.")] int interClickDelayMs = 50,
        [Description("Screen X coordinate (optional).")] int? x = null,
        [Description("Screen Y coordinate (optional).")] int? y = null)
    {
        if (!TryParseMouseButton(button, out var parsed, out var error))
        {
            return ErrorResult(error);
        }

        if ((x is null) != (y is null))
        {
            return ErrorResult("Provide both x and y, or neither.");
        }

        if (x is not null && y is not null)
        {
            InputSimulator.MoveMouseToScreenPoint(x.Value, y.Value);
        }

        InputSimulator.DoubleClick(parsed, interClickDelayMs: interClickDelayMs);
        return JsonResult(new { ok = true, button = NormalizeMouseButton(parsed), interClickDelayMs = Math.Clamp(interClickDelayMs, 0, 2000), x, y });
    }

    [McpServerTool(Name = "mouse_drag", UseStructuredContent = true)]
    [Description("Drag the mouse from one screen coordinate to another.")]
    public static CallToolResult MouseDrag(
        [Description("Start screen X coordinate.")] int fromX,
        [Description("Start screen Y coordinate.")] int fromY,
        [Description("End screen X coordinate.")] int toX,
        [Description("End screen Y coordinate.")] int toY,
        [Description("Mouse button: left|right|middle|x1|x2.")] string button = "left",
        [Description("Number of steps for smooth drag (1 = instant).")] int steps = 20,
        [Description("Total duration in milliseconds for smooth drag movement.")] int durationMs = 200,
        [Description("Delay after mouse down in milliseconds.")] int afterDownDelayMs = 50,
        [Description("Delay before mouse up in milliseconds.")] int beforeUpDelayMs = 50)
    {
        if (!TryParseMouseButton(button, out var parsed, out var error))
        {
            return ErrorResult(error);
        }

        InputSimulator.DragMouse(
            fromX: fromX,
            fromY: fromY,
            toX: toX,
            toY: toY,
            button: parsed,
            steps: steps,
            durationMs: durationMs,
            afterDownDelayMs: afterDownDelayMs,
            beforeUpDelayMs: beforeUpDelayMs);

        return JsonResult(new
        {
            ok = true,
            button = NormalizeMouseButton(parsed),
            fromX,
            fromY,
            toX,
            toY,
            steps = Math.Max(1, steps),
            durationMs = Math.Max(0, durationMs),
            afterDownDelayMs = Math.Clamp(afterDownDelayMs, 0, 5000),
            beforeUpDelayMs = Math.Clamp(beforeUpDelayMs, 0, 5000),
        });
    }

    [McpServerTool(Name = "mouse_scroll", UseStructuredContent = true)]
    [Description("Scroll the mouse wheel. Optional x/y moves the cursor before scrolling.")]
    public static CallToolResult MouseScroll(
        [Description("Vertical wheel delta in notches (positive up, negative down).")] int deltaY,
        [Description("Horizontal wheel delta in notches (positive right, negative left).")] int deltaX = 0,
        [Description("Screen X coordinate (optional).")] int? x = null,
        [Description("Screen Y coordinate (optional).")] int? y = null)
    {
        if ((x is null) != (y is null))
        {
            return ErrorResult("Provide both x and y, or neither.");
        }

        if (x is not null && y is not null)
        {
            InputSimulator.MoveMouseToScreenPoint(x.Value, y.Value);
        }

        InputSimulator.ScrollWheel(deltaY: deltaY, deltaX: deltaX);

        return JsonResult(new { ok = true, deltaY, deltaX, x, y });
    }

    [McpServerTool(Name = "invoke", UseStructuredContent = true)]
    [Description("Invoke an element using UI Automation patterns (Invoke/LegacyIAccessible).")]
    public static CallToolResult Invoke(
        [Description("Window title substring match.")] string? windowTitle = null,
        [Description("Window title regex match.")] string? windowTitleRegex = null,
        [Description("Window handle (HWND) as hex string, e.g. 0x1234.")] string? windowHwnd = null,
        [Description("Target element Name.")] string? name = null,
        [Description("Target element AutomationId.")] string? automationId = null,
        [Description("Target element ClassName.")] string? className = null,
        [Description("Target element ControlType (e.g., Button).")] string? controlType = null,
        [Description("Which match index to invoke (0-based).")] int index = 0,
        [Description("Bring the target window to foreground before invoking.")] bool focusWindow = true)
    {
        var element = TryFindElement(windowTitle, windowTitleRegex, windowHwnd, name, null, automationId, className, controlType, index);
        if (element is null)
        {
            return ErrorResult("Element not found.");
        }

        if (focusWindow)
        {
            FocusTopLevelWindow(element);
        }

        if (!TryInvoke(element, out string? invokedBy))
        {
            return ErrorResult("Element does not support InvokePattern/LegacyIAccessiblePattern.");
        }

        return JsonResult(new { ok = true, method = invokedBy });
    }

    [McpServerTool(Name = "wait_for", UseStructuredContent = true, ReadOnly = true, Idempotent = true)]
    [Description("Wait for an element to appear (exists) or disappear (gone).")]
    public static async Task<CallToolResult> WaitFor(
        [Description("State to wait for: exists | gone")] string state,
        [Description("Window title substring match.")] string? windowTitle = null,
        [Description("Window title regex match.")] string? windowTitleRegex = null,
        [Description("Window handle (HWND) as hex string, e.g. 0x1234.")] string? windowHwnd = null,
        [Description("Target element Name.")] string? name = null,
        [Description("Target element Name regex.")] string? nameRegex = null,
        [Description("Target element AutomationId.")] string? automationId = null,
        [Description("Target element ClassName.")] string? className = null,
        [Description("Target element ControlType (e.g., Button, Edit).")] string? controlType = null,
        [Description("Which match index to consider (0-based).")] int index = 0,
        [Description("Timeout in milliseconds.")] int timeoutMs = 10000,
        [Description("Poll interval in milliseconds.")] int pollIntervalMs = 200)
    {
        bool waitExists = state.Equals("exists", StringComparison.OrdinalIgnoreCase);
        bool waitGone = state.Equals("gone", StringComparison.OrdinalIgnoreCase);
        if (!waitExists && !waitGone)
        {
            return ErrorResult("Invalid state. Use 'exists' or 'gone'.");
        }

        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds <= timeoutMs)
        {
            var element = TryFindElement(windowTitle, windowTitleRegex, windowHwnd, name, nameRegex, automationId, className, controlType, index);
            if (waitExists && element is not null)
            {
                return JsonResult(new { ok = true, state = "exists", elapsedMs = sw.ElapsedMilliseconds, element = DescribeElement(element) });
            }

            if (waitGone && element is null)
            {
                return JsonResult(new { ok = true, state = "gone", elapsedMs = sw.ElapsedMilliseconds });
            }

            await Task.Delay(pollIntervalMs);
        }

        return ErrorResult($"Timeout waiting for state='{state}' after {timeoutMs}ms.");
    }

    [McpServerTool(Name = "screenshot", UseStructuredContent = true, ReadOnly = true)]
    [Description("Capture a screenshot of the virtual desktop, a window, or an element. Returns an image/png content block.")]
    public static CallToolResult Screenshot(
        [Description("Window title substring match.")] string? windowTitle = null,
        [Description("Window title regex match.")] string? windowTitleRegex = null,
        [Description("Window handle (HWND) as hex string, e.g. 0x1234.")] string? windowHwnd = null,
        [Description("Target element Name.")] string? name = null,
        [Description("Target element AutomationId.")] string? automationId = null,
        [Description("Target element ClassName.")] string? className = null,
        [Description("Target element ControlType (e.g., Edit).")] string? controlType = null,
        [Description("Which match index to capture (0-based).")] int index = 0,
        [Description("Capture full desktop if true.")] bool fullDesktop = false,
        [Description("If set, saves PNG to this path and returns JSON instead of inline image data.")] string? outputPath = null)
    {
        NativeRect rect;

        if (fullDesktop)
        {
            rect = GetVirtualDesktopRect();
        }
        else if (!string.IsNullOrWhiteSpace(name) || !string.IsNullOrWhiteSpace(automationId) || !string.IsNullOrWhiteSpace(className) || !string.IsNullOrWhiteSpace(controlType))
        {
            var element = TryFindElement(windowTitle, windowTitleRegex, windowHwnd, name, null, automationId, className, controlType, index);
            if (element is null)
            {
                return ErrorResult("Element not found.");
            }

            var br = element.Current.BoundingRectangle;
            if (br.IsEmpty)
            {
                return ErrorResult("Element has empty BoundingRectangle; cannot screenshot.");
            }

            rect = new NativeRect
            {
                Left = (int)Math.Round(br.Left),
                Top = (int)Math.Round(br.Top),
                Right = (int)Math.Round(br.Right),
                Bottom = (int)Math.Round(br.Bottom),
            };
        }
        else if (HasAnyWindowFilter(windowTitle, windowTitleRegex, windowHwnd))
        {
            var window = TryResolveWindow(windowTitle, windowTitleRegex, windowHwnd);
            if (window is null)
            {
                return ErrorResult("Window not found.");
            }

            rect = window.Rect;
        }
        else
        {
            rect = GetVirtualDesktopRect();
        }

        byte[] pngBytes = CaptureRectAsPng(rect);
        if (!string.IsNullOrWhiteSpace(outputPath))
        {
            string fullPath = Path.GetFullPath(outputPath);
            string? dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllBytes(fullPath, pngBytes);
            return JsonResult(new
            {
                ok = true,
                savedTo = fullPath,
                bytes = pngBytes.Length,
                rect = new { left = rect.Left, top = rect.Top, right = rect.Right, bottom = rect.Bottom },
            });
        }

        string base64 = Convert.ToBase64String(pngBytes);

        return new CallToolResult
        {
            Content =
            [
                new ImageContentBlock { Data = base64, MimeType = "image/png" },
            ],
            IsError = false,
        };
    }

    private static CallToolResult JsonResult(object value)
    {
        string json = JsonSerializer.Serialize(value, JsonOptions);
        return new CallToolResult
        {
            Content =
            [
                new TextContentBlock { Text = json },
            ],
            IsError = false,
        };
    }

    private static CallToolResult ErrorResult(string message) =>
        new()
        {
            Content =
            [
                new TextContentBlock { Text = message },
            ],
            IsError = true,
        };

    private static long? ParseHwnd(string? hwnd)
    {
        if (string.IsNullOrWhiteSpace(hwnd))
        {
            return null;
        }

        string trimmed = hwnd.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed[2..];
        }

        if (long.TryParse(trimmed, System.Globalization.NumberStyles.HexNumber, null, out long value))
        {
            return value;
        }

        if (long.TryParse(trimmed, out value))
        {
            return value;
        }

        return null;
    }

    private static WindowInfo? TryResolveWindow(string? windowTitle, string? windowTitleRegex, string? windowHwnd)
    {
        var hwndValue = ParseHwnd(windowHwnd);
        if (hwndValue is not null)
        {
            var hwnd = new IntPtr(hwndValue.Value);
            var match = Win32.EnumerateTopLevelWindows(visibleOnly: false, includeEmptyTitles: true)
                .FirstOrDefault(w => w.Hwnd == hwnd);
            return match;
        }

        var matches = Win32.EnumerateTopLevelWindows(visibleOnly: true, includeEmptyTitles: false)
            .Where(w =>
            {
                if (!string.IsNullOrWhiteSpace(windowTitle) && !w.Title.Contains(windowTitle, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(windowTitleRegex) && !System.Text.RegularExpressions.Regex.IsMatch(w.Title, windowTitleRegex, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    return false;
                }

                return true;
            })
            .ToList();

        return matches.Count == 0 ? null : matches[0];
    }

    private static AutomationElement? TryFindElement(
        string? windowTitle,
        string? windowTitleRegex,
        string? windowHwnd,
        string? name,
        string? nameRegex,
        string? automationId,
        string? className,
        string? controlType,
        int index)
    {
        var windowQuery = new UiaWindowQuery(
            Title: windowTitle,
            TitleRegex: windowTitleRegex,
            Hwnd: ParseHwnd(windowHwnd),
            Index: 0);

        var windowRoot = UiaQuery.ResolveWindowRoot(windowQuery);
        if (windowRoot is null && HasAnyWindowFilter(windowTitle, windowTitleRegex, windowHwnd))
        {
            return null;
        }

        var root = windowRoot ?? AutomationElement.RootElement;

        var elementQuery = new UiaElementQuery(
            Name: name,
            NameRegex: nameRegex,
            AutomationId: automationId,
            ClassName: className,
            ControlType: controlType,
            Index: index);

        return UiaQuery.FindElement(root, elementQuery);
    }

    private static bool TryParseMouseButton(string? value, out MouseButton button, out string error)
    {
        button = MouseButton.Left;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(value))
        {
            return true;
        }

        string normalized = value.Trim();
        if (normalized.StartsWith("MouseButton.", StringComparison.OrdinalIgnoreCase))
        {
            normalized = normalized["MouseButton.".Length..];
        }

        button = normalized.ToLowerInvariant() switch
        {
            "left" => MouseButton.Left,
            "right" => MouseButton.Right,
            "middle" => MouseButton.Middle,
            "x1" => MouseButton.X1,
            "x2" => MouseButton.X2,
            _ => MouseButton.Left,
        };

        if (!normalized.Equals("left", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("right", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("middle", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("x1", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("x2", StringComparison.OrdinalIgnoreCase)
            && !normalized.StartsWith("MouseButton.", StringComparison.OrdinalIgnoreCase))
        {
            error = "Invalid button. Use left|right|middle|x1|x2.";
            return false;
        }

        return true;
    }

    private static string NormalizeMouseButton(MouseButton button) =>
        button switch
        {
            MouseButton.Left => "left",
            MouseButton.Right => "right",
            MouseButton.Middle => "middle",
            MouseButton.X1 => "x1",
            MouseButton.X2 => "x2",
            _ => "left",
        };

    private static string NormalizeKeyName(string value) =>
        string.IsNullOrWhiteSpace(value) ? value : value.Trim();

    private static bool TryParseVirtualKey(string value, out ushort vk, out string error)
    {
        vk = 0;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(value))
        {
            error = "Key is required.";
            return false;
        }

        string k = value.Trim();

        if (k.StartsWith("VK_", StringComparison.OrdinalIgnoreCase))
        {
            k = k["VK_".Length..];
        }

        if (k.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(k[2..], System.Globalization.NumberStyles.HexNumber, null, out int hex)
            && hex >= 0
            && hex <= ushort.MaxValue)
        {
            vk = (ushort)hex;
            return true;
        }

        if (int.TryParse(k, out int dec) && dec >= 0 && dec <= ushort.MaxValue)
        {
            vk = (ushort)dec;
            return true;
        }

        if (k.Length == 1)
        {
            char c = k[0];
            if (char.IsLetter(c))
            {
                vk = (ushort)char.ToUpperInvariant(c);
                return true;
            }

            if (char.IsDigit(c))
            {
                vk = (ushort)c;
                return true;
            }
        }

        if (k.StartsWith("F", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(k[1..], out int f)
            && f is >= 1 and <= 24)
        {
            vk = (ushort)(0x70 + (f - 1)); // VK_F1..VK_F24
            return true;
        }

        vk = k.ToLowerInvariant() switch
        {
            "ctrl" or "control" or "lcontrol" => 0x11, // VK_CONTROL
            "rcontrol" => 0xA3, // VK_RCONTROL
            "shift" or "lshift" => 0x10, // VK_SHIFT
            "rshift" => 0xA1, // VK_RSHIFT
            "alt" or "menu" or "lalt" => 0x12, // VK_MENU
            "ralt" => 0xA5, // VK_RMENU
            "win" or "meta" or "lwin" => 0x5B, // VK_LWIN
            "rwin" => 0x5C, // VK_RWIN
            "enter" or "return" => 0x0D, // VK_RETURN
            "tab" => 0x09, // VK_TAB
            "escape" or "esc" => 0x1B, // VK_ESCAPE
            "backspace" or "bksp" => 0x08, // VK_BACK
            "space" or "spacebar" => 0x20, // VK_SPACE
            "delete" or "del" => 0x2E, // VK_DELETE
            "insert" or "ins" => 0x2D, // VK_INSERT
            "home" => 0x24, // VK_HOME
            "end" => 0x23, // VK_END
            "pageup" or "pgup" or "prior" => 0x21, // VK_PRIOR
            "pagedown" or "pgdn" or "next" => 0x22, // VK_NEXT
            "left" or "arrowleft" => 0x25, // VK_LEFT
            "up" or "arrowup" => 0x26, // VK_UP
            "right" or "arrowright" => 0x27, // VK_RIGHT
            "down" or "arrowdown" => 0x28, // VK_DOWN
            "apps" or "contextmenu" => 0x5D, // VK_APPS
            _ => (ushort)0,
        };

        if (vk != 0)
        {
            return true;
        }

        error = "Unknown key. Use a name like Enter/Ctrl/Shift/Alt/A/F5 or a VK code like 0x41.";
        return false;
    }

    private static bool HasAnyWindowFilter(string? windowTitle, string? windowTitleRegex, string? windowHwnd) =>
        !string.IsNullOrWhiteSpace(windowTitle)
        || !string.IsNullOrWhiteSpace(windowTitleRegex)
        || !string.IsNullOrWhiteSpace(windowHwnd);

    private static void FocusTopLevelWindow(AutomationElement element)
    {
        AutomationElement? current = element;
        while (current is not null)
        {
            int hwnd = current.Current.NativeWindowHandle;
            if (hwnd != 0)
            {
                Win32.SetForegroundWindow(new IntPtr(hwnd));
                return;
            }

            current = TreeWalker.RawViewWalker.GetParent(current);
        }
    }

    private static bool TryInvoke(AutomationElement element, out string? invokedBy)
    {
        invokedBy = null;

        if (element.TryGetCurrentPattern(InvokePattern.Pattern, out var invokeObj) && invokeObj is InvokePattern invokePattern)
        {
            invokePattern.Invoke();
            invokedBy = "invokePattern";
            return true;
        }

        return false;
    }

    private static void TryFocusOrClick(AutomationElement element)
    {
        try
        {
            element.SetFocus();
            return;
        }
        catch
        {
            // ignore
        }

        var rect = element.Current.BoundingRectangle;
        if (rect.IsEmpty)
        {
            return;
        }

        int cx = (int)Math.Round(rect.X + (rect.Width / 2));
        int cy = (int)Math.Round(rect.Y + (rect.Height / 2));
        InputSimulator.LeftClickAtScreenPoint(cx, cy);
        Thread.Sleep(50);
    }

    private static object DescribeElement(AutomationElement element)
    {
        var current = element.Current;
        var br = current.BoundingRectangle;

        string? processName = null;
        try
        {
            processName = Process.GetProcessById(current.ProcessId).ProcessName;
        }
        catch
        {
            // ignore
        }

        int[]? runtimeId = null;
        try
        {
            runtimeId = element.GetRuntimeId();
        }
        catch
        {
            // ignore
        }

        return new
        {
            name = current.Name,
            automationId = current.AutomationId,
            className = current.ClassName,
            frameworkId = current.FrameworkId,
            controlType = current.ControlType?.ProgrammaticName,
            isEnabled = current.IsEnabled,
            isOffscreen = current.IsOffscreen,
            processId = current.ProcessId,
            processName,
            nativeWindowHandle = current.NativeWindowHandle,
            runtimeId,
            boundingRect = new
            {
                left = br.Left,
                top = br.Top,
                right = br.Right,
                bottom = br.Bottom,
                width = br.Width,
                height = br.Height,
            },
            patterns = new
            {
                invoke = element.TryGetCurrentPattern(InvokePattern.Pattern, out _),
                value = element.TryGetCurrentPattern(ValuePattern.Pattern, out _),
                toggle = element.TryGetCurrentPattern(TogglePattern.Pattern, out _),
                selectionItem = element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out _),
            },
        };
    }

    private static NativeRect GetVirtualDesktopRect()
    {
        const int SM_XVIRTUALSCREEN = 76;
        const int SM_YVIRTUALSCREEN = 77;
        const int SM_CXVIRTUALSCREEN = 78;
        const int SM_CYVIRTUALSCREEN = 79;

        int x = Win32.GetSystemMetric(SM_XVIRTUALSCREEN);
        int y = Win32.GetSystemMetric(SM_YVIRTUALSCREEN);
        int w = Win32.GetSystemMetric(SM_CXVIRTUALSCREEN);
        int h = Win32.GetSystemMetric(SM_CYVIRTUALSCREEN);

        return new NativeRect { Left = x, Top = y, Right = x + w, Bottom = y + h };
    }

    private static byte[] CaptureRectAsPng(NativeRect rect)
    {
        int width = Math.Max(1, rect.Width);
        int height = Math.Max(1, rect.Height);

        using var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bitmap);
        g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);

        using var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }
}
