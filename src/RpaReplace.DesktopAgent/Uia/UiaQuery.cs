using System.Text.RegularExpressions;
using System.Reflection;
using RpaReplace.DesktopAgent.Interop;
using System.Windows.Automation;

namespace RpaReplace.DesktopAgent.Uia;

internal sealed record UiaWindowQuery(
    string? Title = null,
    string? TitleRegex = null,
    long? Hwnd = null,
    int? ProcessId = null,
    int Index = 0);

internal sealed record UiaElementQuery(
    string? Name = null,
    string? NameRegex = null,
    string? AutomationId = null,
    string? ClassName = null,
    string? ControlType = null,
    int Index = 0);

internal static class UiaQuery
{
    private static readonly Lazy<IReadOnlyDictionary<string, ControlType>> ControlTypeMap =
        new(() =>
        {
            var map = new Dictionary<string, ControlType>(StringComparer.OrdinalIgnoreCase);

            foreach (var field in typeof(ControlType).GetFields(BindingFlags.Public | BindingFlags.Static))
            {
                if (field.FieldType != typeof(ControlType))
                {
                    continue;
                }

                try
                {
                    if (field.GetValue(null) is ControlType ct)
                    {
                        map[field.Name] = ct;
                    }
                }
                catch
                {
                    // ignore
                }
            }

            foreach (var prop in typeof(ControlType).GetProperties(BindingFlags.Public | BindingFlags.Static))
            {
                if (prop.PropertyType != typeof(ControlType))
                {
                    continue;
                }

                try
                {
                    if (prop.GetValue(null) is ControlType ct)
                    {
                        map[prop.Name] = ct;
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return map;
        });

    public static AutomationElement? ResolveWindowRoot(UiaWindowQuery query)
    {
        if (query.Hwnd is not null)
        {
            return AutomationElement.FromHandle(new IntPtr(query.Hwnd.Value));
        }

        var matches = Win32.EnumerateTopLevelWindows(visibleOnly: false, includeEmptyTitles: true)
            .Where(w => WindowMatches(w, query))
            .ToList();

        if (query.Index >= 0 && query.Index < matches.Count)
        {
            return AutomationElement.FromHandle(matches[query.Index].Hwnd);
        }

        var uiaWindows = AutomationElement.RootElement.FindAll(
            TreeScope.Children,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

        var uiaMatches = new List<AutomationElement>();
        for (int i = 0; i < uiaWindows.Count; i++)
        {
            var win = uiaWindows[i];
            if (!WindowMatches(win, query))
            {
                continue;
            }

            uiaMatches.Add(win);
        }

        return query.Index >= 0 && query.Index < uiaMatches.Count ? uiaMatches[query.Index] : null;
    }

    private static bool WindowMatches(AutomationElement window, UiaWindowQuery query)
    {
        var current = window.Current;

        if (query.ProcessId is not null && current.ProcessId != query.ProcessId.Value)
        {
            return false;
        }

        string title = current.Name ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(query.Title))
        {
            if (!title.Contains(query.Title, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(query.TitleRegex))
        {
            var regex = new Regex(query.TitleRegex, RegexOptions.IgnoreCase);
            if (!regex.IsMatch(title))
            {
                return false;
            }
        }

        return true;
    }

    public static AutomationElement? FindElement(AutomationElement root, UiaElementQuery query)
    {
        // Fast path: use native UIA query whenever possible (avoids walking the entire tree).
        if (string.IsNullOrWhiteSpace(query.NameRegex) && query.Index == 0)
        {
            Condition condition = BuildCondition(query);
            return root.FindFirst(TreeScope.Descendants, condition);
        }

        // Slower paths (regex / index>0): avoid FindAll(...) which can be extremely expensive on large UI trees.
        var expectedIndex = query.Index < 0 ? 0 : query.Index;
        Regex? nameRegex = null;
        if (!string.IsNullOrWhiteSpace(query.NameRegex))
        {
            nameRegex = new Regex(query.NameRegex, RegexOptions.IgnoreCase);
        }

        int seen = 0;
        foreach (var el in EnumerateDescendants(root))
        {
            if (!ElementMatches(el, query, nameRegex))
            {
                continue;
            }

            if (seen == expectedIndex)
            {
                return el;
            }

            seen++;
        }

        return null;
    }

    private static bool WindowMatches(WindowInfo window, UiaWindowQuery query)
    {
        if (query.ProcessId is not null && window.ProcessId != query.ProcessId.Value)
        {
            return false;
        }

        string title = window.Title ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(query.Title))
        {
            if (!title.Contains(query.Title, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(query.TitleRegex))
        {
            var regex = new Regex(query.TitleRegex, RegexOptions.IgnoreCase);
            if (!regex.IsMatch(title))
            {
                return false;
            }
        }

        return true;
    }

    private static ControlType? ControlTypeFromString(string value)
    {
        string normalized = value.Trim();
        if (normalized.StartsWith("ControlType.", StringComparison.OrdinalIgnoreCase))
        {
            normalized = normalized["ControlType.".Length..];
        }

        return ControlTypeMap.Value.TryGetValue(normalized, out var ct) ? ct : null;
    }

    private static Condition BuildCondition(UiaElementQuery query)
    {
        var conditions = new List<Condition>();

        if (!string.IsNullOrWhiteSpace(query.Name))
        {
            conditions.Add(new PropertyCondition(AutomationElement.NameProperty, query.Name));
        }

        if (!string.IsNullOrWhiteSpace(query.AutomationId))
        {
            conditions.Add(new PropertyCondition(AutomationElement.AutomationIdProperty, query.AutomationId));
        }

        if (!string.IsNullOrWhiteSpace(query.ClassName))
        {
            conditions.Add(new PropertyCondition(AutomationElement.ClassNameProperty, query.ClassName));
        }

        if (!string.IsNullOrWhiteSpace(query.ControlType))
        {
            var controlType = ControlTypeFromString(query.ControlType);
            if (controlType is not null)
            {
                conditions.Add(new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));
            }
        }

        return conditions.Count switch
        {
            0 => Condition.TrueCondition,
            1 => conditions[0],
            _ => new AndCondition(conditions.ToArray()),
        };
    }

    private static IEnumerable<AutomationElement> EnumerateDescendants(AutomationElement root)
    {
        var walker = TreeWalker.RawViewWalker;

        AutomationElement? current = null;
        try
        {
            current = walker.GetFirstChild(root);
        }
        catch
        {
            yield break;
        }

        if (current is null)
        {
            yield break;
        }

        var stack = new Stack<AutomationElement>();
        stack.Push(current);

        while (stack.Count > 0)
        {
            var el = stack.Pop();
            yield return el;

            AutomationElement? sibling = null;
            try
            {
                sibling = walker.GetNextSibling(el);
            }
            catch
            {
                // ignore
            }

            AutomationElement? child = null;
            try
            {
                child = walker.GetFirstChild(el);
            }
            catch
            {
                // ignore
            }

            if (sibling is not null)
            {
                stack.Push(sibling);
            }

            if (child is not null)
            {
                stack.Push(child);
            }
        }
    }

    private static bool ElementMatches(AutomationElement element, UiaElementQuery query, Regex? nameRegex)
    {
        AutomationElement.AutomationElementInformation current;
        try
        {
            current = element.Current;
        }
        catch
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(query.Name) && !string.Equals(current.Name ?? string.Empty, query.Name, StringComparison.Ordinal))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(query.AutomationId) && !string.Equals(current.AutomationId ?? string.Empty, query.AutomationId, StringComparison.Ordinal))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(query.ClassName) && !string.Equals(current.ClassName ?? string.Empty, query.ClassName, StringComparison.Ordinal))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(query.ControlType))
        {
            var ct = ControlTypeFromString(query.ControlType);
            if (ct is not null && current.ControlType != ct)
            {
                return false;
            }
        }

        if (nameRegex is not null)
        {
            if (!nameRegex.IsMatch(current.Name ?? string.Empty))
            {
                return false;
            }
        }

        return true;
    }
}
