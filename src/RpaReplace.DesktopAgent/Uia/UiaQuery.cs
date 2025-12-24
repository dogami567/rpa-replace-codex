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
            var regex = new Regex(query.TitleRegex, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            if (!regex.IsMatch(title))
            {
                return false;
            }
        }

        return true;
    }

    public static AutomationElement? FindElement(AutomationElement root, UiaElementQuery query)
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

        Condition condition = conditions.Count switch
        {
            0 => Condition.TrueCondition,
            1 => conditions[0],
            _ => new AndCondition(conditions.ToArray()),
        };

        if (string.IsNullOrWhiteSpace(query.NameRegex) && query.Index == 0)
        {
            return root.FindFirst(TreeScope.Descendants, condition);
        }

        if (!string.IsNullOrWhiteSpace(query.NameRegex))
        {
            var regex = new Regex(query.NameRegex, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var results = root.FindAll(TreeScope.Descendants, condition);
            var filtered = new List<AutomationElement>();
            for (int i = 0; i < results.Count; i++)
            {
                var el = results[i];
                if (!regex.IsMatch(el.Current.Name ?? string.Empty))
                {
                    continue;
                }

                filtered.Add(el);
            }

            return query.Index >= 0 && query.Index < filtered.Count ? filtered[query.Index] : null;
        }

        var all = root.FindAll(TreeScope.Descendants, condition);
        return query.Index >= 0 && query.Index < all.Count ? all[query.Index] : null;
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
            var regex = new Regex(query.TitleRegex, RegexOptions.Compiled | RegexOptions.IgnoreCase);
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
}
