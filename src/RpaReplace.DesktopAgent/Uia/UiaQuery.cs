using System.Text.RegularExpressions;
using System.Diagnostics;
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
    string? NameContains = null,
    string? NameRegex = null,
    string? AutomationId = null,
    string? ClassName = null,
    string? ControlType = null,
    int Index = 0,
    string? Tree = null,
    string? Scope = null,
    int TimeoutMs = 0,
    int MaxDepth = 0,
    int MaxNodes = 0);

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
        Regex? titleRegex = null;
        if (!string.IsNullOrWhiteSpace(query.TitleRegex))
        {
            titleRegex = new Regex(query.TitleRegex, RegexOptions.IgnoreCase);
        }

        if (query.Hwnd is not null)
        {
            return AutomationElement.FromHandle(new IntPtr(query.Hwnd.Value));
        }

        var matches = Win32.EnumerateTopLevelWindows(visibleOnly: false, includeEmptyTitles: true)
            .Where(w => WindowMatches(w, query, titleRegex))
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
            if (!WindowMatches(win, query, titleRegex))
            {
                continue;
            }

            uiaMatches.Add(win);
        }

        return query.Index >= 0 && query.Index < uiaMatches.Count ? uiaMatches[query.Index] : null;
    }

    private static bool WindowMatches(AutomationElement window, UiaWindowQuery query, Regex? titleRegex)
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

        if (titleRegex is not null)
        {
            if (!titleRegex.IsMatch(title))
            {
                return false;
            }
        }

        return true;
    }

    public static AutomationElement? FindElement(AutomationElement root, UiaElementQuery query)
    {
        bool isSlowPath = !string.IsNullOrWhiteSpace(query.NameRegex)
            || !string.IsNullOrWhiteSpace(query.NameContains)
            || query.Index != 0;

        string tree = NormalizeTree(query.Tree, isSlowPath);
        TreeWalker walker = tree switch
        {
            "control" => TreeWalker.ControlViewWalker,
            "content" => TreeWalker.ContentViewWalker,
            _ => TreeWalker.RawViewWalker,
        };

        TreeScope scope = NormalizeScope(query.Scope);
        int maxDepth = query.MaxDepth;
        if (scope == TreeScope.Children)
        {
            maxDepth = 1;
        }

        // Fast path: use native UIA query whenever possible (avoids walking the entire tree).
        if (string.IsNullOrWhiteSpace(query.NameRegex)
            && string.IsNullOrWhiteSpace(query.NameContains)
            && query.Index == 0)
        {
            Condition condition = BuildCondition(query);
            condition = ApplyTreeCondition(condition, tree);
            return root.FindFirst(scope, condition);
        }

        // Slower paths (regex / contains / index>0): avoid FindAll(...) which can be extremely expensive on large UI trees.
        var expectedIndex = query.Index < 0 ? 0 : query.Index;
        Regex? nameRegex = null;
        if (!string.IsNullOrWhiteSpace(query.NameRegex))
        {
            nameRegex = new Regex(query.NameRegex, RegexOptions.IgnoreCase);
        }

        int timeoutMs = query.TimeoutMs;
        int maxNodes = query.MaxNodes;

        if (timeoutMs <= 0 && isSlowPath)
        {
            timeoutMs = 3000;
        }

        if (maxNodes <= 0 && isSlowPath)
        {
            maxNodes = 100_000;
        }

        long? deadlineTicks = GetDeadlineTicks(timeoutMs);

        int seen = 0;
        foreach (var el in EnumerateDescendants(root, walker, maxDepth, maxNodes, deadlineTicks))
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

    private static string NormalizeTree(string? tree, bool isSlowPath)
    {
        if (string.IsNullOrWhiteSpace(tree))
        {
            return isSlowPath ? "control" : "raw";
        }

        string normalized = tree.Trim();
        if (normalized.StartsWith("Tree.", StringComparison.OrdinalIgnoreCase))
        {
            normalized = normalized["Tree.".Length..];
        }

        return normalized.ToLowerInvariant() switch
        {
            "control" => "control",
            "content" => "content",
            "raw" => "raw",
            _ => isSlowPath ? "control" : "raw",
        };
    }

    private static TreeScope NormalizeScope(string? scope)
    {
        if (string.IsNullOrWhiteSpace(scope))
        {
            return TreeScope.Descendants;
        }

        string normalized = scope.Trim();
        if (normalized.StartsWith("TreeScope.", StringComparison.OrdinalIgnoreCase))
        {
            normalized = normalized["TreeScope.".Length..];
        }

        return normalized.ToLowerInvariant() switch
        {
            "children" => TreeScope.Children,
            _ => TreeScope.Descendants,
        };
    }

    private static Condition ApplyTreeCondition(Condition condition, string tree) =>
        tree switch
        {
            "control" => new AndCondition(condition, new PropertyCondition(AutomationElement.IsControlElementProperty, true)),
            "content" => new AndCondition(condition, new PropertyCondition(AutomationElement.IsContentElementProperty, true)),
            _ => condition,
        };

    private static long? GetDeadlineTicks(int timeoutMs)
    {
        if (timeoutMs <= 0)
        {
            return null;
        }

        long now = Stopwatch.GetTimestamp();
        long ticks = (long)Math.Round(timeoutMs / 1000d * Stopwatch.Frequency);
        return now + ticks;
    }

    private static bool WindowMatches(WindowInfo window, UiaWindowQuery query, Regex? titleRegex)
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

        if (titleRegex is not null)
        {
            if (!titleRegex.IsMatch(title))
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
            conditions.Add(new PropertyCondition(AutomationElement.NameProperty, query.Name, PropertyConditionFlags.IgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(query.AutomationId))
        {
            conditions.Add(new PropertyCondition(AutomationElement.AutomationIdProperty, query.AutomationId, PropertyConditionFlags.IgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(query.ClassName))
        {
            conditions.Add(new PropertyCondition(AutomationElement.ClassNameProperty, query.ClassName, PropertyConditionFlags.IgnoreCase));
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
        return EnumerateDescendants(root, TreeWalker.RawViewWalker, maxDepth: 0, maxNodes: 0, deadlineTicks: null);
    }

    private static IEnumerable<AutomationElement> EnumerateDescendants(
        AutomationElement root,
        TreeWalker walker,
        int maxDepth,
        int maxNodes,
        long? deadlineTicks)
    {
        maxDepth = Math.Max(0, maxDepth);
        maxNodes = Math.Max(0, maxNodes);

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
        var depthStack = new Stack<int>();
        stack.Push(current);
        depthStack.Push(1);

        int yielded = 0;

        while (stack.Count > 0 && depthStack.Count > 0)
        {
            var el = stack.Pop();
            int depth = depthStack.Pop();

            if (deadlineTicks is not null && Stopwatch.GetTimestamp() >= deadlineTicks.Value)
            {
                yield break;
            }

            yield return el;
            yielded++;

            if (maxNodes > 0 && yielded >= maxNodes)
            {
                yield break;
            }

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
                if (maxDepth == 0 || depth < maxDepth)
                {
                    child = walker.GetFirstChild(el);
                }
            }
            catch
            {
                // ignore
            }

            if (sibling is not null)
            {
                stack.Push(sibling);
                depthStack.Push(depth);
            }

            if (child is not null)
            {
                stack.Push(child);
                depthStack.Push(depth + 1);
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

        if (!string.IsNullOrWhiteSpace(query.Name) && !string.Equals(current.Name ?? string.Empty, query.Name, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(query.NameContains) && (current.Name ?? string.Empty).IndexOf(query.NameContains, StringComparison.OrdinalIgnoreCase) < 0)
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(query.AutomationId) && !string.Equals(current.AutomationId ?? string.Empty, query.AutomationId, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(query.ClassName) && !string.Equals(current.ClassName ?? string.Empty, query.ClassName, StringComparison.OrdinalIgnoreCase))
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
