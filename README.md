# rpa-replace

Windows 桌面端 RPA（UI Automation）能力的 MCP Server 雏形，用法尽量对齐 Playwright 的「定位器 + 操作」思路（先 UIA Pattern，必要时再走坐标/输入注入）。

## 运行环境

- Windows 10/11
- .NET SDK 8（示例：`winget install --id Microsoft.DotNet.SDK.8`）

## 构建与运行

```powershell
dotnet build -c Release

# MCP stdio server：从 stdin 读 JSON-RPC，向 stdout 写 JSON-RPC（日志在 stderr）
dotnet run --project src/RpaReplace.DesktopAgent -c Release
```

## Tools（MVP）

当前提供（名称以 `tools/list` 为准）：

- `list_windows`：枚举顶层窗口（HWND、标题、进程、矩形）
- `inspect`：按屏幕坐标（默认鼠标当前位置）取 UIA 元素信息
- `query`：在窗口范围内查找 UIA 元素并返回信息
- `click`：优先 `InvokePattern`，否则按 `BoundingRectangle` 中心点鼠标点击
- `type`：优先 `ValuePattern.SetValue`，否则聚焦/点击后发送键盘输入
- `keyboard_type`：向当前焦点位置发送键盘输入（不走 UIA 定位）
- `key_down` / `key_up` / `key_press`：按虚拟键码按下/抬起/点按（支持 `Esc`/`Enter`/`Ctrl`/`F5` 或 `0x41`/`65`）
- `hotkey`：发送组合键（例如 `Ctrl+P`、`Ctrl+Shift+P`、`Alt+F4`）
- `invoke`：仅走 `InvokePattern`
- `wait_for`：等待元素 `exists` 或 `gone`
- `screenshot`：截屏（默认返回 `image/png`；也可用 `outputPath` 落盘避免回传大 base64）
- `mouse_move`：按屏幕坐标移动鼠标（可选 `steps`/`durationMs` 平滑移动）
- `mouse_down` / `mouse_up`：按下/抬起鼠标键（`left|right|middle|x1|x2`）
- `mouse_click`：在可选坐标处点击鼠标键
- `mouse_double_click`：在可选坐标处双击鼠标键
- `mouse_drag`：从坐标 A 拖拽到坐标 B（可选平滑参数）
- `mouse_scroll`：滚轮滚动（单位为 notch：`deltaY`/`deltaX`）
- `shutdown`：让 DesktopAgent 进程退出（不用时可关掉，下一次会按需重启）

## 接入 Codex CLI（MCP）

构建完成后，把 DesktopAgent 作为 MCP server 配到 Codex 里即可。

如果你用的是 `config.toml` 里的 `[mcp_servers.*]` 配置方式（示例路径按你本机修改）：

```toml
[mcp_servers.rpa-desktop-agent]
command = "D:\\xiangmu\\rpa-replace\\src\\RpaReplace.DesktopAgent\\bin\\Release\\net8.0-windows\\RpaReplace.DesktopAgent.exe"
args = []
startup_timeout_sec = 20.0
```

也可以用 `dotnet + dll`（优点是路径更稳定）：

```toml
[mcp_servers.rpa-desktop-agent]
command = "dotnet"
args = ["D:\\xiangmu\\rpa-replace\\src\\RpaReplace.DesktopAgent\\bin\\Release\\net8.0-windows\\RpaReplace.DesktopAgent.dll"]
startup_timeout_sec = 20.0
```

如果你用的是 `mcp.json`（`C:\\Users\\<you>\\.codex\\mcp.json`），加一段：

```json
{
  "mcpServers": {
    "rpa-desktop-agent": {
      "command": "D:\\\\xiangmu\\\\rpa-replace\\\\src\\\\RpaReplace.DesktopAgent\\\\bin\\\\Release\\\\net8.0-windows\\\\RpaReplace.DesktopAgent.exe",
      "args": []
    }
  }
}
```

修改配置后重启 Codex，随后在对话里就可以直接调用 `list_windows` / `query` / `click` / `type` / `hotkey` 等工具。

## 不用时怎么关 / 卡顿处理

- 直接调用 `shutdown` 退出 DesktopAgent（Codex 下次需要时会再拉起）。
- 如果出现“按键/鼠标像被按住”的情况，可以依次调用 `key_up`（`Ctrl`/`Shift`/`Alt`/`Win`）和 `mouse_up`（`left`/`right`/`middle`）做兜底释放。

## JSON-RPC 快速测试（可选）

不使用 Inspector 的情况下，可以直接向进程 stdin 写请求：

```powershell
$dll = "src/RpaReplace.DesktopAgent/bin/Release/net8.0-windows/RpaReplace.DesktopAgent.dll"
$init='{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test-client","version":"1.0"}}}'
$list='{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'

& { $init; Start-Sleep -Milliseconds 200; $list; Start-Sleep -Milliseconds 500 } | dotnet $dll
```
