# CLAUDE.md — Excel Automation Project

## MCP Servers

本项目配置了 3 个 Excel MCP Server（见 `.claude/settings.json`）：

| Server | 状态 | 执行位置 | CodeAct |
|--------|------|---------|---------|
| `excel-pywin32` | ✅ 已启用 | 跨进程 COM | Python exec() |
| `excel-addin` | ✅ 已启用 | 进程内 (PyAddin) | Python exec() via bridge |
| `excel-csharp` | ❌ disabled (需 dotnet build) | .NET COM | CSharpCodeProvider |

## 使用方式

### 基本工作流

1. `excel_open(path)` — 打开工作簿
2. `excel_inspect()` — 查看结构
3. `excel_execute_code(code)` — 写代码操作（最灵活）
4. `excel_save()` — 保存

### execute_code 写法

代码中可用变量：`app`, `wb`, `ws`, `result`

```python
# 写值
ws.Range("A1").Value2 = "Hello"

# 批量写
ws.Range("A1:C3").Value2 = ((1,2,3),(4,5,6),(7,8,9))

# 格式化（颜色是 BGR！红=0x0000FF）
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Color = 0x0000FF

# 读数据
data = ws.Range("A1:E10").Value2
result = {"rows": len(data)}

# 设置 result 变量返回数据给调用者
result = "操作完成"
```

### execute_actions 写法

```json
[
  {"action": "write_cell", "sheet": "Sheet1", "cell": "A1", "value": 42},
  {"action": "set_format", "sheet": "Sheet1", "range": "A1", "params": {"bold": true}}
]
```

## 注意事项

- 索引从 1 开始（COM 约定）
- 颜色是 BGR（不是 RGB）：红=0x0000FF, 蓝=0xFF0000
- `excel-addin` 需要先注册 add-in 并手动启动 Excel：
  ```
  python python_addin/excel_pyaddin.py  # 注册
  # 然后手动打开 Excel
  ```
- COM 对象模型参考见 `docs/excel-com-reference.md`
