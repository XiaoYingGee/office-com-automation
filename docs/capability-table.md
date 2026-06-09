# Excel 自动化能力表

> 四种后端的操作覆盖矩阵。
> ✅ 已实现 ｜ ⬜ 待实现（COM 支持，代码未写）

---

## 文件/生命周期

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 打开工作簿 | open | ✅ | ✅ | ✅ | ✅ |
| 新建工作簿 | create | ✅ | ✅ | ✅ | ✅ |
| 保存 | save | ✅ | ✅ | ⬜ | ⬜ |
| 关闭 | close | ✅ | ✅ | ✅ | ✅ |

## 单元格读写

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 写单元格 | write_cell | ✅ | ✅ | ✅ | ✅ |
| 读单元格 | read_cell | ✅ | ✅ | ✅ | ✅ |
| 批量写区域 | write_range | ✅ | ✅ | ✅ | ✅ |
| 批量读区域 | read_range | ✅ | ✅ | ✅ | ✅ |

## 区域操作

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 清除内容 | clear_range | ✅ | ✅ | ✅ | ✅ |
| 合并单元格 | merge_cells | ✅ | ✅ | ✅ | ✅ |
| 取消合并 | unmerge_cells | ✅ | ✅ | ✅ | ✅ |
| 复制区域值 | copy_values | ⬜ | ⬜ | ⬜ | ⬜ |
| 选择性粘贴 | paste_special | ⬜ | ⬜ | ⬜ | ⬜ |
| 自动填充 | auto_fill | ⬜ | ⬜ | ⬜ | ⬜ |

## 格式化

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 字体格式 | set_format | ✅ | ✅ | ✅ | ✅ |
| 背景色 | set_format | ✅ | ✅ | ✅ | ✅ |
| 数字格式 | set_format | ✅ | ✅ | ✅ | ✅ |
| 对齐/换行 | set_format | ✅ | ✅ | ✅ | ✅ |
| 边框 | set_border | ⬜ | ⬜ | ⬜ | ⬜ |
| 条件格式 | add_conditional_format | ⬜ | ⬜ | ⬜ | ⬜ |
| 删除条件格式 | clear_conditional_format | ⬜ | ⬜ | ⬜ | ⬜ |

## 行列结构

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 插入行 | insert_rows | ✅ | ✅ | ✅ | ✅ |
| 删除行 | delete_rows | ✅ | ✅ | ✅ | ✅ |
| 插入列 | insert_cols | ✅ | ✅ | ✅ | ✅ |
| 删除列 | delete_cols | ✅ | ✅ | ✅ | ✅ |
| 自适应列宽 | autofit_columns | ✅ | ✅ | ⬜ | ⬜ |
| 设置行高 | set_row_height | ⬜ | ⬜ | ⬜ | ⬜ |
| 设置列宽 | set_col_width | ⬜ | ⬜ | ⬜ | ⬜ |
| 行列分组 | group_rows / group_cols | ⬜ | ⬜ | ⬜ | ⬜ |
| 取消分组 | ungroup_rows / ungroup_cols | ⬜ | ⬜ | ⬜ | ⬜ |
| 隐藏行列 | hide_rows / hide_cols | ⬜ | ⬜ | ⬜ | ⬜ |
| 显示行列 | unhide_rows / unhide_cols | ⬜ | ⬜ | ⬜ | ⬜ |

## 工作表管理

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 新增工作表 | add_sheet | ✅ | ✅ | ✅ | ✅ |
| 重命名 | rename_sheet | ✅ | ✅ | ✅ | ✅ |
| 删除工作表 | delete_sheet | ✅ | ✅ | ✅ | ✅ |
| 复制工作表 | copy_sheet | ⬜ | ⬜ | ⬜ | ⬜ |
| 移动工作表 | move_sheet | ⬜ | ⬜ | ⬜ | ⬜ |
| 保护工作表 | protect_sheet | ⬜ | ⬜ | ⬜ | ⬜ |
| 取消保护 | unprotect_sheet | ⬜ | ⬜ | ⬜ | ⬜ |
| 冻结窗格 | freeze_panes | ⬜ | ⬜ | ⬜ | ⬜ |
| 取消冻结 | unfreeze_panes | ⬜ | ⬜ | ⬜ | ⬜ |

## 数据操作

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 排序 | sort_range | ✅ | ✅ | ⬜ | ✅ |
| 自动筛选 | auto_filter | ✅ | ✅ | ⬜ | ⬜ |
| 查找替换 | find_replace | ✅ | ✅ | ⬜ | ✅ |
| 重算公式 | calculate | ✅ | ✅ | ⬜ | ✅ |
| 数据验证 | add_validation | ⬜ | ⬜ | ⬜ | ⬜ |
| 删除验证 | clear_validation | ⬜ | ⬜ | ⬜ | ⬜ |
| 命名区域 | add_named_range | ⬜ | ⬜ | ⬜ | ⬜ |
| 删除命名区域 | delete_named_range | ⬜ | ⬜ | ⬜ | ⬜ |

## 结构查看

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 查看工作簿结构 | inspect | ✅ | ✅ | ✅ | ✅ |
| 查看工作表数据 | inspect_sheet | ✅ | ✅ | ⬜ | ⬜ |

## 图表

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 创建图表 | add_chart | ⬜ | ⬜ | ⬜ | ⬜ |
| 删除图表 | delete_chart | ⬜ | ⬜ | ⬜ | ⬜ |
| 修改图表类型 | modify_chart | ⬜ | ⬜ | ⬜ | ⬜ |
| 设置图表标题 | set_chart_title | ⬜ | ⬜ | ⬜ | ⬜ |

## 图片

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 插入图片 | add_picture | ⬜ | ⬜ | ⬜ | ⬜ |
| 删除图片 | delete_picture | ⬜ | ⬜ | ⬜ | ⬜ |

## 批注/超链接

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 添加批注 | add_comment | ⬜ | ⬜ | ⬜ | ⬜ |
| 删除批注 | delete_comment | ⬜ | ⬜ | ⬜ | ⬜ |
| 添加超链接 | add_hyperlink | ⬜ | ⬜ | ⬜ | ⬜ |
| 删除超链接 | delete_hyperlink | ⬜ | ⬜ | ⬜ | ⬜ |

## 导出/打印

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 导出 PDF | export_pdf | ✅ | ✅ | ⬜ | ✅ |
| 导出图片 | export_image | ⬜ | ⬜ | ⬜ | ⬜ |
| 打印设置 | set_page_setup | ⬜ | ⬜ | ⬜ | ⬜ |

## 宏/VBA

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 运行宏 | run_macro | ⬜ | ⬜ | ⬜ | ⬜ |

## 数据透视表

| 操作 | action | pywin32 | VBA | PyAddin | C# Addin |
|------|--------|:---:|:---:|:---:|:---:|
| 创建透视表 | add_pivot_table | ⬜ | ⬜ | ⬜ | ⬜ |
| 刷新透视表 | refresh_pivot | ⬜ | ⬜ | ⬜ | ⬜ |

---

## 统计

| 后端 | 已实现 | 待实现 | 覆盖率 |
|------|:---:|:---:|:---:|
| pywin32 | 24 | 38 | 39% |
| VBA | 24 | 38 | 39% |
| PyAddin | 18 | 44 | 29% |
| C# Addin | 22 | 40 | 35% |

## 说明

- 所有四种后端均通过 COM 接口访问同一个 Excel 进程，理论上能力完全一致
- ⬜ 表示 COM 支持该操作，但代码尚未实现对应的 action dispatch
- PyAddin 的 ⬜ 操作可通过在 `dispatch_action()` 中添加 case 立即获得（复用 ExcelCOM 方法）
- C# Addin 的 ⬜ 操作需要在 `Connect.cs` 的 `ExecuteAction` switch 中添加对应分支
- **性能差异只在 IPC 边界**：in-process（VBA/PyAddin/C# Addin）零 IPC，out-of-process（pywin32）逐属性往返
- **并发限制**：COM STA 模型，同一时刻只能有一个线程操作 Excel
