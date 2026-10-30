# 样式检查脚本

`scripts/inspect_pptx_style.py` 只读 PPTX，输出 JSON 清单。它辅助定位格式，不制作幻灯片，也不代替渲染验收。不运行脚本也可以通过现有工具完成模板检查。

## 运行

Python 3.10 或更高版本，安装依赖：

```sh
python -m pip install -r requirements.txt
python scripts/inspect_pptx_style.py "template.pptx"
python scripts/inspect_pptx_style.py "template.pptx" --slides 2,4-6 --output "style-report.json"
```

路径相对于技能目录；从其他目录运行时传入完整脚本路径。页码从 1 起，范围包含两端，重复页码保留第一次出现的顺序。无 `--slides` 时统计全稿、不列元素详情；指定后，仅选中页参与样式汇总，并按指定顺序列详情，全稿页摘要仍保留。优先选择同类内容页，避免封面、目录混入正文字号统计。

标准输出及 `--output` 文件均为 UTF-8。错误写入标准错误并返回非零退出码；源文件不会保存或覆盖，输出路径不能等同于源文件。

## 报告含义

- `format_version: 2`：新版格式。原有的页摘要、区域样式和元素详情仍保留；`--slides` 现在也限定样式汇总范围。主题颜色使用原始 XML 名称，例如 `theme:accent1`，并保留亮度、透明度等修饰；HSL/scRGB 保留原始通道属性，不转成已解析的 RGB。
- `font` / `font_east_asian` / `font_complex_script`：分别为 Latin、东亚及复杂文字的显式字体。`+mn-lt` 等值是主题引用，不是已解析的字体名。
- `explicit_text_styles`：run 或动态字段中直接声明的字号、字重、颜色等。`<inherited>` 表示此处未声明；不能当成默认值。
- `paragraph_defaults`：段落自己的默认文字格式和对齐设置，独立列出，不冒充 run 格式。段落、表格行列从 1 起，缩进 `level` 从 0 起。
- `path` / `parent_path`：由形状 ID 组成的组合层级。`shape_count` 是顶层对象数，`total_shape_count` 包括组合及其后代；文本和图片计数包含后代。
- `box`：相对于整页的百分比名义边界；普通对象不计算旋转后的外接框。未旋转组合支持平移、缩放、嵌套与翻转。
- `coordinate_status`：`slide` 表示边界可用。旋转组、退化组及缺失几何的对象用其他状态标明，不能解析的 `box` 为 `null`。
- `region`：按元素顶部位置划分的粗略区域提示，不是语义识别或规定的三层高度。未知几何归入 `unresolved`。
- `table_cells`：单元格文本、合并状态、显式格式及段落默认值。汇总时跳过被合并覆盖的单元格，不重复统计内容。

## 核对边界

脚本不完整解析母版、布局、占位符、主题继承或字体替换，也不读取图内文字、图表内部文字、SmartArt、表格主题样式或只在母版上出现的装饰。需要精确匹配时，继续检查对应源文件对象、布局/母版与 `ppt/theme/`、`ppt/slideMasters/`、`ppt/slideLayouts/` 中的 OOXML，并查看渲染结果。源格式和实际渲染效果分别核对，不能从区域统计直接推断本页语义或精确外观。
