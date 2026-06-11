---
name: "yaml-word-template"
description: "基于 YAML 配置生成学术规范 Word/PDF 文档。用户要生成论文、研究报告、带图片表格和参考文献的规范文档时调用。"
---

# YAML Word Template

## 这个 skill 做什么

这个 skill 用来把结构化的 YAML 内容转换成格式规范的 `Word (.docx)` 和 `PDF (.pdf)` 文档。

它基于当前工作区里的现成模板与脚本：

- 模板目录：`c:\Users\ASUS\Desktop\Pyword模板\Format_Reference`
- 主脚本：`c:\Users\ASUS\Desktop\Pyword模板\Format_Reference\Temp_General.py`
- 配置模板：`c:\Users\ASUS\Desktop\Pyword模板\Format_Reference\Demo_Config.yaml`
- 内容模板：`c:\Users\ASUS\Desktop\Pyword模板\Format_Reference\Demo_Content.yaml`

它适合生成：

- 毕业论文
- 研究报告
- 技术文档
- 带章节结构的正式材料
- 包含图片、表格、参考文献、上下标的规范化 Word 文档

## 何时调用

在下面这些场景优先调用本 skill：

- 用户希望把一份文章、论文、报告做成规范 `Word` 文档
- 用户希望通过 `YAML` 维护文档结构与样式
- 用户要求自动生成目录、图片图题、表格表题、参考文献
- 用户需要把内容和格式分开管理
- 用户要批量复用一套格式模板去产出多个文档
- 用户明确提到 `Format_Reference`、`Demo_Config.yaml`、`Demo_Content.yaml` 或 `Temp_General.py`

如果用户只是想润色一段文本，而不是生成 Word/PDF 成品，不要优先调用本 skill。

## 核心能力

- 用 `Demo_Config.yaml` 管页面、字体、段落、图表样式
- 用 `Demo_Content.yaml` 管章节和正文内容
- 支持 `一级标题`、`二级标题`、`三级标题`、`正文`、`图片`、`表格`、`全文大标题`
- 支持单图、双图、三图、四图自适应排版
- 支持三线表生成
- 支持 `{{参考文献}}` 自动编号与文末参考文献列表生成
- 支持 `<sub>...</sub>`、`<sup>...</sup>`、`<i>...</i>` 标签
- 输出 `.docx` 和 `.pdf`

## 工作方式

执行本 skill 时，按下面流程工作：

1. 先确认用户要生成的文档类型、标题、作者和输出目录。
2. 优先复用 `Format_Reference` 中的模板，不要从零重新发明 YAML 结构。
3. 如用户要创建新文档，复制一份模板目录或至少复制 `Demo_Config.yaml`、`Demo_Content.yaml` 到目标目录。
4. 按用户需求修改配置文件与内容文件。
5. 图片路径优先写相对路径，便于模板迁移。
6. 用 `python Temp_General.py <配置文件路径> <输出文件路径>` 或直接运行脚本生成文档。
7. 检查 `.docx` 和 `.pdf` 是否成功输出。
8. 如生成失败，优先排查 YAML 结构、图片路径、字体配置、PDF 转换环境。

## 文件职责

### `Demo_Config.yaml`

负责这些内容：

- 文档基本信息
- 页面设置
- 文本样式
- 表格样式
- 图片样式
- 外部内容文件引用

常见关键字段：

```yaml
document_info:
  title: "毕业论文模板"
  subtitle: "基于 YAML 配置的文档生成"
  author: "张三-2022001"
  date: "2025-01-05"
  organization: "某某大学"

page_setup:
  page_width: 21
  page_height: 29.7
  left_margin: 2.54
  right_margin: 2.54
  top_margin: 2.54
  bottom_margin: 2.54

chapter_configs:
  - "Demo_Content.yaml"

content_structure:
  sections: []
```

### `Demo_Content.yaml`

负责这些内容：

- 全文标题
- 各级章节
- 正文段落
- 图片块
- 表格块

常见结构：

```yaml
sections:
  - type: "全文大标题"
    text: "基于人工智能的图像识别研究"

  - type: "一级标题"
    text: "1. 绪论"
    content_paragraphs:
      - text: "本章主要介绍研究背景。"
        style: "正文"

  - type: "图片"
    paths: ["Example_Image/1.jpg", "Example_Image/2.jpg"]
    captions: ["用户界面", "配置界面"]

  - type: "表格"
    table_name: "实验参数表"
    headers: ["参数名称", "参数值", "说明"]
    data_rows:
      - ["测试文档数量", "100", "随机生成的测试文档"]
    caption: "实验参数设置表"
```

## 内容编写规则

### 标题与段落

- `type: "一级标题"`、`"二级标题"`、`"三级标题"` 用于章节结构
- 标题下的正文放在 `content_paragraphs` 中
- 普通独立段落可直接使用 `type: "正文"`

### 图片

- 单图、多图都用 `type: "图片"`
- 图片路径放在 `paths`
- 图题放在 `captions`
- 建议图片与图题数量一致

### 表格

- 使用 `type: "表格"`
- 表头放 `headers`
- 数据放 `data_rows`
- 表题放 `caption`

### 参考文献

在正文中使用双花括号写引用：

```text
本研究基于{{张三. 论文标题[J]. 期刊名, 2024, 10(2): 1-10.}}的方法。
```

处理规则：

- 引用会自动编号
- 连续引用会自动合并成区间
- 文末会自动生成参考文献列表
- 不要手动再写一份参考文献列表

### 特殊格式

使用内嵌标签：

```text
H<sub>2</sub>O
E=mc<sup>2</sup>
<i>in vitro</i>
```

## 执行命令

在 `Format_Reference` 目录内，最基础的执行方式是：

```powershell
python .\Temp_General.py
```

如果使用自定义配置文件与输出路径，优先使用：

```powershell
python .\Temp_General.py .\Demo_Config.yaml .\输出文档.docx
```

如果脚本未自动产出 PDF，需要检查本机 Word/COM 环境，因为当前实现通过 Windows Office 自动转换 PDF。

## 执行时的偏好

调用本 skill 时，默认遵循这些原则：

- 优先保留“配置与内容分离”的设计
- 除非用户要求，否则不要把所有内容塞回单一 YAML
- 优先复用已有样式定义，而不是新造一套字段名
- 新增字段时先保证与现有脚本兼容
- 如果用户要求新格式，尽量通过改 `Demo_Config.yaml` 实现
- 如果用户要求新增内容结构，再考虑扩展 `Temp_General.py`

## 常见任务模板

### 任务 1：按现有模板生成文档

适用场景：

- 用户已经有标题、作者、正文内容
- 用户只想快速生成规范文档

做法：

1. 修改 `Demo_Config.yaml`
2. 修改 `Demo_Content.yaml`
3. 运行 `Temp_General.py`
4. 返回生成结果路径

### 任务 2：为新项目复制一套模板

适用场景：

- 用户要为另一份论文或报告新建模板目录

做法：

1. 复制 `Format_Reference` 到新目录
2. 替换示例内容、图片、作者信息
3. 确认图片路径和输出文件名
4. 运行生成脚本

### 任务 3：扩展模板能力

适用场景：

- 用户要新增封面、摘要、目录、特殊版式、更多 section 类型

做法：

1. 先读 `Temp_General.py`
2. 判断是调整 YAML 即可，还是需要改 Python 逻辑
3. 小改优先局部增量实现
4. 修改后生成一次样例文档验证

## 失败排查清单

- YAML 缩进是否正确
- `chapter_configs` 引用的文件是否存在
- 图片路径是否相对于配置文件目录可解析
- 字体名是否可在 Windows/Word 中使用
- 输出 `.docx` 是否被占用
- 本机是否具备 Word PDF 转换所需的 Office/COM 环境

## 响应方式

调用本 skill 时，输出应尽量包含：

- 采用了哪份配置文件
- 改了哪些 YAML 字段
- 是否修改了 `Temp_General.py`
- 生成出的 `.docx` / `.pdf` 路径
- 若失败，明确指出是配置问题、路径问题还是脚本问题

## 示例触发语句

- “帮我把这篇论文做成规范 Word 文档”
- “用 YAML 配置生成一个毕业论文模板”
- “按 `Format_Reference` 的格式给我出一份报告”
- “帮我往这个 Word 模板里加图片、表格和参考文献”
- “把正文和格式拆开维护，最后导出 docx 和 pdf”
