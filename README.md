# Pyword

Pyword 是一个用 YAML 驱动的 Word 文档生成工具。

它适合把“内容”和“格式”分开维护：你用 YAML 写标题、正文、图片、表格、引用，程序负责按固定模板生成 `.docx`，并在可用时继续导出 `.pdf`。

这个项目尤其适合下面几类场景：

- 论文、课程报告、项目说明书、研究报告
- 需要统一格式的长文档
- 想把一份文档拆成多个章节单独维护
- 希望把“Vibe Writing”变成可复用、可批量生成的文档流程

## 它能做什么

- 用 YAML 定义文档内容和版式
- 生成 Word 文档
- 自动导出 PDF
- 支持标题层级、目录、页眉页脚、页边距等常见排版需求
- 支持图片、表格、参考文献引用
- 支持 `<sub>`、`<sup>`、`<i>` 这类轻量文本标记
- 支持“总配置 + 多章节文件”的拆分写法

## 先看两个 Demo

项目里已经带了两套 demo，建议按顺序看。

### Demo 1：快速上手

目录：`Demo_01_Basic/`

适合：

- 第一次使用 Pyword
- 想先理解最基础的 YAML 写法
- 内容还不多，只需要一份完整文档

核心文件：

- `Demo_01_Basic/Demo_Config.yaml`
- `Demo_01_Basic/Demo_Content.yaml`

运行：

```bash
python Temp_General.py Demo_01_Basic/Demo_Config.yaml
```

输出：

- `Format_Reference.docx`
- `Format_Reference.pdf`

### Demo 2：章节拆分

目录：`Demo_02_Split/`

适合：

- 论文、长文档
- 多人协作
- 按章节维护内容
- 想复用同一套样式，但把正文拆开管理

核心文件：

- `Demo_02_Split/Demo_Split_Config.yaml`
- `Demo_02_Split/Demo_Split_Chapter_*.yaml`

运行：

```bash
python Temp_General.py Demo_02_Split/Demo_Split_Config.yaml
```

输出：

- `Format_Reference_Split_Demo.docx`
- `Format_Reference_Split_Demo.pdf`

## 快速开始

### 1. 准备环境

推荐环境：

- Windows
- Python 3.10+
- 已安装 Microsoft Word

这个项目当前是 Windows 优先的实现，因为 PDF 导出依赖 Word 的 COM 接口。

### 2. 安装依赖

先安装 Python 依赖：

```bash
pip install pyyaml python-docx docx2pdf pywin32 comtypes
```

### 3. 运行 Demo

进入项目目录后，任选一个 demo：

```bash
python Temp_General.py Demo_01_Basic/Demo_Config.yaml
```

或者：

```bash
python Temp_General.py Demo_02_Split/Demo_Split_Config.yaml
```

### 4. 查看输出文件

生成成功后，你会在当前目录看到：

- `.docx` 文件
- 对应的 `.pdf` 文件

如果 `.docx` 生成成功但 `.pdf` 没导出，通常是 Word 环境或 COM 调用不可用。

## 目录结构

```text
.
├── Demo_01_Basic/
│   ├── Demo_Config.yaml
│   └── Demo_Content.yaml
├── Demo_02_Split/
│   ├── Demo_Split_Config.yaml
│   ├── Demo_Split_Chapter_00_FrontMatter.yaml
│   ├── Demo_Split_Chapter_01_Intro.yaml
│   ├── Demo_Split_Chapter_02_Structure.yaml
│   ├── Demo_Split_Chapter_03_Practice.yaml
│   └── Demo_Split_Chapter_04_Collaboration.yaml
├── Example_Image/
├── Temp_General.py
├── README.md
└── README_YAML.txt
```

## 最核心的概念

### 基础配置

基础配置文件一般负责这些内容：

- 文档标题、作者、日期
- 页边距、页眉页脚、纸张尺寸
- 正文、标题、图题、表格等样式
- 默认输出文件名

### 内容文件

内容文件主要放 `sections`，也就是正文结构本身：

- 一级标题
- 二级标题
- 段落
- 图片
- 表格
- 参考文献

### 章节拆分

如果文档比较长，可以把每一章拆成单独的 YAML 文件，再由总配置统一装配。

Demo 2 用到两个关键字段：

- `import_base_config`：复用基础样式配置
- `chapter_configs`：按顺序拼接多个章节文件

这就是 Pyword 里“可拆分写作”的核心。

## 你会经常改的地方

如果你只想快速改出自己的文档，通常只需要改这些文件：

- 改样式：`Demo_01_Basic/Demo_Config.yaml`
- 改正文：`Demo_01_Basic/Demo_Content.yaml`
- 改章节式长文档：`Demo_02_Split/` 下的各个章节文件

## 常见能力

Pyword 当前已经支持这些常见写法：

- 标题层级：`一级标题`、`二级标题`、`三级标题`
- 图片插入与图注
- 表格与表题
- 文内参考文献引用
- 下标：`H<sub>2</sub>O`
- 上标：`E=mc<sup>2</sup>`
- 斜体：`<i>italic</i>`

## 常见问题

### PDF 为什么导不出来？

先确认两件事：

- 本机安装了 Microsoft Word
- 当前环境允许 Word 的 COM 自动化调用

如果这两项有问题，通常会出现“Word 成功，PDF 失败”的情况。

### 图片路径应该怎么写？

图片路径是相对于“当前 YAML 文件”解析的。

例如：

- `Demo_01_Basic/` 里的内容文件引用上一级图片目录，要写 `../Example_Image/1.jpg`
- `Demo_02_Split/` 里的章节文件同样要写 `../Example_Image/...`

### 我应该从哪个 demo 开始？

建议这样选：

- 想先跑通流程：从 Demo 1 开始
- 想写论文或长文档：从 Demo 2 开始

## 详细 YAML 说明

如果你想继续看字段含义、章节拆分方式、图片路径规则和更多示例，请看：

- `README_YAML.txt`

## 一个推荐工作流

如果你准备把它真正用起来，推荐顺序是：

1. 先运行 Demo 1，确认环境没问题
2. 把 Demo 1 的正文换成你自己的内容
3. 当文档开始变长，再切到 Demo 2 的拆章节结构
4. 固定样式后，长期只改 YAML，不改 Python 主程序

这套流程的好处很直接：样式稳定，内容可拆，协作也更轻松。
