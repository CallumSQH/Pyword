================================================================================
                  Word 文档模板生成器 - YAML 配置使用说明
                              版本 2.2
================================================================================

一、简介
================================================================================

本工具用于通过 YAML 配置文件快速生成符合规范的 Word 文档，并同步导出 PDF。
适合论文、课程报告、项目说明书、研究报告等需要统一格式的场景。

当前目录提供两套 demo，并且已经拆分到两个独立子文件夹中：

1. Demo 1：基础单内容示例
   - 目录：Demo_01_Basic/
   - 核心文件：Demo_01_Basic/Demo_Config.yaml + Demo_01_Basic/Demo_Content.yaml
   - 适合：第一次上手、内容量不大、希望快速改一份完整示例
   - 说明：这套 demo 保持基础用法，适合先熟悉标题、图片、表格、引用等能力

2. Demo 2：总配置 + 多章节拆分示例
   - 目录：Demo_02_Split/
   - 核心文件：Demo_02_Split/Demo_Split_Config.yaml + Demo_02_Split/Demo_Split_Chapter_*.yaml
   - 适合：论文、长文档、多人协作、按章维护的写作方式
   - 说明：这套 demo 把“如何拆分章节”本身写成了指导说明，可直接照着扩展


二、快速开始
================================================================================

【运行 Demo 1：基础单文件内容示例】

    python Temp_General.py Demo_01_Basic/Demo_Config.yaml

输出文件：
    - Format_Reference.docx
    - Format_Reference.pdf

【运行 Demo 2：章节拆分示例】

    python Temp_General.py Demo_02_Split/Demo_Split_Config.yaml

输出文件：
    - Format_Reference_Split_Demo.docx
    - Format_Reference_Split_Demo.pdf


三、目录结构
================================================================================

├── Demo_01_Basic/
│   ├── Demo_Config.yaml                 # Demo 1：基础配置
│   └── Demo_Content.yaml                # Demo 1：正文内容
├── Demo_02_Split/
│   ├── Demo_Split_Config.yaml           # Demo 2：总配置入口
│   ├── Demo_Split_Chapter_00_FrontMatter.yaml
│   ├── Demo_Split_Chapter_01_Intro.yaml
│   ├── Demo_Split_Chapter_02_Structure.yaml
│   ├── Demo_Split_Chapter_03_Practice.yaml
│   └── Demo_Split_Chapter_04_Collaboration.yaml
├── Example_Image/                       # 示例图片资源
├── Temp_General.py                      # 主程序
└── README_YAML.txt                      # 当前说明文档


四、先理解两个 demo 的区别
================================================================================

【Demo 1：适合快速开始】

特点：
- 一个配置文件负责样式与页面参数
- 一个内容文件负责正文 sections
- 理解成本最低，最快看到生成效果

推荐场景：
- 第一次使用这个模板
- 文档内容不多
- 希望先熟悉标题、图片、表格、引用等基础写法

【Demo 2：适合真实项目】

特点：
- 一个总配置文件统一入口
- 多个章节文件按顺序拼装
- 更接近论文、研究报告、长文档的组织方式

推荐场景：
- 文档天然按章组织
- 一份文档会长期维护
- 多人协作，每人负责不同章节
- 希望复用同一套样式，但不断更换正文内容


五、最常用的写法
================================================================================

【1】基础配置文件负责什么

基础配置一般负责：
- document_info：标题、作者、日期等文档信息
- page_setup：页面大小、边距、页眉页脚距离等
- text_styles：正文、各级标题、参考文献样式
- table / image：表格与图片排版规则

【2】内容文件负责什么

内容文件一般只放 sections。
也就是说，正文内容、章节标题、图片、表格、说明段落都写在这里。

【3】章节拆分文件负责什么

在 Demo 2 里，每个章节文件仍然只放 sections，但只负责自己的那一章。
例如：
- 01 文件负责绪论
- 02 文件负责结构设计
- 03 文件负责写作实践
- 04 文件负责协作与发布建议

总配置通过 chapter_configs 把它们按顺序拼起来。


六、Demo 2 的核心配置方式
================================================================================

Demo 2 使用两种关键能力：

【1】import_base_config

作用：复用已有基础配置。

示例：
    import_base_config: "../Demo_01_Basic/Demo_Config.yaml"

这表示：
- 页面设置沿用 Demo 1
- 标题、正文、图题、表格样式沿用 Demo 1
- 只在新 demo 中覆盖必要的标题、作者、输出文件名和章节列表

【2】chapter_configs

作用：按顺序装配多个章节文件。

示例：
    chapter_configs:
      - "Demo_Split_Chapter_00_FrontMatter.yaml"
      - "Demo_Split_Chapter_01_Intro.yaml"
      - "Demo_Split_Chapter_02_Structure.yaml"
      - "Demo_Split_Chapter_03_Practice.yaml"
      - "Demo_Split_Chapter_04_Collaboration.yaml"

建议：
- 使用 00、01、02 这类可排序前缀
- 文件名尽量能看出章节职责
- 新增章节时，只需要新增 YAML 文件并追加到 chapter_configs

【3】generation_config

作用：给不同 demo 指定不同输出文件名，避免互相覆盖。

示例：
    generation_config:
      output_filename: "Format_Reference_Split_Demo.docx"


七、sections 的常见写法
================================================================================

【1】一级标题 / 二级标题 / 三级标题

    - type: "一级标题"
      text: "1. 绪论"
      content_paragraphs:
        - text: "本章介绍研究背景与目标。"
          style: "正文"

    - type: "二级标题"
      text: "1.1 研究背景"
      content_paragraphs:
        - text: "这一节补充说明背景。"
          style: "正文"

【2】正文段落

    - type: "正文"
      text: "这是一段普通正文。"

如果需要软换行，可直接在 text 中使用 \n。

【3】图片

如果图片文件与当前 YAML 在同一目录层级，可以直接写：

    - type: "图片"
      paths: ["Example_Image/1.jpg"]
      captions: ["单图示例"]

如果图片目录在当前 YAML 的上一级目录，例如 Demo_02_Split 下的章节文件引用根目录图片，应写成：

    - type: "图片"
      paths: ["../Example_Image/1.jpg"]
      captions: ["单图示例"]

并排图片示例：

    - type: "图片"
      paths: ["../Example_Image/2.jpg", "../Example_Image/3.jpg"]
      captions: ["左图说明", "右图说明"]

【4】表格

    - type: "表格"
      table_name: "参数表"
      headers: ["字段", "说明"]
      data_rows:
        - ["title", "文档标题"]
        - ["author", "作者信息"]
      caption: "字段说明表"
      style: "居中"


八、引用、上下标与格式标记
================================================================================

【1】参考文献引用

在正文中写：

    这是一个带引用的说明{{张三. 示例文献[J]. 示例期刊, 2025, 1(1): 1-5.}}

程序会自动：
- 给正文中的引用编号
- 在文末生成参考文献列表
- 复用同一条引用的编号

【2】上下标

下标：
    H<sub>2</sub>O

上标：
    E=mc<sup>2</sup>

混合：
    SO<sub>4</sub><sup>2-</sup>

【3】斜体

    <i>这是斜体文本</i>

这些能力在 Demo 2 中仍然可用，说明“拆章节”不会影响正文标记能力。


九、推荐的工作流
================================================================================

【如果你是第一次使用】

1. 先运行 Demo 1
2. 改一两段正文，熟悉基础语法
3. 确认图片、表格、引用都能正常工作

【如果你要做长文档】

1. 以 Demo 2 为起点
2. 复制 Demo_02_Split/Demo_Split_Config.yaml 改成你的总配置
3. 按章节复制 Demo_02_Split/Demo_Split_Chapter_*.yaml
4. 每新增一章，就把文件名追加到 chapter_configs
5. 每次修改后重新生成 docx / pdf 检查排版

【如果你要多人协作】

建议规则：
- 一个人维护总配置
- 每个人负责自己章节文件
- 样式统一在基础配置中调整
- 图片尽量使用相对路径


十、常见问题
================================================================================

【1】为什么我改了 YAML，但输出文件没变化？

请确认你运行的是正确的配置文件，例如：

    python Temp_General.py Demo_02_Split/Demo_Split_Config.yaml

如果你运行的是 Demo 1，它不会读取 Demo 2 的章节文件。

【2】为什么第二套 demo 没有覆盖第一套输出？

因为 Demo 2 设置了：

    generation_config:
      output_filename: "Format_Reference_Split_Demo.docx"

这样可以避免不同 demo 互相覆盖产物。

【3】图片路径应该怎么写？

规则很简单：相对路径总是相对于“当前正在读取的 YAML 文件”来解析。

例如：
- Demo_01_Basic/Demo_Content.yaml 可以直接写 Example_Image/1.jpg 吗？不建议，因为图片目录不在同级
- Demo_02_Split/Demo_Split_Chapter_01_Intro.yaml 如果要引用根目录图片，应写 ../Example_Image/1.jpg

换句话说，进入了子文件夹以后，引用上一级目录资源时要加 ../。

【4】我想继续增加章节，需要改程序吗？

通常不需要。
你只需要：
- 新建一个章节 YAML
- 把文件名追加到 chapter_configs
- 重新运行生成命令


十一、结论
================================================================================

这套模板现在提供两条清晰路径：

- 想快速看效果，用 Demo_01_Basic/
- 想做长文档、论文或协作项目，用 Demo_02_Split/

也就是说，第一套 demo 负责“最快上手”，第二套 demo 负责“长期可维护”。
如果你后续继续扩展，优先复制 Demo 2 的结构，而不是把所有正文重新塞回一个大 YAML。
