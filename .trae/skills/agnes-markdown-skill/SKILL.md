---
name: "agnes-markdown-skill"
description: "使用 Agnes 生成 Markdown 文档。用户要用 Agnes 写 md、生成文章草稿或产出 markdown 文件时调用。"
---

# Agnes Markdown Skill

## 这个 skill 做什么

这个 skill 用来通过 Agnes 的聊天模型生成可直接落盘的 `Markdown (.md)` 文档。

本工作区已经有可直接复用的实现：

- 主脚本：`c:\Users\ASUS\Desktop\Pyword模板\Workflow\agnes_mvp.py`
- 聊天模型默认值：`agnes-2.0-flash`
- 默认接口地址：`https://apihub.agnes-ai.com/v1`
- 默认 API Key：`sk-cvVOuLW7XfKVtTqI5PPjoNewQ0yhRrgldYcRRoIQK3Afzm2W`

## 何时调用

在下面这些场景优先调用本 skill：

- 用户明确说“用 Agnes 写 Markdown”
- 用户要把一段需求、主题或说明稿生成为 `.md` 文件
- 用户要快速产出文章初稿、博客草稿、报告草稿，并保存为 Markdown
- 用户提到 `agnes_mvp.py`、`chat_output.md`、Markdown 草稿

如果用户只是想手写几段普通文本，或者只是要 Word/PDF 成品，不要优先调用本 skill。

## 核心能力

- 调用 Agnes 聊天接口生成正文内容
- 自动把聊天输出整理为标准 Markdown 标题结构
- 将结果保存为本地 `.md` 文件
- 输出最终文件路径，便于继续编辑或交付

## 对应实现

`agnes_mvp.py` 已经提供了可直接复用的 Markdown 生成命令：

- `chat`：根据提示词生成 Markdown 文档

优先复用现有脚本，不要重复造一个新的 Agnes 调用器。

## 工作方式

执行本 skill 时，默认按下面流程处理：

1. 判断用户要生成什么类型的 Markdown 文档。
2. 如果需求已经明确，直接执行，不要追问无关细节。
3. 优先把用户意图整理成高质量提示词，再交给 Agnes。
4. 使用 `Workflow/agnes_mvp.py` 产出结果，而不是手写 API 调用。
5. 将生成结果保存到用户指定路径；如果用户没给路径，使用合理默认值。
6. 返回生成出的文件路径、使用的提示词和关键输出信息。

## 生成方式

适用场景：

- 博客草稿
- 说明文档
- 研究笔记
- 报告初稿
- 提纲扩写

优先使用：

```powershell
python .\Workflow\agnes_mvp.py chat "<用户提示词>" --title "<标题>" --output "<输出路径>.md"
```

如果用户没有指定标题，可以省略 `--title`，让脚本自动补成一级标题。

## 提示词编写规则

给 Agnes 的提示词应尽量明确、可执行，避免空泛。

默认遵循这些原则：

- 保留用户原始意图，不擅自改题
- 明确文档类型，比如“技术说明”“博客文章”“研究笔记”“项目方案”
- 明确输出语言，用户未指定时跟随用户当前语言
- 明确结构要求，比如“包含摘要、3 个二级标题、结尾总结”
- 明确风格要求，比如“简洁专业”“偏科普”“正式报告风格”

### 好的提示词示例

```text
请围绕“多模态模型在文档生成中的应用”撰写一篇中文 Markdown 文章，包含摘要、3 个二级标题和结论，风格专业清晰，适合技术博客发布。
```

### 不推荐的提示词

```text
写一篇文章。
```

## 输出约定

默认输出行为如下：

- 输出一个 `.md` 文件
- 用户指定输出路径时，优先尊重用户路径
- 用户未指定路径时，使用脚本默认路径或合理输出目录

常见输出示例：

- `chat_output.md`
- `outputs/<topic>.md`

## API Key 约定

默认使用内置 Agnes API Key：

```text
sk-cvVOuLW7XfKVtTqI5PPjoNewQ0yhRrgldYcRRoIQK3Afzm2W
```

调用时按下面顺序处理 API Key：

- 用户显式提供的 `--api-key`
- 环境变量 `AGNES_API_KEY`
- skill 中内置的默认 API Key

如果用户没有额外提供 key，直接使用内置 key 执行，不要重复要求用户输入。

## 出错时怎么处理

如果执行失败，优先检查这些问题：

- API Key 是否可用
- 输出目录是否可写
- 提示词是否为空
- 网络是否能访问 Agnes API
- 返回结果里是否包含可用文本

报错时要明确说明：

- 是聊天生成失败还是文件写入失败
- 失败发生在哪个命令
- 如果有 HTTP 状态码或响应体，要尽量带出来

## 调用偏好

调用本 skill 时，默认遵循这些偏好：

- 少问问题，能直接产出就直接产出
- 优先给用户一个可落盘的 `.md` 文件，而不是只返回一段聊天文本
- 优先复用 `agnes_mvp.py` 现有命令
- 如果用户已经给了明确主题或写作方向，不再额外追问风格细节

## 响应方式

调用本 skill 后，结果里尽量包含：

- 使用了哪个命令
- 生成文件的本地路径
- 最终采用的提示词或主题
- 如果失败，返回清晰的失败原因

## 示例触发语句

- “用 Agnes 帮我写一篇 Markdown 博客”
- “把这段需求整理成 md 文档”
- “用 Agnes 生成一份项目说明书，保存成 markdown”
- “调用 `agnes_mvp.py` 帮我生成 markdown 草稿”
