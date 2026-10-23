<div align="center">

# KenEasy Word Converter

**打开就能用的本地 ODT 与万能文档转 Word 工具** — 完美将 `.odt`、`.fodt`、`.rtf`、`.html`、Markdown、TXT 转换为标准 Microsoft Word (`.docx`) 或纯文本 (`.txt`)。
文件全在你的浏览器本地转换，**零上传、不装软件、不注册、纯私密**。

<br/>

### 👉 点这里直接使用

# [🚀 打开 Word Converter（在线）](https://ngiken.github.io/KenEasy-word/)

**https://ngiken.github.io/KenEasy-word/**

> 上面这个才是**工具网页**。
> 你现在看的 GitHub 页面只是**源代码仓库**，不能当工具点着用。

<br/>

[中文](README.md) · [English](README.en.md)

<br/>

<img alt="Use online" src="https://img.shields.io/badge/%E5%9C%A8%E7%BA%BF%E4%BD%BF%E7%94%A8-ngiken.github.io-fb7299?style=for-the-badge">
<img alt="Version" src="https://img.shields.io/badge/version-0.1.0-0071e3?style=for-the-badge">
<img alt="Privacy" src="https://img.shields.io/badge/privacy-local%20only-00aeec?style=for-the-badge">
<img alt="Tests" src="https://img.shields.io/badge/e2e-32%2F32-27c499?style=for-the-badge">
<img alt="License" src="https://img.shields.io/badge/license-MIT-9aa4b2?style=for-the-badge">

<br/><br/>

<img src="docs/screenshots/hero-light.png" alt="KenEasy Word Converter v0.1.0 浅色主界面" width="1200">

<sub>Apple-inspired 自适应毛玻璃界面 · 离线开箱即用</sub>

</div>

---

## 30 秒上手

1. 打开：**[在线工具](https://ngiken.github.io/KenEasy-word/)**
2. 拖入 `.odt`、`.fodt`、`.rtf`、`.html`、`.md` 或 `.txt` 文件（支持多选批量）
3. 选择预设（标准排版、学术论文缩进、紧凑布局、合并文档、纯文本提取）或微调页边距与字号
4. 拖动 `⠿` 手柄调整文档次序（键盘按 ↑ / ↓ 亦可移动）
5. 点 **开始转换** → 秒级生成标准 `.docx` 或打包为 `.zip` 下载

没有任何登录门槛、没有任何服务器排队、没有任何文件泄露风险。

### 操作演示

<p align="center">
  <img src="docs/screenshots/quick-tour.gif" alt="导入文档、选择版式、一键导出 Word 的三步演示" width="960">
</p>

<p align="center"><strong>① 拖入异构文档　→　② 调整版式与次序　→　③ 本地高保真生成 Word</strong></p>

---

## 这是什么

日常办公中，我们经常遇到 LibreOffice / OpenOffice 的 `.odt`、旧系统导出的 `.rtf`、网页离线 `.html`、开发者文档 `.md` 等非标准文档。找第三方在线转换网站往往要排队、收费、甚至把私密合同与简历上传到云端。

**KenEasy Word Converter** 是一个**纯前端、100% 本地运算**的高保真文档转换器：

| 你想做的事 | 它怎么帮你 |
| --- | --- |
| **ODT / FODT 转 Word** | 解析 XML 结构树，保留各级标题、粗体斜体、颜色、多层列表、表格与**内嵌图片**，生成规范 ECMA-376 OpenXML `.docx` |
| **RTF / 富文本转 Word** | 内建状态机分词器，完美解析字符集、对齐方式与富文本标记 |
| **网页 HTML 转 Word** | 解析 DOM 元素，将 `<h1>`-`<h6>`、段落、引用、表格、超链接清洗转为原生 Word 样式 |
| **Markdown 转 Word** | 纯前端词法解析，支持表格、代码块、任务列表，开箱生成正式排版 |
| **批量合并文档** | 将多份异构文档（如 ODT + RTF + MD）按次序拼接，插入独立分页符，合并为一份完整的 Word 汇编报告 |
| **纯文本快速提取** | 一键去除繁复格式，快速生成结构化的 `.txt` 摘要 |
| **极端隐私与合规场景** | 包含强力 CSP `connect-src 'none'` 安全策略，网络断网亦可全速运行，绝对零数据离机 |

界面支持 **中文 / English** 和 **浅色 / 深色** 苹果风外观，会自动记住偏好；首次打开跟随系统偏好。

---

## 界面预览

| 浅色工作区 | 深色工作区 |
| --- | --- |
| <img src="docs/screenshots/workspace-light.png" alt="浅色设置与文档转换队列"> | <img src="docs/screenshots/workspace-dark.png" alt="深色设置与文档转换队列"> |

<p align="center">
  <img src="docs/screenshots/mobile-light.png" alt="KenEasy Word Converter 手机界面" width="360">
</p>

<p align="center"><sub>桌面端宽屏双栏布局、移动端触控友好自适应单栏。</sub></p>

---

## 支持格式

| 输入格式 | 说明 | 保真支持项 |
| --- | --- | --- |
| **`.odt`** | OpenDocument 文本文件 (LibreOffice / OpenOffice) | 各级标题、粗斜下划线、文字颜色、无序/有序列表、表格行列、内嵌媒体图片 |
| **`.fodt`** | 单文件 Flat XML OpenDocument | 纯 XML 文本，解析同 ODT |
| **`.rtf`** | Rich Text Format 富文本 | 基础文字排版、段落对齐、Unicode 多字节中文 |
| **`.html` / `.htm`** | 网页文章与离线文档 | 标签清洗转换、标题层级、表格、超链接、代码块 |
| **`.md` / `.markdown`** | Markdown 标记文件 | ATX 标题、强调、引用块、GFM 表格、列表缩进 |
| **`.txt` / `.text`** | 纯文本与日志 | 自动识别标题与空行分段，保持换行 |

> **提示**：微软旧版的二进制 `.doc` (Word 97-2003) 结构为复合二进制流（CFBF），若有需求请在 Word / WPS 中另存为 `.docx` 或 `.odt` 后使用本工具。

---

## 排版预设与专业控制

- **标准排版 (Standard)**：适合商务文档与工作报告，默认 A4 纸张、2.54cm 标准边距、现代字体族排版。
- **学术论文 (Academic)**：经典宋体/Times 字体搭配，自动启用中文惯例的**首行缩进 2 字符**。
- **紧凑布局 (Compact)**：窄页边距 (1.27cm)，最大限度容纳表格与密集数据。
- **多文合并 (Merge)**：将队列中所有异构文件无缝拼接为一份完整 `.docx`，每个原文件独立分页。
- **纯文本 (Plain Text)**：提取纯文字结构内容，快速导出 `.txt`。

---

## 技术架构

```text
web/
  index.html       # 语义化 HTML 骨架（配置 CSP connect-src 'none'）
  styles.css       # 响应式毛玻璃微拟物设计系统（暗黑/明亮双套语义变量）
  config.js        # 数据层：格式限制、纸张尺寸度量、排版预设
  i18n.js          # 内容层：全量中英双语国际化
  doc-engine.js    # 核心领域层：ODT/FODT/RTF/HTML/MD 解析器与 OpenXML 生成引擎
  app.js           # 表现层：文件队列、排序拖拽交互、状态机与流水线调度
  vendor/          # 本地固定版本的 JSZip 与 SortableJS
```

- **高保真 OpenXML 打包**：在客户端动态装配 ECMA-376 标准命名空间 XML（`[Content_Types].xml`、`_rels/.rels`、`word/document.xml`、`word/styles.xml`、`word/media/`），生成的 `.docx` 能够在 Microsoft Office 2007 至 365、WPS Office、Pages 等现代文字处理软件中原生秒开。
- **真正的零后端**：所有依赖已打包内嵌，不需要 `npm install` 即可在静态空间托管。

---

## 本地开发与测试

运行实时 E2E 测试套件（包含 32 项自动化断言、多格式解析与 OpenXML 校验）：

```bash
# 执行端到端真实无头 Chromium 回归测试
node scratch/e2e/run.mjs
```

本地预览网页：

```bash
python -m http.server 5173 --directory web
```

在浏览器打开 `http://localhost:5173/` 即可。

---

## 开源协议

本项目采用 [MIT 许可证](LICENSE)。
