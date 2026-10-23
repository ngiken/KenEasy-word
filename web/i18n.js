/* KenEasy Word Converter — isolated bilingual copy and translation service. */
(function (global) {
  "use strict";

  var COPY = {
    zh: {
      pageTitle: "KenEasy Word Converter — 本地 ODT 与万能文档转 Word / TXT",
      pageDesc: "拖入 ODT、FODT、RTF、HTML、Markdown 或文本文件，在浏览器本地高保真转换为 Microsoft Word (.docx) 或 TXT。无需安装、无需上传、零隐私泄露。",
      skipToMain: "跳到主要内容",
      eyebrow: "A CALMER WORD WORKFLOW",
      heroTitle: "从各种复杂文档，到一份整洁 Word。",
      tagline: "拖入、排序、自定义版式，一键转换。所有解析与排版都在你的设备上安静进行。",
      badgeNoInstall: "无需安装",
      badgeLocal: "本地处理",
      badgeFidelity: "高保真 ODT",
      badgeMixed: "多格式支持",
      utilityAria: "工具导航",
      featuresAria: "产品特点",
      guaranteesAria: "产品保障",
      presetsAria: "版式预设",
      workflowAria: "转换步骤",
      workflowImport: "导入文件",
      workflowTune: "设置版式",
      workflowExport: "生成 Word",
      workflowStepLabel: "第 {n} 步：{name}，{state}",
      workflowStateCurrent: "当前步骤",
      workflowStateComplete: "已完成",
      workflowStateUpcoming: "尚未开始",
      themeToLight: "切换到浅色外观",
      themeToDark: "切换到深色外观",
      githubBtn: "GitHub",
      githubAria: "在 GitHub 打开本项目，欢迎点 Star",
      stepImport: "01 · 导入",
      dropzoneAria: "拖入或点击选择要转换的文件",
      dropTitle: "把文件放到这里",
      dropHint: "拖放 ODT、FODT、RTF、HTML、Markdown 或文本文件，也可以一次选择多个文件。",
      chooseFiles: "选择文件",
      dropFormats: "ODT · FODT · RTF · HTML · MD · TXT · CSV · 单个最大 {maxFileSize}",
      localPromise: "文件不会离开此设备",
      stepOptions: "02 · 设置",
      optionsTitle: "输出方案",
      optionsSubtitle: "先选一个版式预设，再按需要微调。",
      presetsLabel: "快捷预设",
      presetStandard: "标准 Word",
      presetStandardDesc: "A4 · 25.4mm 留白 · 现代排版",
      presetAcademic: "学术报告",
      presetAcademicDesc: "A4 · 首行缩进 2 字符 · 衬线经典",
      presetCompact: "紧凑省纸",
      presetCompactDesc: "A4 · 12.7mm 窄边距 · 高效阅读",
      presetMerge: "合并为一个 Word",
      presetMergeDesc: "所有文件按顺序合并为单一文档",
      presetPlainText: "纯净纯文本",
      presetPlainTextDesc: "提纯正文与结构 · 导出 .txt",
      presetApplied: "已应用“{name}”预设",
      optTargetFormat: "目标格式",
      formatDocx: "Microsoft Word (.docx)",
      formatTxt: "纯文本 (.txt)",
      formatHtml: "网页文档 (.html)",
      formatMd: "Markdown (.md)",
      optDelivery: "生成方式",
      deliverySeparate: "分别输出（每个文件单独生成）",
      deliveryMerge: "合并输出（拼为一个完整文档）",
      optZipBundle: "多文件打包为 ZIP 下载",
      optPageSize: "页面尺寸",
      pageSizeHelp: "仅对 Word (.docx) 输出生效。",
      optMargin: "页边距",
      marginNormal: "标准（25.4 mm / 1 英寸）",
      marginNormalDesc: "经典公文与商务标准",
      marginNarrow: "紧凑（12.7 mm / 0.5 英寸）",
      marginNarrowDesc: "边距缩紧，容纳更多内容",
      marginWide: "宽边距（31.8 mm）",
      marginWideDesc: "大面积呼吸感，适合读物与手稿",
      optTypography: "字体风格",
      fontModern: "现代商务（Calibri / 微软雅黑）",
      fontClassic: "经典公文（Times New Roman / 宋体）",
      fontSans: "清新无衬线（Arial / 黑体）",
      optFirstLineIndent: "首行缩进 2 字符（中文段落惯例）",
      optFilename: "合并文件名",
      filenameHelpMerge: "合并文件将保存为 {name}.{ext}",
      filenameHelpSeparate: "分别输出时沿用各个原文件名并修改扩展名。",
      stepQueue: "03 · 生成",
      queueTitle: "文件队列",
      queueEmptyCount: "0 个",
      queueCount: "{n} 个",
      queueTotal: "共 {size}",
      queueHint: "拖动手柄调整顺序；键盘聚焦手柄后可用 ↑ / ↓ 排序。",
      btnReverse: "倒序",
      btnClear: "清空",
      btnConvert: "开始转换",
      btnWorking: "正在转换…",
      emptyTitle: "你的文件会出现在这里",
      emptyState: "导入后可以查看识别类型、调整顺序和预览。",
      actionEmpty: "先添加文件，再开始转换",
      actionReady: "将转换 {n} 个文件 · {size}",
      statusPreparing: "正在准备…",
      statusParsing: "正在解析 {name}…",
      statusGenerating: "正在生成 {ext} 文档…",
      statusBundling: "正在打包 ZIP…",
      statusDone: "转换完成！已下载 {n} 个文件。",
      trustLocalTitle: "100% 本地处理",
      trustLocalText: "所有转换在浏览器内完成，文件绝不上传",
      trustFidelityTitle: "高保真 ODT 转 Word",
      trustFidelityText: "完美保留段落、标题、表格、列表与内嵌图片",
      trustOfflineTitle: "完全离线可用",
      trustOfflineText: "内置纯 JS 依赖，无网络亦可运行",
      tipsTitle: "格式转换与排版支持说明",
      tip1: "<strong>100% 隐私安全</strong>：转换全部在浏览器内存中运算完成，没有后端服务器，文件<strong>绝不上云</strong>。",
      tip2: "<strong>ODT 格式高保真</strong>：深度解析 LibreOffice / OpenOffice ODT XML 结构，完整转换标题层级、粗体斜体、文字颜色、背景高亮、列表、表格与图片。",
      tip3: "<strong>FODT / RTF / HTML / MD</strong>：全面支持各种特殊文档结构，一键重排为规范的 Microsoft Word 格式。",
      tip4: "<strong>合并功能</strong>：开启“合并为一个 Word”后，可将多份零散材料依序组合，自动插入分页符与统一章节样式。",
      tip5: "<strong>纯文本导出</strong>：选择 .txt 可快速提纯文本，自动清理冗余格式，并保留清晰的表格与清单结构。",
      tip6: "<strong>完全免安装</strong>：随时在手机或电脑浏览器打开即用，支持离线与局域网静态部署。",
      footerNote: "开源 · 纯前端 · 零依赖云端",
      footerGithub: "GitHub 仓库",
      toastAdded: "已添加 {n} 个文件",
      toastRemoved: "已移除 {name}",
      toastCleared: "队列已清空",
      toastConverted: "转换完成，已下载！",
      toastError: "处理出错：{msg}",
      unsupportedFormat: "不支持的文件格式",
      unsupportedDoc: "旧版 .doc 或 .wps 请在 Office 中另存为 .docx 或 .odt 再转换",
      fileTooLarge: "文件超过 {limit} 大小限制",
      tooManyFiles: "已达到最大文件数限制 ({limit} 个)",
    },
    en: {
      pageTitle: "KenEasy Word Converter — Local ODT & Universal Document to Word / TXT",
      pageDesc: "Drop ODT, FODT, RTF, HTML, Markdown or text files, and convert them to high-fidelity Microsoft Word (.docx) or TXT locally in your browser. No install, no upload, zero data leakage.",
      skipToMain: "Skip to main content",
      eyebrow: "A CALMER WORD WORKFLOW",
      heroTitle: "From any document, to a clean Word file.",
      tagline: "Drop, reorder, customize styling, and convert with one click. All parsing and formatting happens quietly on your device.",
      badgeNoInstall: "No Install",
      badgeLocal: "100% Local",
      badgeFidelity: "High-Fidelity ODT",
      badgeMixed: "Multi-Format",
      utilityAria: "Utility navigation",
      featuresAria: "Product features",
      guaranteesAria: "Guarantees",
      presetsAria: "Formatting presets",
      workflowAria: "Workflow steps",
      workflowImport: "Import Files",
      workflowTune: "Formatting",
      workflowExport: "Export Word",
      workflowStepLabel: "Step {n}: {name}, {state}",
      workflowStateCurrent: "current step",
      workflowStateComplete: "completed",
      workflowStateUpcoming: "not started",
      themeToLight: "Switch to light theme",
      themeToDark: "Switch to dark theme",
      githubBtn: "GitHub",
      githubAria: "Open this project on GitHub (Star welcome!)",
      stepImport: "01 · Import",
      dropzoneAria: "Drop or click to select files to convert",
      dropTitle: "Drop your files here",
      dropHint: "Drag and drop ODT, FODT, RTF, HTML, Markdown or text files, or pick multiple files.",
      chooseFiles: "Choose Files",
      dropFormats: "ODT · FODT · RTF · HTML · MD · TXT · CSV · Max {maxFileSize} per file",
      localPromise: "Files never leave this device",
      stepOptions: "02 · Options",
      optionsTitle: "Output Recipe",
      optionsSubtitle: "Pick a quick preset first, then fine-tune if needed.",
      presetsLabel: "Quick Presets",
      presetStandard: "Standard Word",
      presetStandardDesc: "A4 · 25.4mm margin · Modern layout",
      presetAcademic: "Academic / Report",
      presetAcademicDesc: "A4 · 2-char indent · Classic serif",
      presetCompact: "Compact & Lean",
      presetCompactDesc: "A4 · 12.7mm narrow margin · High density",
      presetMerge: "Merge into One Word",
      presetMergeDesc: "Combine all files into a single Word document",
      presetPlainText: "Clean Plain Text",
      presetPlainTextDesc: "Purify body text · Export .txt",
      presetApplied: "Applied \"{name}\" preset",
      optTargetFormat: "Target Format",
      formatDocx: "Microsoft Word (.docx)",
      formatTxt: "Plain Text (.txt)",
      formatHtml: "HTML Document (.html)",
      formatMd: "Markdown (.md)",
      optDelivery: "Output Mode",
      deliverySeparate: "Separate files (convert each individually)",
      deliveryMerge: "Merge output (combine into one document)",
      optZipBundle: "Bundle multiple files as a ZIP archive",
      optPageSize: "Page Size",
      pageSizeHelp: "Applies to Word (.docx) export.",
      optMargin: "Page Margins",
      marginNormal: "Normal (25.4 mm / 1 inch)",
      marginNormalDesc: "Classic business standard",
      marginNarrow: "Narrow (12.7 mm / 0.5 inch)",
      marginNarrowDesc: "More text per page",
      marginWide: "Wide (31.8 mm)",
      marginWideDesc: "Generous whitespace for manuscripts",
      optTypography: "Typography",
      fontModern: "Modern (Calibri / Microsoft YaHei)",
      fontClassic: "Classic (Times New Roman / SimSun)",
      fontSans: "Clean Sans (Arial / SimHei)",
      optFirstLineIndent: "First-line paragraph indent (Chinese convention)",
      optFilename: "Output Filename",
      filenameHelpMerge: "Merged file will be saved as {name}.{ext}",
      filenameHelpSeparate: "Separate files keep their original names with updated extensions.",
      stepQueue: "03 · Convert",
      queueTitle: "Document Queue",
      queueEmptyCount: "0 files",
      queueCount: "{n} files",
      queueTotal: "total {size}",
      queueHint: "Drag handles to reorder; keyboard focus handles and use ↑ / ↓ to sort.",
      btnReverse: "Reverse",
      btnClear: "Clear",
      btnConvert: "Convert Now",
      btnWorking: "Converting…",
      emptyTitle: "Your documents will appear here",
      emptyState: "Once added, you can review detected types, reorder, and convert.",
      actionEmpty: "Add documents first, then convert",
      actionReady: "Ready to convert {n} file(s) · {size}",
      statusPreparing: "Preparing documents…",
      statusParsing: "Parsing {name}…",
      statusGenerating: "Generating {ext} document…",
      statusBundling: "Creating ZIP archive…",
      statusDone: "Conversion complete! Downloaded {n} file(s).",
      trustLocalTitle: "100% Local Processing",
      trustLocalText: "All conversion executes in-browser, files are never uploaded",
      trustFidelityTitle: "High-Fidelity ODT to Word",
      trustFidelityText: "Faithfully preserves headings, bold/italic, tables, lists, and images",
      trustOfflineTitle: "Fully Offline Capable",
      trustOfflineText: "Pure JS offline dependencies, no internet required",
      tipsTitle: "Format & Compatibility Guide",
      tip1: "<strong>100% Private</strong>: Computations happen entirely in browser memory. There is no backend, and files <strong>never leave your machine</strong>.",
      tip2: "<strong>High-Fidelity ODT</strong>: Deeply parses LibreOffice / OpenOffice XML structure, retaining heading hierarchies, colors, highlights, lists, tables, and embedded images.",
      tip3: "<strong>FODT / RTF / HTML / MD</strong>: Converts unusual and diverse document formats into clean, compliant Microsoft Word documents.",
      tip4: "<strong>Document Merging</strong>: Merge mode cleanly stitches multiple files together with page breaks and unified styling.",
      tip5: "<strong>Plain Text Extraction</strong>: Choose .txt to extract crisp, distraction-free text with formatted ASCII tables and bullet points.",
      tip6: "<strong>Zero Installation</strong>: Works instantly on any modern desktop or mobile browser.",
      footerNote: "Open Source · Pure Front-End · Zero Cloud Dependency",
      footerGithub: "GitHub Repository",
      toastAdded: "Added {n} file(s)",
      toastRemoved: "Removed {name}",
      toastCleared: "Queue cleared",
      toastConverted: "Converted and downloaded successfully!",
      toastError: "Error: {msg}",
      unsupportedFormat: "Unsupported document format",
      unsupportedDoc: "Legacy .doc or .wps should be saved as .docx or .odt first",
      fileTooLarge: "File exceeds {limit} limit",
      tooManyFiles: "Queue limit reached ({limit} files max)",
    },
  };

  var currentLang = "zh";

  function formatBytes(bytes) {
    if (!bytes || bytes <= 0) return "0 B";
    var k = 1024;
    var sizes = ["B", "KB", "MB", "GB"];
    var i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  }

  var i18n = {
    setLanguage: function (lang) {
      if (COPY[lang]) currentLang = lang;
      document.documentElement.lang = currentLang === "zh" ? "zh-CN" : "en";
      i18n.translateDom();
    },

    getLanguage: function () {
      return currentLang;
    },

    t: function (key, params) {
      var dict = COPY[currentLang] || COPY.zh;
      var text = dict[key] || (COPY.zh[key] || key);
      if (params && typeof params === "object") {
        for (var p in params) {
          text = text.replace(new RegExp("\\{" + p + "\\}", "g"), params[p]);
        }
      }
      return text;
    },

    formatBytes: formatBytes,

    translateDom: function (root) {
      var scope = root || document;

      var textNodes = scope.querySelectorAll("[data-i18n]");
      for (var i = 0; i < textNodes.length; i++) {
        var el = textNodes[i];
        var key = el.getAttribute("data-i18n");
        el.innerHTML = i18n.t(key);
      }

      var ariaNodes = scope.querySelectorAll("[data-i18n-aria-label]");
      for (var a = 0; a < ariaNodes.length; a++) {
        var aEl = ariaNodes[a];
        var aKey = aEl.getAttribute("data-i18n-aria-label");
        aEl.setAttribute("aria-label", i18n.t(aKey));
      }

      var titleNodes = scope.querySelectorAll("[data-i18n-title]");
      for (var t = 0; t < titleNodes.length; t++) {
        var tEl = titleNodes[t];
        var tKey = tEl.getAttribute("data-i18n-title");
        tEl.setAttribute("title", i18n.t(tKey));
      }

      var phNodes = scope.querySelectorAll("[data-i18n-placeholder]");
      for (var p = 0; p < phNodes.length; p++) {
        var pEl = phNodes[p];
        var pKey = pEl.getAttribute("data-i18n-placeholder");
        pEl.setAttribute("placeholder", i18n.t(pKey));
      }

      document.title = i18n.t("pageTitle");
      var descMeta = document.querySelector('meta[name="description"]');
      if (descMeta) descMeta.setAttribute("content", i18n.t("pageDesc"));
    },
  };

  global.KenEasyWordI18n = Object.freeze(i18n);
})(typeof window !== "undefined" ? window : globalThis);
