/* KenEasy Word Converter — UI orchestration and application layer.
 * Strictly decoupled: domain rules in config.js and doc-engine.js, copy in i18n.js.
 */
(function (global) {
  "use strict";

  var config = global.KenEasyWordConfig;
  var i18n = global.KenEasyWordI18n;
  var engine = global.KenEasyWordEngine;

  if (!config || !i18n || !engine) {
    throw new Error("KenEasy Word Converter core modules failed to load.");
  }

  // Elements cache
  var els = {};
  var items = [];
  var uidCounter = 0;
  var sortable = null;
  var isBusy = false;
  var currentTheme = "light";
  var toastTimer = null;

  function $(id) {
    return document.getElementById(id);
  }

  function t(key, params) {
    return i18n.t(key, params);
  }

  function showToast(msg, isError) {
    var container = els.toastContainer || $("toastContainer");
    if (!container) return;
    var toast = document.createElement("div");
    toast.className = "toast" + (isError ? " toast-error" : "");
    toast.textContent = msg;
    container.appendChild(toast);

    requestAnimationFrame(function () {
      toast.classList.add("toast-show");
    });

    setTimeout(function () {
      toast.classList.remove("toast-show");
      setTimeout(function () {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, 3200);
  }

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 2000);
  }

  // -------------------------------------------------------------
  // Theme Management
  // -------------------------------------------------------------
  function initTheme() {
    var saved = null;
    try {
      saved = localStorage.getItem(config.storageKeys.appearance);
    } catch (e) {}

    if (saved === "light" || saved === "dark") {
      currentTheme = saved;
    } else {
      currentTheme = global.matchMedia && global.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    }
    applyTheme(currentTheme, false);
  }

  function applyTheme(theme, save) {
    currentTheme = theme;
    document.documentElement.setAttribute("data-theme", theme);
    var themeColorMeta = document.querySelector('meta[name="theme-color"]');
    if (themeColorMeta) {
      themeColorMeta.setAttribute("content", theme === "dark" ? "#0b0c0f" : "#f5f5f7");
    }
    if (save) {
      try {
        localStorage.setItem(config.storageKeys.appearance, theme);
      } catch (e) {}
    }
    if (els.themeToggle) {
      var labelKey = theme === "dark" ? "themeToLight" : "themeToDark";
      els.themeToggle.setAttribute("aria-label", t(labelKey));
      els.themeToggle.setAttribute("title", t(labelKey));
    }
  }

  function toggleTheme() {
    applyTheme(currentTheme === "dark" ? "light" : "dark", true);
  }

  // -------------------------------------------------------------
  // Language Management
  // -------------------------------------------------------------
  function initLanguage() {
    var saved = null;
    try {
      saved = localStorage.getItem(config.storageKeys.language);
    } catch (e) {}
    var lang = saved === "en" || saved === "zh" ? saved : "zh";
    setLanguage(lang, false);
  }

  function setLanguage(lang, save) {
    i18n.setLanguage(lang);
    if (save) {
      try {
        localStorage.setItem(config.storageKeys.language, lang);
      } catch (e) {}
    }
    if (els.langZh) els.langZh.setAttribute("aria-pressed", String(lang === "zh"));
    if (els.langEn) els.langEn.setAttribute("aria-pressed", String(lang === "en"));
    renderPresets();
    renderQueue();
    syncOptionsVisibility();
    updateWorkflowState();
  }

  // -------------------------------------------------------------
  // Presets
  // -------------------------------------------------------------
  function renderPresets() {
    if (!els.presetList) return;
    els.presetList.innerHTML = "";

    var activePresetId = getCurrentActivePresetId();

    for (var i = 0; i < config.presets.length; i++) {
      var p = config.presets[i];
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "preset-btn" + (p.id === activePresetId ? " active" : "");
      btn.setAttribute("data-preset-id", p.id);

      var nameText = t(p.nameKey);
      var descText = t(p.descKey);

      btn.innerHTML = "<strong>" + nameText + "</strong><span>" + descText + "</span>";

      btn.addEventListener("click", (function (preset) {
        return function () {
          applyPreset(preset);
        };
      })(p));

      els.presetList.appendChild(btn);
    }
  }

  function applyPreset(preset) {
    if (!preset) return;
    var s = preset.values || preset.settings;
    if (!s) return;

    if (s.targetFormat && els.optTargetFormat) els.optTargetFormat.value = s.targetFormat;
    if (s.delivery && els.optDelivery) els.optDelivery.value = s.delivery;
    if (typeof s.zipBundle === "boolean" && els.optZipBundle) els.optZipBundle.checked = s.zipBundle;
    if (s.pageSize && els.optPageSize) els.optPageSize.value = s.pageSize;
    if (s.margin && els.optMargin) els.optMargin.value = s.margin;
    if (s.typography && els.optTypography) els.optTypography.value = s.typography;
    if (typeof s.firstLineIndent === "boolean" && els.optFirstLineIndent) els.optFirstLineIndent.checked = s.firstLineIndent;

    renderPresets();
    syncOptionsVisibility();
    showToast(t("presetApplied", { name: t(preset.nameKey) }));
  }

  function getCurrentActivePresetId() {
    var curFormat = els.optTargetFormat ? els.optTargetFormat.value : "docx";
    var curDelivery = els.optDelivery ? els.optDelivery.value : "separate";
    var curMargin = els.optMargin ? els.optMargin.value : "normal";
    var curIndent = els.optFirstLineIndent ? els.optFirstLineIndent.checked : false;

    if (curFormat === "txt") return "plaintext";
    if (curDelivery === "merge") return "mergeDocx";
    if (curMargin === "narrow") return "compact";
    if (curIndent === true) return "academic";
    return "standard";
  }

  // -------------------------------------------------------------
  // Settings & Visibility
  // -------------------------------------------------------------
  function syncOptionsVisibility() {
    var format = els.optTargetFormat ? els.optTargetFormat.value : "docx";
    var delivery = els.optDelivery ? els.optDelivery.value : "separate";
    var isWord = format === "docx";

    var wordOnlyElements = document.querySelectorAll(".word-only-option");
    for (var i = 0; i < wordOnlyElements.length; i++) {
      wordOnlyElements[i].style.display = isWord ? "" : "none";
    }

    if (els.zipBundleGroup) {
      els.zipBundleGroup.style.display = delivery === "separate" ? "" : "none";
    }

    if (els.filenameHelp) {
      var ext = format;
      var name = (els.optFilename && els.optFilename.value.trim()) || "converted";
      if (delivery === "merge") {
        els.filenameHelp.textContent = t("filenameHelpMerge", { name: name, ext: ext });
      } else {
        els.filenameHelp.textContent = t("filenameHelpSeparate");
      }
    }
  }

  // -------------------------------------------------------------
  // Queue Management
  // -------------------------------------------------------------
  function addFiles(fileList) {
    if (!fileList || fileList.length === 0) return;
    var countAdded = 0;

    for (var i = 0; i < fileList.length; i++) {
      var file = fileList[i];

      // Limit check
      if (items.length >= config.limits.maxFiles) {
        showToast(t("tooManyFiles", { limit: config.limits.maxFiles }), true);
        break;
      }

      if (file.size > config.limits.maxFileBytes) {
        showToast(file.name + ": " + t("fileTooLarge", { limit: i18n.formatBytes(config.limits.maxFileBytes) }), true);
        continue;
      }

      var kind = engine.detectKind(file);
      if (kind === "unknown") {
        showToast(file.name + ": " + t("unsupportedFormat"), true);
        continue;
      }
      if (kind === "unsupported-doc") {
        showToast(file.name + ": " + t("unsupportedDoc"), true);
        continue;
      }

      uidCounter++;
      items.push({
        id: "item-" + uidCounter,
        file: file,
        name: file.name,
        size: file.size,
        kind: kind,
      });
      countAdded++;
    }

    if (countAdded > 0) {
      showToast(t("toastAdded", { n: countAdded }));
      renderQueue();
      updateWorkflowState();
    }
  }

  function removeItem(id) {
    var index = -1;
    for (var i = 0; i < items.length; i++) {
      if (items[i].id === id) {
        index = i;
        break;
      }
    }
    if (index !== -1) {
      var name = items[index].name;
      items.splice(index, 1);
      renderQueue();
      updateWorkflowState();
      showToast(t("toastRemoved", { name: name }));
    }
  }

  function clearQueue() {
    if (items.length === 0) return;
    items = [];
    renderQueue();
    updateWorkflowState();
    showToast(t("toastCleared"));
  }

  function reverseQueue() {
    if (items.length <= 1) return;
    items.reverse();
    renderQueue();
  }

  function getKindBadgeClass(kind) {
    if (kind === "odt" || kind === "fodt") return "kind-badge pink";
    if (kind === "docx") return "kind-badge blue";
    if (kind === "rtf" || kind === "html") return "kind-badge blue";
    if (kind === "md" || kind === "text") return "kind-badge emerald";
    return "kind-badge muted";
  }

  function renderQueue() {
    var list = els.queueList;
    if (!list) return;

    var totalBytes = 0;
    for (var b = 0; b < items.length; b++) totalBytes += items[b].size;

    if (els.queueCountLabel) {
      els.queueCountLabel.textContent = items.length === 0 ? t("queueEmptyCount") : t("queueCount", { n: items.length });
    }
    if (els.queueTotalSizeLabel) {
      els.queueTotalSizeLabel.textContent = t("queueTotal", { size: i18n.formatBytes(totalBytes) });
    }

    if (items.length === 0) {
      list.innerHTML =
        '<div class="empty-queue">' +
        '<svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="12" y1="18" x2="12" y2="12"></line><line x1="9" y1="15" x2="15" y2="15"></line></svg>' +
        "<strong>" +
        t("emptyTitle") +
        "</strong>" +
        "<p>" +
        t("emptyState") +
        "</p>" +
        "</div>";
      if (els.convertButton) els.convertButton.disabled = true;
      if (els.actionStatus) els.actionStatus.textContent = t("actionEmpty");
      if (els.actionHint) els.actionHint.textContent = t("emptyState");
      return;
    }

    if (els.convertButton) els.convertButton.disabled = isBusy;
    if (els.actionStatus) els.actionStatus.textContent = t("actionReady", { n: items.length, size: i18n.formatBytes(totalBytes) });
    if (els.actionHint) els.actionHint.textContent = t("workflowStateUpcoming");

    list.innerHTML = "";

    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var div = document.createElement("div");
      div.className = "queue-item";
      div.setAttribute("data-id", it.id);

      var badgeClass = getKindBadgeClass(it.kind);
      var badgeText = it.kind.toUpperCase();

      div.innerHTML =
        '<span class="drag-handle" tabindex="0" title="' +
        t("queueHint") +
        '">' +
        '<svg viewBox="0 0 20 20"><path d="M7 4h2v2H7V4zm4 0h2v2h-2V4zM7 9h2v2H7V9zm4 0h2v2h-2V9zm-4 5h2v2H7v-2zm4 0h2v2h-2v-2z"></path></svg>' +
        "</span>" +
        '<span class="' +
        badgeClass +
        '">' +
        badgeText +
        "</span>" +
        '<div class="item-info">' +
        '<div class="item-name" title="' +
        it.name +
        '">' +
        it.name +
        "</div>" +
        '<div class="item-meta">' +
        i18n.formatBytes(it.size) +
        "</div>" +
        "</div>" +
        '<div class="item-actions">' +
        '<button type="button" class="item-remove" data-action="remove" title="' +
        t("toastRemoved", { name: "" }).replace(/\s+/, "") +
        '">' +
        '<svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>' +
        "</button>" +
        "</div>";

      // Attach remove event
      (function (itemId) {
        var rmBtn = div.querySelector('[data-action="remove"]');
        if (rmBtn) {
          rmBtn.addEventListener("click", function (e) {
            e.stopPropagation();
            removeItem(itemId);
          });
        }
      })(it.id);

      // Keyboard reordering on handle
      (function (itemIdx) {
        var handle = div.querySelector(".drag-handle");
        if (handle) {
          handle.addEventListener("keydown", function (e) {
            if (e.key === "ArrowUp" && itemIdx > 0) {
              e.preventDefault();
              var temp = items[itemIdx];
              items[itemIdx] = items[itemIdx - 1];
              items[itemIdx - 1] = temp;
              renderQueue();
              focusHandle(itemIdx - 1);
            } else if (e.key === "ArrowDown" && itemIdx < items.length - 1) {
              e.preventDefault();
              var temp2 = items[itemIdx];
              items[itemIdx] = items[itemIdx + 1];
              items[itemIdx + 1] = temp2;
              renderQueue();
              focusHandle(itemIdx + 1);
            }
          });
        }
      })(i);

      list.appendChild(div);
    }

    initSortable();
  }

  function focusHandle(idx) {
    setTimeout(function () {
      var all = els.queueList.querySelectorAll(".drag-handle");
      if (all && all[idx]) all[idx].focus();
    }, 50);
  }

  function initSortable() {
    if (!global.Sortable || !els.queueList) return;
    if (sortable) sortable.destroy();

    sortable = new global.Sortable(els.queueList, {
      handle: ".drag-handle",
      ghostClass: "sortable-ghost",
      animation: 180,
      onEnd: function () {
        var domIds = Array.prototype.map.call(els.queueList.children, function (el) {
          return el.getAttribute("data-id");
        });
        var reordered = [];
        for (var d = 0; d < domIds.length; d++) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].id === domIds[d]) {
              reordered.push(items[i]);
              break;
            }
          }
        }
        items = reordered;
      },
    });
  }

  function updateWorkflowState() {
    var step1 = document.querySelector('[data-workflow-step="1"]');
    var step2 = document.querySelector('[data-workflow-step="2"]');
    var step3 = document.querySelector('[data-workflow-step="3"]');

    if (items.length === 0) {
      if (step1) step1.setAttribute("aria-current", "step");
      if (step2) step2.removeAttribute("aria-current");
      if (step3) step3.removeAttribute("aria-current");
    } else if (!isBusy) {
      if (step1) step1.removeAttribute("aria-current");
      if (step2) step2.setAttribute("aria-current", "step");
      if (step3) step3.removeAttribute("aria-current");
    } else {
      if (step1) step1.removeAttribute("aria-current");
      if (step2) step2.removeAttribute("aria-current");
      if (step3) step3.setAttribute("aria-current", "step");
    }
  }

  function setProgress(percent, statusText) {
    if (els.progressWrap && els.progressFill) {
      if (percent >= 0) {
        els.progressWrap.style.display = "block";
        els.progressFill.style.width = percent + "%";
      } else {
        els.progressWrap.style.display = "none";
      }
    }
    if (els.actionStatus && statusText) {
      els.actionStatus.textContent = statusText;
    }
  }

  // -------------------------------------------------------------
  // Conversion Pipeline
  // -------------------------------------------------------------
  async function startConversion() {
    if (items.length === 0 || isBusy) return;

    var format = els.optTargetFormat ? els.optTargetFormat.value : "docx";
    var delivery = els.optDelivery ? els.optDelivery.value : "separate";
    var zipBundle = els.optZipBundle ? els.optZipBundle.checked : true;
    var pageSize = els.optPageSize ? els.optPageSize.value : "a4";
    var margin = els.optMargin ? els.optMargin.value : "normal";
    var typography = els.optTypography ? els.optTypography.value : "modern";
    var firstLineIndent = els.optFirstLineIndent ? els.optFirstLineIndent.checked : false;
    var customName = (els.optFilename && els.optFilename.value.trim()) || "converted";

    var options = {
      targetFormat: format,
      pageSize: pageSize,
      margin: margin,
      typography: typography,
      firstLineIndent: firstLineIndent,
      filename: customName,
    };

    isBusy = true;
    if (els.convertButton) {
      els.convertButton.disabled = true;
      els.convertButton.innerHTML = "<span>" + t("btnWorking") + "</span>";
    }
    updateWorkflowState();
    setProgress(5, t("statusPreparing"));

    try {
      if (delivery === "merge") {
        // Merge mode: parse all, combine ASTs, export 1 document
        var asts = [];
        for (var i = 0; i < items.length; i++) {
          var it = items[i];
          var pct = Math.round(10 + (i / items.length) * 50);
          setProgress(pct, t("statusParsing", { name: it.name }));
          var ast = await engine.parseFile(it.file);
          asts.push(ast);
        }

        setProgress(70, t("statusGenerating", { ext: format.toUpperCase() }));
        var mergedAst = engine.mergeAsts(asts, customName);
        var result = await engine.convertAst(mergedAst, options);

        setProgress(100, t("statusDone", { n: 1 }));
        downloadBlob(result.blob, result.filename);
        showToast(t("toastConverted"));
      } else {
        // Separate mode: convert each
        var convertedList = [];
        for (var j = 0; j < items.length; j++) {
          var item = items[j];
          var stepPct = Math.round(10 + (j / items.length) * 60);
          setProgress(stepPct, t("statusParsing", { name: item.name }));

          var singleAst = await engine.parseFile(item.file);
          var fileOpt = Object.assign({}, options, {
            filename: item.name.replace(/\.[^.]+$/, ""),
          });
          var singleResult = await engine.convertAst(singleAst, fileOpt);
          convertedList.push({
            name: singleResult.filename,
            blob: singleResult.blob,
          });
        }

        if (zipBundle && convertedList.length > 1) {
          setProgress(85, t("statusBundling"));
          var zipBlob = await engine.bundleZip(convertedList, customName + ".zip");
          setProgress(100, t("statusDone", { n: convertedList.length }));
          downloadBlob(zipBlob, customName + ".zip");
          showToast(t("toastConverted"));
        } else {
          // Download individually
          for (var k = 0; k < convertedList.length; k++) {
            (function (resItem, delay) {
              setTimeout(function () {
                downloadBlob(resItem.blob, resItem.name);
              }, delay);
            })(convertedList[k], k * 280);
          }
          setProgress(100, t("statusDone", { n: convertedList.length }));
          showToast(t("toastConverted"));
        }
      }
    } catch (err) {
      console.error("Conversion error:", err);
      showToast(t("toastError", { msg: err.message || "Unknown error" }), true);
      setProgress(-1, t("toastError", { msg: err.message || "" }));
    } finally {
      isBusy = false;
      if (els.convertButton) {
        els.convertButton.disabled = false;
        els.convertButton.innerHTML = "<span>" + t("btnConvert") + "</span>";
      }
      setTimeout(function () {
        setProgress(-1);
        updateWorkflowState();
        renderQueue();
      }, 2500);
    }
  }

  // -------------------------------------------------------------
  // Initialization & Event Binding
  // -------------------------------------------------------------
  function cacheElements() {
    els.themeToggle = $("themeToggle");
    els.langZh = $("langZh");
    els.langEn = $("langEn");
    els.dropzone = $("dropzone");
    els.fileInput = $("fileInput");
    els.chooseFilesButton = $("chooseFilesButton");
    els.presetList = $("presetList");
    els.optTargetFormat = $("optTargetFormat");
    els.optDelivery = $("optDelivery");
    els.optZipBundle = $("optZipBundle");
    els.zipBundleGroup = $("zipBundleGroup");
    els.optPageSize = $("optPageSize");
    els.optMargin = $("optMargin");
    els.optTypography = $("optTypography");
    els.optFirstLineIndent = $("optFirstLineIndent");
    els.optFilename = $("optFilename");
    els.filenameHelp = $("filenameHelp");
    els.queueList = $("queueList");
    els.queueCountLabel = $("queueCountLabel");
    els.queueTotalSizeLabel = $("queueTotalSizeLabel");
    els.clearQueueButton = $("clearQueueButton");
    els.reverseQueueButton = $("reverseQueueButton");
    els.convertButton = $("convertButton");
    els.actionStatus = $("actionStatus");
    els.actionHint = $("actionHint");
    els.progressWrap = $("progressWrap");
    els.progressFill = $("progressFill");
    els.toastContainer = $("toastContainer");
  }

  function bindEvents() {
    // Theme toggle
    if (els.themeToggle) {
      els.themeToggle.addEventListener("click", toggleTheme);
    }

    // Language buttons
    if (els.langZh) {
      els.langZh.addEventListener("click", function () {
        setLanguage("zh", true);
      });
    }
    if (els.langEn) {
      els.langEn.addEventListener("click", function () {
        setLanguage("en", true);
      });
    }

    // Dropzone & File Input
    if (els.dropzone && els.fileInput) {
      var dz = els.dropzone;
      var fi = els.fileInput;

      if (els.chooseFilesButton) {
        els.chooseFilesButton.addEventListener("click", function (e) {
          e.stopPropagation();
          fi.click();
        });
      }

      dz.addEventListener("click", function () {
        fi.click();
      });

      dz.addEventListener("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          fi.click();
        }
      });

      fi.addEventListener("change", function () {
        if (fi.files && fi.files.length > 0) {
          addFiles(fi.files);
          fi.value = "";
        }
      });

      // Drag and drop events
      ["dragenter", "dragover"].forEach(function (eventName) {
        dz.addEventListener(eventName, function (e) {
          e.preventDefault();
          e.stopPropagation();
          dz.classList.add("drag-active");
        });
      });

      ["dragleave", "drop"].forEach(function (eventName) {
        dz.addEventListener(eventName, function (e) {
          e.preventDefault();
          e.stopPropagation();
          dz.classList.remove("drag-active");
        });
      });

      dz.addEventListener("drop", function (e) {
        var dt = e.dataTransfer;
        if (dt && dt.files && dt.files.length > 0) {
          addFiles(dt.files);
        }
      });

      // Global window drag prevention
      window.addEventListener("dragover", function (e) {
        e.preventDefault();
      });
      window.addEventListener("drop", function (e) {
        e.preventDefault();
      });
    }

    // Settings changes
    if (els.optTargetFormat) {
      els.optTargetFormat.addEventListener("change", function () {
        syncOptionsVisibility();
        renderPresets();
      });
    }
    if (els.optDelivery) {
      els.optDelivery.addEventListener("change", function () {
        syncOptionsVisibility();
        renderPresets();
      });
    }
    if (els.optFilename) {
      els.optFilename.addEventListener("input", syncOptionsVisibility);
    }

    // Queue buttons
    if (els.clearQueueButton) els.clearQueueButton.addEventListener("click", clearQueue);
    if (els.reverseQueueButton) els.reverseQueueButton.addEventListener("click", reverseQueue);

    // Convert button
    if (els.convertButton) els.convertButton.addEventListener("click", startConversion);
  }

  function init() {
    cacheElements();
    initTheme();
    initLanguage();
    bindEvents();
    renderPresets();
    renderQueue();
    syncOptionsVisibility();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})(typeof window !== "undefined" ? window : globalThis);
