/* KenEasy Word Converter — unified document parsing and conversion engine.
 * Pure client-side, zero backend, offline-capable via JSZip.
 */
(function (global) {
  "use strict";

  var XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  // Helper: escape XML entities
  function escapeXml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  // Helper: parse measurement (cm, mm, in, pt, px) to dxa (1/20 of a pt; 1 in = 1440 dxa; 1 pt = 20 dxa; 1 cm = 567 dxa)
  function parseToDxa(val, defaultDxa) {
    if (!val) return defaultDxa || 0;
    var str = String(val).trim().toLowerCase();
    var num = parseFloat(str);
    if (isNaN(num)) return defaultDxa || 0;
    if (str.indexOf("cm") !== -1) return Math.round(num * 567);
    if (str.indexOf("mm") !== -1) return Math.round(num * 56.7);
    if (str.indexOf("in") !== -1) return Math.round(num * 1440);
    if (str.indexOf("pt") !== -1) return Math.round(num * 20);
    if (str.indexOf("px") !== -1) return Math.round(num * 15);
    return Math.round(num * 20);
  }

  // Helper: parse measurement to EMU (English Metric Units; 1 in = 914400 EMU; 1 pt = 12700 EMU; 1 cm = 360000 EMU)
  function parseToEmu(val, defaultEmu) {
    if (!val) return defaultEmu || 3600000;
    var str = String(val).trim().toLowerCase();
    var num = parseFloat(str);
    if (isNaN(num)) return defaultEmu || 3600000;
    if (str.indexOf("cm") !== -1) return Math.round(num * 360000);
    if (str.indexOf("mm") !== -1) return Math.round(num * 36000);
    if (str.indexOf("in") !== -1) return Math.round(num * 914400);
    if (str.indexOf("pt") !== -1) return Math.round(num * 12700);
    if (str.indexOf("px") !== -1) return Math.round(num * 9525);
    return Math.round(num * 12700);
  }

  // Helper: normalize hex color to 6-char RRGGBB without #
  function normalizeHexColor(colorStr) {
    if (!colorStr) return null;
    var s = String(colorStr).trim().replace(/^#/, "");
    if (s.length === 3) {
      return s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
    }
    if (s.length === 6) return s.toUpperCase();
    return null;
  }

  // -------------------------------------------------------------
  // Base AST Models
  // -------------------------------------------------------------
  function createDocumentAst(title) {
    return {
      title: title || "Document",
      elements: [], // Headings, paragraphs, tables, images, breaks
    };
  }

  // -------------------------------------------------------------
  // ODT (OpenDocument Text) & FODT (Flat XML) Parser
  // -------------------------------------------------------------
  var OdtParser = {
    parseOdt: async function (arrayBuffer, filename) {
      if (!global.JSZip) throw new Error("JSZip is required to parse ODT files.");
      var zip = await global.JSZip.loadAsync(arrayBuffer);
      var contentFile = zip.file("content.xml");
      if (!contentFile) throw new Error("Invalid ODT file: content.xml not found.");
      var contentXml = await contentFile.async("string");

      var stylesXml = "";
      var stylesFile = zip.file("styles.xml");
      if (stylesFile) {
        stylesXml = await stylesFile.async("string");
      }

      // Collect images in Pictures/
      var mediaMap = {};
      var pictureFiles = zip.file(/^Pictures\/.+/i);
      for (var i = 0; i < pictureFiles.length; i++) {
        var pf = pictureFiles[i];
        var u8 = await pf.async("uint8array");
        var ext = pf.name.split(".").pop().toLowerCase();
        var mime = "image/png";
        if (ext === "jpg" || ext === "jpeg") mime = "image/jpeg";
        else if (ext === "gif") mime = "image/gif";
        else if (ext === "webp") mime = "image/webp";
        else if (ext === "svg") mime = "image/svg+xml";
        mediaMap[pf.name] = { data: u8, mimeType: mime, ext: ext };
      }

      return OdtParser.parseXml(contentXml, stylesXml, mediaMap, filename);
    },

    parseFodt: function (xmlText, filename) {
      // Flat XML: parse without zip
      return OdtParser.parseXml(xmlText, "", {}, filename);
    },

    parseXml: function (contentXml, stylesXml, mediaMap, filename) {
      var parser = new DOMParser();
      var doc = parser.parseFromString(contentXml, "application/xml");
      var parseError = doc.querySelector("parsererror");
      if (parseError) {
        // Inject standard ODT namespaces into root element as robust fallback
        var commonNamespaces = ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ';
        var fixedXml = contentXml.replace(/(<[a-zA-Z0-9_:-]+)(\s|>)/, "$1" + commonNamespaces + "$2");
        doc = parser.parseFromString(fixedXml, "application/xml");
        parseError = doc.querySelector("parsererror");
        if (parseError) {
          throw new Error("XML parsing error in ODT: " + parseError.textContent.slice(0, 100));
        }
      }

      // Parse styles (both automatic and global)
      var styles = OdtParser.extractStyles(doc);
      if (stylesXml) {
        try {
          var sDoc = parser.parseFromString(stylesXml, "application/xml");
          var globalStyles = OdtParser.extractStyles(sDoc);
          for (var k in globalStyles) {
            if (!styles[k]) styles[k] = globalStyles[k];
          }
        } catch (e) {
          console.warn("Could not parse styles.xml:", e);
        }
      }

      var ast = createDocumentAst(filename ? filename.replace(/\.[^.]+$/, "") : "ODT Document");

      // Find <office:text> or root
      var body = doc.getElementsByTagNameNS("*", "body")[0] || doc.querySelector("body");
      var officeText = body ? (body.getElementsByTagNameNS("*", "text")[0] || body) : doc.documentElement;

      OdtParser.walkChildren(officeText, ast, styles, mediaMap);
      return ast;
    },

    extractStyles: function (doc) {
      var styles = {};
      var styleNodes = doc.querySelectorAll("style, [name]");
      for (var i = 0; i < styleNodes.length; i++) {
        var node = styleNodes[i];
        var name = node.getAttribute("style:name") || node.getAttribute("name");
        var family = node.getAttribute("style:family") || node.getAttribute("family");
        if (!name) continue;

        var sObj = { family: family, text: {}, para: {}, table: {}, cell: {} };

        // Text properties
        var tp = node.querySelector("text-properties, [style\\:text-properties]");
        if (tp) {
          if (tp.getAttribute("fo:font-weight") === "bold") sObj.text.bold = true;
          if (tp.getAttribute("fo:font-style") === "italic") sObj.text.italic = true;
          if (tp.getAttribute("style:text-underline-style") === "solid" || tp.getAttribute("text-underline-style")) sObj.text.underline = true;
          if (tp.getAttribute("style:text-line-through-style") === "solid") sObj.text.strike = true;
          var color = tp.getAttribute("fo:color");
          if (color) sObj.text.color = normalizeHexColor(color);
          var bg = tp.getAttribute("fo:background-color");
          if (bg && bg !== "transparent") sObj.text.bgColor = normalizeHexColor(bg);
          var sz = tp.getAttribute("fo:font-size");
          if (sz) {
            var pt = parseFloat(sz);
            if (!isNaN(pt)) sObj.text.fontSizePt = pt;
          }
          var fontName = tp.getAttribute("style:font-name") || tp.getAttribute("fo:font-family");
          if (fontName) sObj.text.fontFamily = fontName;
        }

        // Paragraph properties
        var pp = node.querySelector("paragraph-properties, [style\\:paragraph-properties]");
        if (pp) {
          var align = pp.getAttribute("fo:text-align");
          if (align) {
            if (align === "start") align = "left";
            if (align === "end") align = "right";
            sObj.para.align = align;
          }
          var mb = pp.getAttribute("fo:margin-bottom");
          if (mb) sObj.para.spacingAfter = parseToDxa(mb, 120);
          var mt = pp.getAttribute("fo:margin-top");
          if (mt) sObj.para.spacingBefore = parseToDxa(mt, 0);
          var ti = pp.getAttribute("fo:text-indent");
          if (ti) sObj.para.textIndent = parseToDxa(ti, 0);
        }

        // Table / Cell properties
        var tcp = node.querySelector("table-cell-properties, [style\\:table-cell-properties]");
        if (tcp) {
          var cbg = tcp.getAttribute("fo:background-color");
          if (cbg && cbg !== "transparent") sObj.cell.bgColor = normalizeHexColor(cbg);
        }

        styles[name] = sObj;
      }
      return styles;
    },

    walkChildren: function (parentElement, ast, styles, mediaMap, listLevel) {
      var children = parentElement.children;
      for (var i = 0; i < children.length; i++) {
        var el = children[i];
        var tag = el.localName || el.nodeName.split(":").pop();

        if (tag === "h") {
          // Heading
          var level = parseInt(el.getAttribute("text:outline-level") || el.getAttribute("outline-level") || "1", 10);
          if (isNaN(level) || level < 1) level = 1;
          if (level > 6) level = 6;
          var hStyleName = el.getAttribute("text:style-name") || el.getAttribute("style-name");
          var hStyle = styles[hStyleName] || {};
          var hRuns = OdtParser.extractRuns(el, styles, hStyle.text, mediaMap);
          ast.elements.push({
            type: "heading",
            level: level,
            runs: hRuns,
            align: hStyle.para ? hStyle.para.align : "left",
          });
        } else if (tag === "p") {
          // Paragraph
          var pStyleName = el.getAttribute("text:style-name") || el.getAttribute("style-name");
          var pStyle = styles[pStyleName] || {};
          var pRuns = OdtParser.extractRuns(el, styles, pStyle.text, mediaMap);

          // Check if paragraph contains an embedded image frame
          var images = OdtParser.extractImages(el, mediaMap);
          for (var imgIdx = 0; imgIdx < images.length; imgIdx++) {
            ast.elements.push(images[imgIdx]);
          }

          if (pRuns.length > 0 || images.length === 0) {
            ast.elements.push({
              type: "paragraph",
              runs: pRuns,
              align: pStyle.para ? pStyle.para.align : "left",
              spacingBefore: pStyle.para ? pStyle.para.spacingBefore : 0,
              spacingAfter: pStyle.para ? pStyle.para.spacingAfter : 140,
              textIndent: pStyle.para ? pStyle.para.textIndent : 0,
            });
          }
        } else if (tag === "list") {
          // List
          var currentLevel = typeof listLevel === "number" ? listLevel + 1 : 0;
          var items = el.children;
          for (var j = 0; j < items.length; j++) {
            var item = items[j];
            var itemTag = item.localName || item.nodeName.split(":").pop();
            if (itemTag === "list-item") {
              // Extract paragraph or sub-list
              for (var k = 0; k < item.children.length; k++) {
                var child = item.children[k];
                var childTag = child.localName || child.nodeName.split(":").pop();
                if (childTag === "p" || childTag === "h") {
                  var itemStyle = styles[child.getAttribute("text:style-name")] || {};
                  var itemRuns = OdtParser.extractRuns(child, styles, itemStyle.text, mediaMap);
                  ast.elements.push({
                    type: "list-item",
                    level: currentLevel,
                    ordered: false,
                    runs: itemRuns,
                  });
                } else if (childTag === "list") {
                  OdtParser.walkChildren(item, ast, styles, mediaMap, currentLevel);
                }
              }
            }
          }
        } else if (tag === "table") {
          // Table
          var tableAst = OdtParser.extractTable(el, styles, mediaMap);
          if (tableAst) ast.elements.push(tableAst);
        } else if (tag === "section") {
          OdtParser.walkChildren(el, ast, styles, mediaMap, listLevel);
        }
      }
    },

    extractRuns: function (element, styles, parentTextStyle, mediaMap) {
      var runs = [];
      var nodes = element.childNodes;
      for (var i = 0; i < nodes.length; i++) {
        var node = nodes[i];
        if (node.nodeType === 3) {
          // Text node
          var text = node.nodeValue;
          if (text) {
            runs.push(Object.assign({ text: text }, parentTextStyle || {}));
          }
        } else if (node.nodeType === 1) {
          var tag = node.localName || node.nodeName.split(":").pop();
          if (tag === "s") {
            // Space count
            var count = parseInt(node.getAttribute("text:c") || node.getAttribute("c") || "1", 10);
            var spaces = "";
            for (var sc = 0; sc < count; sc++) spaces += " ";
            runs.push(Object.assign({ text: spaces }, parentTextStyle || {}));
          } else if (tag === "tab") {
            runs.push(Object.assign({ text: "\t" }, parentTextStyle || {}));
          } else if (tag === "line-break") {
            runs.push(Object.assign({ text: "\n", breakLine: true }, parentTextStyle || {}));
          } else if (tag === "span") {
            var sName = node.getAttribute("text:style-name") || node.getAttribute("style-name");
            var spanStyle = styles[sName] ? styles[sName].text : {};
            var merged = Object.assign({}, parentTextStyle || {}, spanStyle);
            var subRuns = OdtParser.extractRuns(node, styles, merged, mediaMap);
            runs.push.apply(runs, subRuns);
          } else if (tag === "a") {
            var href = node.getAttribute("xlink:href") || node.getAttribute("href") || "";
            var linkStyle = Object.assign({}, parentTextStyle || {}, { link: href, underline: true, color: "0066CC" });
            var linkRuns = OdtParser.extractRuns(node, styles, linkStyle, mediaMap);
            runs.push.apply(runs, linkRuns);
          }
        }
      }
      return runs;
    },

    extractImages: function (element, mediaMap) {
      var images = [];
      var frames = element.querySelectorAll("frame, [draw\\:frame]");
      for (var i = 0; i < frames.length; i++) {
        var frame = frames[i];
        var widthVal = frame.getAttribute("svg:width") || frame.getAttribute("width") || "12cm";
        var heightVal = frame.getAttribute("svg:height") || frame.getAttribute("height") || "8cm";
        var widthEmu = parseToEmu(widthVal, 4320000);
        var heightEmu = parseToEmu(heightVal, 2880000);

        var imgNode = frame.querySelector("image, [draw\\:image]");
        if (imgNode) {
          var href = imgNode.getAttribute("xlink:href") || imgNode.getAttribute("href");
          if (href) {
            // Clean href (may start with ./)
            var cleanHref = href.replace(/^\.\//, "");
            var media = mediaMap[cleanHref];
            if (media) {
              images.push({
                type: "image",
                data: media.data,
                mimeType: media.mimeType,
                ext: media.ext,
                widthEmu: widthEmu,
                heightEmu: heightEmu,
                name: cleanHref.split("/").pop(),
              });
            }
          }
        }
      }
      return images;
    },

    extractTable: function (tableEl, styles, mediaMap) {
      var rows = [];
      var rowNodes = tableEl.querySelectorAll("table-row, [table\\:table-row]");
      for (var r = 0; r < rowNodes.length; r++) {
        var rNode = rowNodes[r];
        var cells = [];
        var cellNodes = rNode.children;
        for (var c = 0; c < cellNodes.length; c++) {
          var cNode = cellNodes[c];
          var cTag = cNode.localName || cNode.nodeName.split(":").pop();
          if (cTag === "table-cell" || cTag === "covered-table-cell") {
            var colSpan = parseInt(cNode.getAttribute("table:number-columns-spanned") || "1", 10);
            var rowSpan = parseInt(cNode.getAttribute("table:number-rows-spanned") || "1", 10);
            var cellStyle = styles[cNode.getAttribute("table:style-name")] || {};

            var cellDoc = createDocumentAst();
            OdtParser.walkChildren(cNode, cellDoc, styles, mediaMap);

            cells.push({
              colSpan: isNaN(colSpan) ? 1 : colSpan,
              rowSpan: isNaN(rowSpan) ? 1 : rowSpan,
              bgColor: cellStyle.cell ? cellStyle.cell.bgColor : null,
              elements: cellDoc.elements,
            });
          }
        }
        if (cells.length > 0) {
          rows.push({ cells: cells });
        }
      }

      if (rows.length === 0) return null;
      return {
        type: "table",
        rows: rows,
      };
    },
  };

  // -------------------------------------------------------------
  // RTF (Rich Text Format) Parser
  // -------------------------------------------------------------
  var RtfParser = {
    parse: function (rtfText, filename) {
      var ast = createDocumentAst(filename ? filename.replace(/\.[^.]+$/, "") : "RTF Document");
      var text = String(rtfText);

      var stateStack = [];
      var curState = {
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        fontSizePt: 12,
        align: "left",
        color: null,
      };

      var colorTable = [];
      var inColorTbl = false;

      var currentParagraph = {
        type: "paragraph",
        runs: [],
        align: "left",
      };

      var i = 0;
      var len = text.length;

      while (i < len) {
        var char = text[i];

        if (char === "{") {
          stateStack.push(Object.assign({}, curState));
          i++;
        } else if (char === "}") {
          if (stateStack.length > 0) {
            curState = stateStack.pop();
          }
          if (inColorTbl) inColorTbl = false;
          i++;
        } else if (char === "\\") {
          // Control word or symbol
          i++;
          if (i >= len) break;
          var nextChar = text[i];

          if (nextChar === "{" || nextChar === "}" || nextChar === "\\") {
            currentParagraph.runs.push(Object.assign({ text: nextChar }, curState));
            i++;
          } else if (nextChar === "'") {
            // Hex character \'xx
            i++;
            var hex = text.substr(i, 2);
            i += 2;
            var code = parseInt(hex, 16);
            if (!isNaN(code)) {
              currentParagraph.runs.push(Object.assign({ text: String.fromCharCode(code) }, curState));
            }
          } else {
            // Read control word
            var word = "";
            while (i < len && /[a-zA-Z]/.test(text[i])) {
              word += text[i];
              i++;
            }
            // Read optional signed integer parameter
            var param = "";
            if (i < len && (text[i] === "-" || /[0-9]/.test(text[i]))) {
              while (i < len && (text[i] === "-" || /[0-9]/.test(text[i]))) {
                param += text[i];
                i++;
              }
            }
            // If followed by space, swallow space
            if (i < len && text[i] === " ") {
              i++;
            }

            var num = param !== "" ? parseInt(param, 10) : null;

            if (word === "par") {
              // Flush paragraph
              ast.elements.push(currentParagraph);
              currentParagraph = {
                type: "paragraph",
                runs: [],
                align: curState.align || "left",
              };
            } else if (word === "b") {
              curState.bold = num === 0 ? false : true;
            } else if (word === "i") {
              curState.italic = num === 0 ? false : true;
            } else if (word === "ul") {
              curState.underline = num === 0 ? false : true;
            } else if (word === "ulnone") {
              curState.underline = false;
            } else if (word === "strike") {
              curState.strike = num === 0 ? false : true;
            } else if (word === "fs") {
              if (num !== null) curState.fontSizePt = num / 2;
            } else if (word === "qc") {
              curState.align = "center";
              currentParagraph.align = "center";
            } else if (word === "qr") {
              curState.align = "right";
              currentParagraph.align = "right";
            } else if (word === "ql") {
              curState.align = "left";
              currentParagraph.align = "left";
            } else if (word === "qj") {
              curState.align = "justify";
              currentParagraph.align = "justify";
            } else if (word === "tab") {
              currentParagraph.runs.push(Object.assign({ text: "\t" }, curState));
            } else if (word === "u") {
              // Unicode codepoint \uN?
              if (num !== null) {
                var uChar = String.fromCharCode(num < 0 ? num + 65536 : num);
                currentParagraph.runs.push(Object.assign({ text: uChar }, curState));
                // Often followed by replacement char like ?, skip if next is non-backslash
                if (i < len && text[i] === "?") i++;
              }
            } else if (word === "colortbl") {
              inColorTbl = true;
            } else if (word === "cf" && num !== null) {
              if (colorTable[num]) curState.color = colorTable[num];
            } else if (word === "red" || word === "green" || word === "blue") {
              // Color table definition
            }
          }
        } else if (char === "\r" || char === "\n") {
          // RTF ignores raw newlines unless preceded by control word
          i++;
        } else {
          // Regular text run
          var runText = "";
          while (i < len && text[i] !== "\\" && text[i] !== "{" && text[i] !== "}" && text[i] !== "\r" && text[i] !== "\n") {
            runText += text[i];
            i++;
          }
          if (runText) {
            currentParagraph.runs.push(Object.assign({ text: runText }, curState));
          }
        }
      }

      if (currentParagraph.runs.length > 0) {
        ast.elements.push(currentParagraph);
      }

      // If document ended up empty, ensure at least one paragraph
      if (ast.elements.length === 0) {
        ast.elements.push({
          type: "paragraph",
          runs: [{ text: " " }],
          align: "left",
        });
      }

      return ast;
    },
  };

  // -------------------------------------------------------------
  // HTML Parser
  // -------------------------------------------------------------
  var HtmlParser = {
    parse: function (htmlString, filename) {
      var parser = new DOMParser();
      var doc = parser.parseFromString(htmlString, "text/html");
      var ast = createDocumentAst(filename ? filename.replace(/\.[^.]+$/, "") : "HTML Document");

      var body = doc.body || doc.documentElement;
      HtmlParser.walkNodes(body, ast, {});
      return ast;
    },

    walkNodes: function (parent, ast, parentStyle) {
      var nodes = parent.childNodes;
      for (var i = 0; i < nodes.length; i++) {
        var node = nodes[i];
        if (node.nodeType === 1) {
          var tag = node.tagName.toLowerCase();

          if (/^h[1-6]$/.test(tag)) {
            var level = parseInt(tag[1], 10);
            var hRuns = HtmlParser.extractRuns(node, parentStyle);
            ast.elements.push({
              type: "heading",
              level: level,
              runs: hRuns,
              align: node.style.textAlign || "left",
            });
          } else if (tag === "p" || tag === "div") {
            var pRuns = HtmlParser.extractRuns(node, parentStyle);
            if (pRuns.length > 0) {
              ast.elements.push({
                type: "paragraph",
                runs: pRuns,
                align: node.style.textAlign || "left",
              });
            } else {
              // Might contain block children
              HtmlParser.walkNodes(node, ast, parentStyle);
            }
          } else if (tag === "ul" || tag === "ol") {
            var ordered = tag === "ol";
            var lis = node.querySelectorAll("li");
            for (var l = 0; l < lis.length; l++) {
              var liRuns = HtmlParser.extractRuns(lis[l], parentStyle);
              ast.elements.push({
                type: "list-item",
                level: 0,
                ordered: ordered,
                runs: liRuns,
              });
            }
          } else if (tag === "table") {
            var tableAst = HtmlParser.extractTable(node, parentStyle);
            if (tableAst) ast.elements.push(tableAst);
          } else if (tag === "blockquote") {
            var bqRuns = HtmlParser.extractRuns(node, Object.assign({}, parentStyle, { italic: true, color: "666666" }));
            ast.elements.push({
              type: "paragraph",
              runs: bqRuns,
              align: "left",
              textIndent: 720,
            });
          } else if (tag === "pre" || tag === "code") {
            var codeText = node.textContent || "";
            ast.elements.push({
              type: "paragraph",
              runs: [{ text: codeText, fontFamily: "Consolas", fontSizePt: 10, bgColor: "F3F4F6" }],
              align: "left",
            });
          } else if (tag === "hr") {
            ast.elements.push({ type: "page-break" });
          } else {
            HtmlParser.walkNodes(node, ast, parentStyle);
          }
        }
      }
    },

    extractRuns: function (element, parentStyle) {
      var runs = [];
      var nodes = element.childNodes;

      for (var i = 0; i < nodes.length; i++) {
        var node = nodes[i];
        if (node.nodeType === 3) {
          var t = node.nodeValue;
          if (t && t.length > 0) {
            runs.push(Object.assign({ text: t }, parentStyle || {}));
          }
        } else if (node.nodeType === 1) {
          var tag = node.tagName.toLowerCase();
          var style = Object.assign({}, parentStyle || {});

          if (tag === "b" || tag === "strong") style.bold = true;
          if (tag === "i" || tag === "em") style.italic = true;
          if (tag === "u") style.underline = true;
          if (tag === "s" || tag === "del" || tag === "strike") style.strike = true;
          if (tag === "mark") style.bgColor = "FFFF00";
          if (tag === "a") {
            style.link = node.getAttribute("href");
            style.underline = true;
            style.color = "0066CC";
          }

          if (node.style.color) {
            var c = normalizeHexColor(node.style.color);
            if (c) style.color = c;
          }
          if (node.style.backgroundColor) {
            var bg = normalizeHexColor(node.style.backgroundColor);
            if (bg) style.bgColor = bg;
          }

          var subRuns = HtmlParser.extractRuns(node, style);
          runs.push.apply(runs, subRuns);
        }
      }
      return runs;
    },

    extractTable: function (tableNode, parentStyle) {
      var rows = [];
      var trs = tableNode.querySelectorAll("tr");
      for (var r = 0; r < trs.length; r++) {
        var cells = [];
        var tds = trs[r].querySelectorAll("th, td");
        for (var c = 0; c < tds.length; c++) {
          var td = tds[c];
          var colSpan = parseInt(td.getAttribute("colspan") || "1", 10);
          var rowSpan = parseInt(td.getAttribute("rowspan") || "1", 10);
          var isHeader = td.tagName.toLowerCase() === "th";
          var cellRuns = HtmlParser.extractRuns(td, Object.assign({}, parentStyle, isHeader ? { bold: true } : {}));

          cells.push({
            colSpan: isNaN(colSpan) ? 1 : colSpan,
            rowSpan: isNaN(rowSpan) ? 1 : rowSpan,
            bgColor: isHeader ? "F3F4F6" : null,
            elements: [
              {
                type: "paragraph",
                runs: cellRuns.length > 0 ? cellRuns : [{ text: "" }],
                align: td.style.textAlign || (isHeader ? "center" : "left"),
              },
            ],
          });
        }
        if (cells.length > 0) rows.push({ cells: cells });
      }

      if (rows.length === 0) return null;
      return { type: "table", rows: rows };
    },
  };

  // -------------------------------------------------------------
  // Markdown Parser
  // -------------------------------------------------------------
  var MarkdownParser = {
    parse: function (mdText, filename) {
      var ast = createDocumentAst(filename ? filename.replace(/\.[^.]+$/, "") : "Markdown Document");
      var lines = String(mdText).split(/\r?\n/);
      var i = 0;

      while (i < lines.length) {
        var line = lines[i];
        var trimmed = line.trim();

        if (!trimmed) {
          i++;
          continue;
        }

        // Heading # ...
        var headingMatch = trimmed.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) {
          var level = headingMatch[1].length;
          ast.elements.push({
            type: "heading",
            level: level,
            runs: MarkdownParser.parseInline(headingMatch[2]),
            align: "left",
          });
          i++;
          continue;
        }

        // Horizontal rule
        if (/^(\*{3,}|-{3,}|_{3,})$/.test(trimmed)) {
          ast.elements.push({ type: "page-break" });
          i++;
          continue;
        }

        // Code block ```
        if (trimmed.startsWith("```")) {
          i++;
          var codeLines = [];
          while (i < lines.length && !lines[i].trim().startsWith("```")) {
            codeLines.push(lines[i]);
            i++;
          }
          if (i < lines.length) i++; // skip closing ```
          ast.elements.push({
            type: "paragraph",
            runs: [
              {
                text: codeLines.join("\n"),
                fontFamily: "Consolas",
                fontSizePt: 10,
                bgColor: "F3F4F6",
              },
            ],
            align: "left",
          });
          continue;
        }

        // Blockquote >
        if (trimmed.startsWith(">")) {
          var bqText = trimmed.replace(/^>\s*/, "");
          ast.elements.push({
            type: "paragraph",
            runs: MarkdownParser.parseInline(bqText, { italic: true, color: "555555" }),
            align: "left",
            textIndent: 720,
          });
          i++;
          continue;
        }

        // Unordered list
        var ulMatch = trimmed.match(/^[-*+]\s+(.+)$/);
        if (ulMatch) {
          ast.elements.push({
            type: "list-item",
            level: 0,
            ordered: false,
            runs: MarkdownParser.parseInline(ulMatch[1]),
          });
          i++;
          continue;
        }

        // Ordered list
        var olMatch = trimmed.match(/^([0-9]+)\.\s+(.+)$/);
        if (olMatch) {
          ast.elements.push({
            type: "list-item",
            level: 0,
            ordered: true,
            runs: MarkdownParser.parseInline(olMatch[2]),
          });
          i++;
          continue;
        }

        // Table (| a | b |)
        if (trimmed.startsWith("|") && trimmed.endsWith("|")) {
          var tableLines = [];
          while (i < lines.length && lines[i].trim().startsWith("|") && lines[i].trim().endsWith("|")) {
            tableLines.push(lines[i].trim());
            i++;
          }
          var tableAst = MarkdownParser.parseTable(tableLines);
          if (tableAst) ast.elements.push(tableAst);
          continue;
        }

        // Plain paragraph
        ast.elements.push({
          type: "paragraph",
          runs: MarkdownParser.parseInline(trimmed),
          align: "left",
        });
        i++;
      }

      return ast;
    },

    parseInline: function (text, baseStyle) {
      // Simple inline tokenizer for **bold**, *italic*, ~~strike~~, `code`, [link](url)
      var runs = [];
      var remaining = text;

      while (remaining.length > 0) {
        // Link [text](url)
        var linkMatch = remaining.match(/^\[([^\]]+)\]\(([^)]+)\)/);
        if (linkMatch) {
          runs.push(
            Object.assign({}, baseStyle || {}, {
              text: linkMatch[1],
              link: linkMatch[2],
              underline: true,
              color: "0066CC",
            })
          );
          remaining = remaining.slice(linkMatch[0].length);
          continue;
        }

        // Bold **text**
        var boldMatch = remaining.match(/^\*\*([^*]+)\*\*/);
        if (boldMatch) {
          runs.push(Object.assign({}, baseStyle || {}, { text: boldMatch[1], bold: true }));
          remaining = remaining.slice(boldMatch[0].length);
          continue;
        }

        // Italic *text*
        var italicMatch = remaining.match(/^\*([^*]+)\*/);
        if (italicMatch) {
          runs.push(Object.assign({}, baseStyle || {}, { text: italicMatch[1], italic: true }));
          remaining = remaining.slice(italicMatch[0].length);
          continue;
        }

        // Strike ~~text~~
        var strikeMatch = remaining.match(/^~~([^~]+)~~/);
        if (strikeMatch) {
          runs.push(Object.assign({}, baseStyle || {}, { text: strikeMatch[1], strike: true }));
          remaining = remaining.slice(strikeMatch[0].length);
          continue;
        }

        // Inline code `text`
        var codeMatch = remaining.match(/^`([^`]+)`/);
        if (codeMatch) {
          runs.push(
            Object.assign({}, baseStyle || {}, {
              text: codeMatch[1],
              fontFamily: "Consolas",
              bgColor: "F3F4F6",
            })
          );
          remaining = remaining.slice(codeMatch[0].length);
          continue;
        }

        // Take regular characters until next special token
        var nextSpecial = remaining.search(/(\*\*|\*|~~|`|\[)/);
        if (nextSpecial === -1) {
          runs.push(Object.assign({}, baseStyle || {}, { text: remaining }));
          break;
        } else if (nextSpecial === 0) {
          // Special char not matching a pair, consume one
          runs.push(Object.assign({}, baseStyle || {}, { text: remaining[0] }));
          remaining = remaining.slice(1);
        } else {
          runs.push(Object.assign({}, baseStyle || {}, { text: remaining.slice(0, nextSpecial) }));
          remaining = remaining.slice(nextSpecial);
        }
      }

      return runs;
    },

    parseTable: function (tableLines) {
      if (tableLines.length < 2) return null;
      var rows = [];

      for (var r = 0; r < tableLines.length; r++) {
        var line = tableLines[r];
        // Skip separator row |---|---|
        if (/^\|[\s\-:|]+\|$/.test(line)) continue;

        var rawCells = line
          .slice(1, -1)
          .split("|")
          .map(function (c) {
            return c.trim();
          });
        var isHeader = r === 0;
        var cells = rawCells.map(function (content) {
          return {
            colSpan: 1,
            rowSpan: 1,
            bgColor: isHeader ? "F3F4F6" : null,
            elements: [
              {
                type: "paragraph",
                runs: MarkdownParser.parseInline(content, isHeader ? { bold: true } : {}),
                align: isHeader ? "center" : "left",
              },
            ],
          };
        });

        rows.push({ cells: cells });
      }

      if (rows.length === 0) return null;
      return { type: "table", rows: rows };
    },
  };

  // -------------------------------------------------------------
  // Plain Text & CSV Parser
  // -------------------------------------------------------------
  var PlainTextParser = {
    parse: function (content, filename) {
      var ast = createDocumentAst(filename ? filename.replace(/\.[^.]+$/, "") : "Text Document");
      var text = String(content);

      // Check if file is CSV / TSV
      var ext = filename ? filename.split(".").pop().toLowerCase() : "";
      if (ext === "csv" || ext === "tsv") {
        var sep = ext === "tsv" ? "\t" : ",";
        var lines = text.split(/\r?\n/).filter(function (l) {
          return l.trim().length > 0;
        });
        if (lines.length > 0) {
          var rows = lines.map(function (l, rIdx) {
            var raw = l.split(sep);
            var isH = rIdx === 0;
            return {
              cells: raw.map(function (item) {
                return {
                  colSpan: 1,
                  rowSpan: 1,
                  bgColor: isH ? "F3F4F6" : null,
                  elements: [
                    {
                      type: "paragraph",
                      runs: [{ text: item.trim(), bold: isH }],
                      align: isH ? "center" : "left",
                    },
                  ],
                };
              }),
            };
          });
          ast.elements.push({ type: "table", rows: rows });
          return ast;
        }
      }

      // Plain text: split into paragraphs by double newlines or single newlines
      var paras = text.split(/\r?\n\r?\n+/);
      for (var p = 0; p < paras.length; p++) {
        var pText = paras[p].trim();
        if (pText) {
          ast.elements.push({
            type: "paragraph",
            runs: [{ text: pText }],
            align: "left",
          });
        }
      }

      if (ast.elements.length === 0) {
        ast.elements.push({
          type: "paragraph",
          runs: [{ text: "" }],
          align: "left",
        });
      }

      return ast;
    },
  };

  // -------------------------------------------------------------
  // DOCX Generator (Office Open XML via JSZip)
  // -------------------------------------------------------------
  var DocxGenerator = {
    generate: async function (ast, options) {
      if (!global.JSZip) throw new Error("JSZip is required to generate DOCX files.");
      options = options || {};

      var zip = new global.JSZip();
      var mediaList = [];
      var hyperlinks = [];

      // Scan AST for images and hyperlinks to assign relationship IDs
      DocxGenerator.collectResources(ast, mediaList, hyperlinks);

      // 1. [Content_Types].xml
      zip.file("[Content_Types].xml", DocxGenerator.buildContentTypes(mediaList));

      // 2. _rels/.rels
      zip.file("_rels/.rels", DocxGenerator.buildRootRels());

      // 3. docProps/core.xml & app.xml
      zip.file("docProps/core.xml", DocxGenerator.buildCoreProps(ast.title));
      zip.file("docProps/app.xml", DocxGenerator.buildAppProps());

      // 4. word/_rels/document.xml.rels
      zip.file("word/_rels/document.xml.rels", DocxGenerator.buildDocumentRels(mediaList, hyperlinks));

      // 5. word/styles.xml
      zip.file("word/styles.xml", DocxGenerator.buildStyles(options));

      // 6. word/document.xml
      zip.file("word/document.xml", DocxGenerator.buildDocumentXml(ast, options, mediaList, hyperlinks));

      // 7. word/media/
      for (var m = 0; m < mediaList.length; m++) {
        var item = mediaList[m];
        zip.file("word/media/" + item.filename, item.data);
      }

      // Generate blob
      var blob = await zip.generateAsync({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      return blob;
    },

    collectResources: function (ast, mediaList, hyperlinks) {
      var imgCounter = 1;
      var linkCounter = 1;

      function inspectElements(elements) {
        for (var i = 0; i < elements.length; i++) {
          var el = elements[i];
          if (el.type === "image" && el.data) {
            var ext = el.ext || "png";
            var fname = "image" + imgCounter + "." + ext;
            var rId = "rIdImg" + imgCounter;
            imgCounter++;
            el.rId = rId;
            el.filename = fname;
            mediaList.push({
              rId: rId,
              filename: fname,
              data: el.data,
              mimeType: el.mimeType,
            });
          } else if (el.type === "table" && el.rows) {
            for (var r = 0; r < el.rows.length; r++) {
              for (var c = 0; c < el.rows[r].cells.length; c++) {
                inspectElements(el.rows[r].cells[c].elements);
              }
            }
          }

          if (el.runs) {
            for (var runIdx = 0; runIdx < el.runs.length; runIdx++) {
              var run = el.runs[runIdx];
              if (run.link) {
                var linkRId = "rIdLink" + linkCounter;
                linkCounter++;
                run.rId = linkRId;
                hyperlinks.push({
                  rId: linkRId,
                  target: run.link,
                });
              }
            }
          }
        }
      }

      inspectElements(ast.elements);
    },

    buildContentTypes: function (mediaList) {
      var extMap = {
        rels: "application/vnd.openxmlformats-package.relationships+xml",
        xml: "application/xml",
        png: "image/png",
        jpeg: "image/jpeg",
        jpg: "image/jpeg",
        gif: "image/gif",
        webp: "image/webp",
      };

      for (var i = 0; i < mediaList.length; i++) {
        var ext = mediaList[i].filename.split(".").pop().toLowerCase();
        if (!extMap[ext]) extMap[ext] = mediaList[i].mimeType || "application/octet-stream";
      }

      var xml = XML_DECLARATION + '\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';
      for (var extKey in extMap) {
        xml += '  <Default Extension="' + extKey + '" ContentType="' + extMap[extKey] + '"/>\n';
      }
      xml += '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n';
      xml += '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n';
      xml += '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n';
      xml += '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n';
      xml += "</Types>";
      return xml;
    },

    buildRootRels: function () {
      return (
        XML_DECLARATION +
        '\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n' +
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\n' +
        '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\n' +
        "</Relationships>"
      );
    },

    buildCoreProps: function (title) {
      var dateStr = new Date().toISOString();
      return (
        XML_DECLARATION +
        '\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' +
        "  <dc:title>" +
        escapeXml(title || "Document") +
        "</dc:title>\n" +
        "  <dc:creator>KenEasy Word Converter</dc:creator>\n" +
        "  <cp:lastModifiedBy>KenEasy Word Converter</cp:lastModifiedBy>\n" +
        '  <dcterms:created xsi:type="dcterms:W3CDTF">' +
        dateStr +
        "</dcterms:created>\n" +
        '  <dcterms:modified xsi:type="dcterms:W3CDTF">' +
        dateStr +
        "</dcterms:modified>\n" +
        "</cp:coreProperties>"
      );
    },

    buildAppProps: function () {
      return (
        XML_DECLARATION +
        '\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n' +
        "  <Application>KenEasy Word Converter</Application>\n" +
        "  <DocSecurity>0</DocSecurity>\n" +
        "  <Company>KenEasy Open Source</Company>\n" +
        "</Properties>"
      );
    },

    buildDocumentRels: function (mediaList, hyperlinks) {
      var xml =
        XML_DECLARATION +
        '\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n';

      for (var i = 0; i < mediaList.length; i++) {
        var m = mediaList[i];
        xml +=
          '  <Relationship Id="' +
          m.rId +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' +
          m.filename +
          '"/>\n';
      }

      for (var j = 0; j < hyperlinks.length; j++) {
        var h = hyperlinks[j];
        xml +=
          '  <Relationship Id="' +
          h.rId +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' +
          escapeXml(h.target) +
          '" TargetMode="External"/>\n';
      }

      xml += "</Relationships>";
      return xml;
    },

    buildStyles: function (options) {
      var latinFont = "Calibri";
      var cjkFont = "Microsoft YaHei";
      if (options.typography === "classic") {
        latinFont = "Times New Roman";
        cjkFont = "SimSun";
      } else if (options.typography === "sans") {
        latinFont = "Arial";
        cjkFont = "SimHei";
      }

      return (
        XML_DECLARATION +
        '\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n' +
        '  <w:docDefaults>\n' +
        '    <w:rPrDefault>\n' +
        '      <w:rPr>\n' +
        '        <w:rFonts w:ascii="' +
        latinFont +
        '" w:hAnsi="' +
        latinFont +
        '" w:eastAsia="' +
        cjkFont +
        '"/>\n' +
        '        <w:sz w:val="24"/>\n' +
        '        <w:color w:val="222222"/>\n' +
        "      </w:rPr>\n" +
        "    </w:rPrDefault>\n" +
        "  </w:docDefaults>\n" +
        '  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">\n' +
        '    <w:name w:val="Normal"/>\n' +
        "    <w:pPr>\n" +
        '      <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>\n' +
        "    </w:pPr>\n" +
        "  </w:style>\n" +
        '  <w:style w:type="paragraph" w:styleId="Heading1">\n' +
        '    <w:name w:val="heading 1"/>\n' +
        '    <w:basedOn w:val="Normal"/>\n' +
        "    <w:pPr>\n" +
        '      <w:spacing w:before="360" w:after="160"/>\n' +
        "    </w:pPr>\n" +
        "    <w:rPr>\n" +
        "      <w:b/>\n" +
        '      <w:sz w:val="36"/>\n' +
        '      <w:color w:val="1F2937"/>\n' +
        "    </w:rPr>\n" +
        "  </w:style>\n" +
        '  <w:style w:type="paragraph" w:styleId="Heading2">\n' +
        '    <w:name w:val="heading 2"/>\n' +
        '    <w:basedOn w:val="Normal"/>\n' +
        "    <w:pPr>\n" +
        '      <w:spacing w:before="280" w:after="120"/>\n' +
        "    </w:pPr>\n" +
        "    <w:rPr>\n" +
        "      <w:b/>\n" +
        '      <w:sz w:val="30"/>\n' +
        '      <w:color w:val="1F2937"/>\n' +
        "    </w:rPr>\n" +
        "  </w:style>\n" +
        '  <w:style w:type="paragraph" w:styleId="Heading3">\n' +
        '    <w:name w:val="heading 3"/>\n' +
        '    <w:basedOn w:val="Normal"/>\n' +
        "    <w:pPr>\n" +
        '      <w:spacing w:before="240" w:after="100"/>\n' +
        "    </w:pPr>\n" +
        "    <w:rPr>\n" +
        "      <w:b/>\n" +
        '      <w:sz w:val="26"/>\n' +
        '      <w:color w:val="374151"/>\n' +
        "    </w:rPr>\n" +
        "  </w:style>\n" +
        "</w:styles>"
      );
    },

    buildDocumentXml: function (ast, options, mediaList, hyperlinks) {
      // Dimensions
      var pageSize = { widthDxa: 11906, heightDxa: 16838 }; // Default A4
      if (options.pageSize === "letter") pageSize = { widthDxa: 12240, heightDxa: 15840 };
      else if (options.pageSize === "a5") pageSize = { widthDxa: 8390, heightDxa: 11906 };
      else if (options.pageSize === "b5") pageSize = { widthDxa: 9978, heightDxa: 14173 };

      var margins = { top: 1440, bottom: 1440, left: 1440, right: 1440 }; // 1 in
      if (options.margin === "narrow") margins = { top: 720, bottom: 720, left: 720, right: 720 };
      else if (options.margin === "wide") margins = { top: 1440, bottom: 1440, left: 2880, right: 2880 };

      var xml =
        XML_DECLARATION +
        '\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' +
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
        ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">\n' +
        "  <w:body>\n";

      for (var i = 0; i < ast.elements.length; i++) {
        xml += DocxGenerator.renderElement(ast.elements[i], options);
      }

      // Final section properties (page size, orientation, margins)
      xml +=
        "    <w:sectPr>\n" +
        '      <w:pgSz w:w="' +
        pageSize.widthDxa +
        '" w:h="' +
        pageSize.heightDxa +
        '"/>\n' +
        '      <w:pgMar w:top="' +
        margins.top +
        '" w:right="' +
        margins.right +
        '" w:bottom="' +
        margins.bottom +
        '" w:left="' +
        margins.left +
        '" w:header="720" w:footer="720" w:gutter="0"/>\n' +
        "    </w:sectPr>\n";

      xml += "  </w:body>\n</w:document>";
      return xml;
    },

    renderElement: function (el, options) {
      if (el.type === "heading") {
        var hLevel = el.level || 1;
        var pStyle = "Heading" + (hLevel <= 3 ? hLevel : 3);
        var pXml = "    <w:p>\n      <w:pPr>\n";
        pXml += '        <w:pStyle w:val="' + pStyle + '"/>\n';
        if (el.align && el.align !== "left") {
          var jcVal = el.align === "justify" ? "both" : el.align;
          pXml += '        <w:jc w:val="' + jcVal + '"/>\n';
        }
        pXml += "      </w:pPr>\n";
        pXml += DocxGenerator.renderRuns(el.runs);
        pXml += "    </w:p>\n";
        return pXml;
      }

      if (el.type === "paragraph") {
        var paraXml = "    <w:p>\n      <w:pPr>\n";
        if (el.align && el.align !== "left") {
          var jc = el.align === "justify" ? "both" : el.align;
          paraXml += '        <w:jc w:val="' + jc + '"/>\n';
        }
        var sb = el.spacingBefore || 0;
        var sa = typeof el.spacingAfter === "number" ? el.spacingAfter : 160;
        paraXml += '        <w:spacing w:before="' + sb + '" w:after="' + sa + '" w:line="276" w:lineRule="auto"/>\n';

        if (options.firstLineIndent || el.textIndent) {
          var indentVal = el.textIndent || 480; // 2 chars / 480 dxa
          paraXml += '        <w:ind w:firstLine="' + indentVal + '"/>\n';
        }

        paraXml += "      </w:pPr>\n";
        paraXml += DocxGenerator.renderRuns(el.runs);
        paraXml += "    </w:p>\n";
        return paraXml;
      }

      if (el.type === "list-item") {
        var lLevel = el.level || 0;
        var listXml = "    <w:p>\n      <w:pPr>\n";
        var leftIndent = 720 + lLevel * 360;
        listXml += '        <w:ind w:left="' + leftIndent + '" w:hanging="360"/>\n';
        listXml += "      </w:pPr>\n";
        var bullet = el.ordered ? (lLevel + 1) + ". " : "• ";
        listXml += '      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + escapeXml(bullet) + "</w:t></w:r>\n";
        listXml += DocxGenerator.renderRuns(el.runs);
        listXml += "    </w:p>\n";
        return listXml;
      }

      if (el.type === "image" && el.rId) {
        var wEmu = el.widthEmu || 4320000;
        var hEmu = el.heightEmu || 2880000;
        var docPrId = parseInt(el.rId.replace(/\D/g, "") || "1", 10);
        return (
          "    <w:p>\n" +
          "      <w:pPr><w:jc w:val=\"center\"/></w:pPr>\n" +
          "      <w:r>\n" +
          "        <w:drawing>\n" +
          '          <wp:inline distT="0" distB="0" distL="0" distR="0">\n' +
          '            <wp:extent cx="' +
          wEmu +
          '" cy="' +
          hEmu +
          '"/>\n' +
          '            <wp:docPr id="' +
          docPrId +
          '" name="' +
          escapeXml(el.name || "Picture") +
          '"/>\n' +
          '            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>\n' +
          '            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n' +
          '              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">\n' +
          '                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">\n' +
          "                  <pic:nvPicPr>\n" +
          '                    <pic:cNvPr id="' +
          docPrId +
          '" name="' +
          escapeXml(el.name || "Image") +
          '"/>\n' +
          "                    <pic:cNvPicPr/>\n" +
          "                  </pic:nvPicPr>\n" +
          "                  <pic:blipFill>\n" +
          '                    <a:blip r:embed="' +
          el.rId +
          '"/>\n' +
          "                    <a:stretch><a:fillRect/></a:stretch>\n" +
          "                  </pic:blipFill>\n" +
          "                  <pic:spPr>\n" +
          '                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="' +
          wEmu +
          '" cy="' +
          hEmu +
          '"/></a:xfrm>\n' +
          '                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n' +
          "                  </pic:spPr>\n" +
          "                </pic:pic>\n" +
          "              </a:graphicData>\n" +
          "            </a:graphic>\n" +
          "          </wp:inline>\n" +
          "        </w:drawing>\n" +
          "      </w:r>\n" +
          "    </w:p>\n"
        );
      }

      if (el.type === "table" && el.rows) {
        var tblXml = "    <w:tbl>\n      <w:tblPr>\n";
        tblXml += '        <w:tblW w:w="0" w:type="auto"/>\n';
        tblXml += "        <w:tblBorders>\n";
        tblXml += '          <w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n';
        tblXml += '          <w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n';
        tblXml += '          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n';
        tblXml += '          <w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n';
        tblXml += '          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>\n';
        tblXml += '          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>\n';
        tblXml += "        </w:tblBorders>\n";
        tblXml += "      </w:tblPr>\n";

        for (var r = 0; r < el.rows.length; r++) {
          tblXml += "      <w:tr>\n";
          var row = el.rows[r];
          for (var c = 0; c < row.cells.length; c++) {
            var cell = row.cells[c];
            tblXml += "        <w:tc>\n          <w:tcPr>\n";
            if (cell.colSpan > 1) {
              tblXml += '            <w:gridSpan w:val="' + cell.colSpan + '"/>\n';
            }
            if (cell.bgColor) {
              tblXml += '            <w:shd w:val="clear" w:color="auto" w:fill="' + cell.bgColor + '"/>\n';
            }
            tblXml += "          </w:tcPr>\n";
            if (cell.elements && cell.elements.length > 0) {
              for (var ceIdx = 0; ceIdx < cell.elements.length; ceIdx++) {
                tblXml += DocxGenerator.renderElement(cell.elements[ceIdx], options);
              }
            } else {
              tblXml += "          <w:p/>\n";
            }
            tblXml += "        </w:tc>\n";
          }
          tblXml += "      </w:tr>\n";
        }
        tblXml += "    </w:tbl>\n";
        return tblXml;
      }

      if (el.type === "page-break") {
        return "    <w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>\n";
      }

      return "";
    },

    renderRuns: function (runs) {
      if (!runs || runs.length === 0) return "";
      var xml = "";
      for (var i = 0; i < runs.length; i++) {
        var r = runs[i];
        var text = r.text || "";

        if (r.breakLine) {
          xml += "      <w:r><w:br/></w:r>\n";
          continue;
        }

        var runBody = "";
        runBody += "        <w:rPr>\n";
        if (r.bold) runBody += "          <w:b/>\n";
        if (r.italic) runBody += "          <w:i/>\n";
        if (r.underline) runBody += '          <w:u w:val="single"/>\n';
        if (r.strike) runBody += "          <w:strike/>\n";
        if (r.color) runBody += '          <w:color w:val="' + r.color + '"/>\n';
        if (r.bgColor) runBody += '          <w:shd w:val="clear" w:color="auto" w:fill="' + r.bgColor + '"/>\n';
        if (r.fontSizePt) {
          runBody += '          <w:sz w:val="' + Math.round(r.fontSizePt * 2) + '"/>\n';
        }
        if (r.fontFamily) {
          runBody +=
            '          <w:rFonts w:ascii="' +
            escapeXml(r.fontFamily) +
            '" w:hAnsi="' +
            escapeXml(r.fontFamily) +
            '"/>\n';
        }
        runBody += "        </w:rPr>\n";
        runBody += '        <w:t xml:space="preserve">' + escapeXml(text) + "</w:t>\n";

        if (r.link && r.rId) {
          xml += '      <w:hyperlink r:id="' + r.rId + '">\n';
          xml += "        <w:r>\n" + runBody + "        </w:r>\n";
          xml += "      </w:hyperlink>\n";
        } else {
          xml += "      <w:r>\n" + runBody + "      </w:r>\n";
        }
      }
      return xml;
    },
  };

  // -------------------------------------------------------------
  // Plain Text & Markdown Exporters
  // -------------------------------------------------------------
  var TextExporter = {
    toPlainText: function (ast) {
      var lines = [];
      for (var i = 0; i < ast.elements.length; i++) {
        var el = ast.elements[i];
        if (el.type === "heading") {
          var prefix = "";
          for (var h = 0; h < (el.level || 1); h++) prefix += "#";
          lines.push(prefix + " " + TextExporter.runsToText(el.runs));
          lines.push("");
        } else if (el.type === "paragraph") {
          lines.push(TextExporter.runsToText(el.runs));
          lines.push("");
        } else if (el.type === "list-item") {
          var indent = "";
          for (var l = 0; l < (el.level || 0); l++) indent += "  ";
          var mark = el.ordered ? "1. " : "• ";
          lines.push(indent + mark + TextExporter.runsToText(el.runs));
        } else if (el.type === "table" && el.rows) {
          lines.push(TextExporter.tableToAscii(el));
          lines.push("");
        } else if (el.type === "page-break") {
          lines.push("\n--- [Page Break] ---\n");
        }
      }
      return lines.join("\n").trim();
    },

    toMarkdown: function (ast) {
      var lines = [];
      for (var i = 0; i < ast.elements.length; i++) {
        var el = ast.elements[i];
        if (el.type === "heading") {
          var prefix = "";
          for (var h = 0; h < (el.level || 1); h++) prefix += "#";
          lines.push(prefix + " " + TextExporter.runsToMarkdown(el.runs));
          lines.push("");
        } else if (el.type === "paragraph") {
          lines.push(TextExporter.runsToMarkdown(el.runs));
          lines.push("");
        } else if (el.type === "list-item") {
          var indent = "";
          for (var l = 0; l < (el.level || 0); l++) indent += "  ";
          var mark = el.ordered ? "1. " : "- ";
          lines.push(indent + mark + TextExporter.runsToMarkdown(el.runs));
        } else if (el.type === "table" && el.rows) {
          lines.push(TextExporter.tableToMarkdown(el));
          lines.push("");
        } else if (el.type === "page-break") {
          lines.push("\n---\n");
        }
      }
      return lines.join("\n").trim();
    },

    toHtml: function (ast) {
      var html = '<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<title>' +
        escapeXml(ast.title) +
        "</title>\n<style>\nbody{font-family:system-ui,-apple-system,sans-serif;max-width:860px;margin:2rem auto;padding:0 1.5rem;line-height:1.6;color:#1f2937;background:#fff;}\ntable{border-collapse:collapse;width:100%;margin:1rem 0;}\nth,td{border:1px solid #d1d5db;padding:8px 12px;text-align:left;}\nth{background:#f3f4f6;}\ncode,pre{background:#f3f4f6;padding:2px 6px;border-radius:4px;font-family:monospace;}\nhr{border:0;border-top:1px dashed #d1d5db;margin:2rem 0;}\n</style>\n</head>\n<body>\n";

      for (var i = 0; i < ast.elements.length; i++) {
        var el = ast.elements[i];
        if (el.type === "heading") {
          var tag = "h" + (el.level || 1);
          html += "<" + tag + ">" + TextExporter.runsToHtml(el.runs) + "</" + tag + ">\n";
        } else if (el.type === "paragraph") {
          html += "<p>" + TextExporter.runsToHtml(el.runs) + "</p>\n";
        } else if (el.type === "list-item") {
          html += "<ul><li>" + TextExporter.runsToHtml(el.runs) + "</li></ul>\n";
        } else if (el.type === "table" && el.rows) {
          html += "<table>\n";
          for (var r = 0; r < el.rows.length; r++) {
            html += "  <tr>\n";
            for (var c = 0; c < el.rows[r].cells.length; c++) {
              var cell = el.rows[r].cells[c];
              var cellTag = r === 0 ? "th" : "td";
              var cellText = "";
              if (cell.elements) {
                for (var ce = 0; ce < cell.elements.length; ce++) {
                  if (cell.elements[ce].runs) cellText += TextExporter.runsToHtml(cell.elements[ce].runs);
                }
              }
              html += "    <" + cellTag + ">" + cellText + "</" + cellTag + ">\n";
            }
            html += "  </tr>\n";
          }
          html += "</table>\n";
        } else if (el.type === "page-break") {
          html += "<hr>\n";
        }
      }

      html += "</body>\n</html>";
      return html;
    },

    runsToText: function (runs) {
      if (!runs) return "";
      return runs
        .map(function (r) {
          return r.text || "";
        })
        .join("");
    },

    runsToMarkdown: function (runs) {
      if (!runs) return "";
      return runs
        .map(function (r) {
          var t = r.text || "";
          if (r.bold && r.italic) t = "***" + t + "***";
          else if (r.bold) t = "**" + t + "**";
          else if (r.italic) t = "*" + t + "*";
          if (r.strike) t = "~~" + t + "~~";
          if (r.link) t = "[" + t + "](" + r.link + ")";
          return t;
        })
        .join("");
    },

    runsToHtml: function (runs) {
      if (!runs) return "";
      return runs
        .map(function (r) {
          var t = escapeXml(r.text || "");
          if (r.bold) t = "<strong>" + t + "</strong>";
          if (r.italic) t = "<em>" + t + "</em>";
          if (r.underline) t = "<u>" + t + "</u>";
          if (r.strike) t = "<del>" + t + "</del>";
          if (r.color) t = '<span style="color:#' + r.color + '">' + t + "</span>";
          if (r.link) t = '<a href="' + escapeXml(r.link) + '" target="_blank">' + t + "</a>";
          return t;
        })
        .join("");
    },

    tableToAscii: function (tableEl) {
      var lines = [];
      for (var r = 0; r < tableEl.rows.length; r++) {
        var row = tableEl.rows[r];
        var cellTexts = row.cells.map(function (c) {
          var txt = "";
          if (c.elements) {
            for (var i = 0; i < c.elements.length; i++) {
              if (c.elements[i].runs) txt += TextExporter.runsToText(c.elements[i].runs);
            }
          }
          return txt.replace(/\r?\n/g, " ");
        });
        lines.push("| " + cellTexts.join(" | ") + " |");
      }
      return lines.join("\n");
    },

    tableToMarkdown: function (tableEl) {
      if (tableEl.rows.length === 0) return "";
      var lines = [];
      var headerRow = tableEl.rows[0];
      var headerTexts = headerRow.cells.map(function (c) {
        var txt = "";
        if (c.elements) {
          for (var i = 0; i < c.elements.length; i++) {
            if (c.elements[i].runs) txt += TextExporter.runsToMarkdown(c.elements[i].runs);
          }
        }
        return txt;
      });
      lines.push("| " + headerTexts.join(" | ") + " |");

      var sep = headerTexts.map(function () {
        return "---";
      });
      lines.push("| " + sep.join(" | ") + " |");

      for (var r = 1; r < tableEl.rows.length; r++) {
        var row = tableEl.rows[r];
        var cellTexts = row.cells.map(function (c) {
          var txt = "";
          if (c.elements) {
            for (var i = 0; i < c.elements.length; i++) {
              if (c.elements[i].runs) txt += TextExporter.runsToMarkdown(c.elements[i].runs);
            }
          }
          return txt;
        });
        lines.push("| " + cellTexts.join(" | ") + " |");
      }
      return lines.join("\n");
    },
  };

  // -------------------------------------------------------------
  // Public Domain API: KenEasyWordEngine
  // -------------------------------------------------------------
  var engine = {
    detectKind: function (file) {
      var name = (file && file.name) || "";
      var ext = name.split(".").pop().toLowerCase();
      var type = (file && file.type) || "";

      // ODT
      if (ext === "odt" || type === "application/vnd.oasis.opendocument.text") return "odt";
      // FODT
      if (ext === "fodt") return "fodt";
      // RTF
      if (ext === "rtf" || type === "application/rtf" || type === "text/rtf") return "rtf";
      // HTML
      if (ext === "html" || ext === "htm" || ext === "xhtml" || type === "text/html") return "html";
      // MD
      if (ext === "md" || ext === "markdown" || type === "text/markdown") return "md";
      // DOCX
      if (ext === "docx" || type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") return "docx";
      // Legacy DOC / WPS (not convertible)
      if (ext === "doc" || ext === "wps") return "unsupported-doc";
      // TXT / CSV / etc.
      if (
        ext === "txt" ||
        ext === "csv" ||
        ext === "tsv" ||
        ext === "log" ||
        ext === "json" ||
        ext === "xml" ||
        ext === "opml" ||
        type.startsWith("text/")
      ) {
        return "text";
      }

      return "unknown";
    },

    parseFile: async function (file) {
      var kind = engine.detectKind(file);
      var filename = file.name || "document";

      if (kind === "odt") {
        var ab = await file.arrayBuffer();
        return await OdtParser.parseOdt(ab, filename);
      }
      if (kind === "fodt") {
        var textFodt = await file.text();
        return OdtParser.parseFodt(textFodt, filename);
      }
      if (kind === "rtf") {
        var textRtf = await file.text();
        return RtfParser.parse(textRtf, filename);
      }
      if (kind === "html") {
        var textHtml = await file.text();
        return HtmlParser.parse(textHtml, filename);
      }
      if (kind === "md") {
        var textMd = await file.text();
        return MarkdownParser.parse(textMd, filename);
      }
      if (kind === "text") {
        var textPlain = await file.text();
        return PlainTextParser.parse(textPlain, filename);
      }
      if (kind === "docx") {
        // Can read text from DOCX using JSZip
        var abDocx = await file.arrayBuffer();
        if (global.JSZip) {
          try {
            var zip = await global.JSZip.loadAsync(abDocx);
            var docXml = await zip.file("word/document.xml").async("string");
            var parser = new DOMParser();
            var xmlDoc = parser.parseFromString(docXml, "application/xml");
            var ast = createDocumentAst(filename.replace(/\.[^.]+$/, ""));
            var ps = xmlDoc.getElementsByTagNameNS("*", "p");
            for (var p = 0; p < ps.length; p++) {
              var pEl = ps[p];
              var runs = [];
              var rs = pEl.getElementsByTagNameNS("*", "r");
              for (var r = 0; r < rs.length; r++) {
                var rEl = rs[r];
                var tEl = rEl.getElementsByTagNameNS("*", "t")[0];
                if (tEl && tEl.textContent) {
                  var isBold = Boolean(rEl.getElementsByTagNameNS("*", "b")[0]);
                  var isItalic = Boolean(rEl.getElementsByTagNameNS("*", "i")[0]);
                  runs.push({ text: tEl.textContent, bold: isBold, italic: isItalic });
                }
              }
              if (runs.length > 0) {
                ast.elements.push({ type: "paragraph", runs: runs, align: "left" });
              }
            }
            return ast;
          } catch (e) {
            console.warn("Could not read DOCX XML directly:", e);
          }
        }
        var rawTxt = await file.text();
        return PlainTextParser.parse(rawTxt, filename);
      }

      throw new Error("Unsupported document format: " + filename);
    },

    convertAst: async function (ast, options) {
      options = options || {};
      var targetFormat = options.targetFormat || "docx";

      if (targetFormat === "docx") {
        var blob = await DocxGenerator.generate(ast, options);
        return {
          blob: blob,
          filename: (options.filename || ast.title || "converted") + ".docx",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        };
      }

      if (targetFormat === "txt") {
        var plain = TextExporter.toPlainText(ast);
        var txtBlob = new Blob([plain], { type: "text/plain; charset=utf-8" });
        return {
          blob: txtBlob,
          filename: (options.filename || ast.title || "converted") + ".txt",
          mimeType: "text/plain; charset=utf-8",
        };
      }

      if (targetFormat === "html") {
        var htmlContent = TextExporter.toHtml(ast);
        var htmlBlob = new Blob([htmlContent], { type: "text/html; charset=utf-8" });
        return {
          blob: htmlBlob,
          filename: (options.filename || ast.title || "converted") + ".html",
          mimeType: "text/html; charset=utf-8",
        };
      }

      if (targetFormat === "md") {
        var mdContent = TextExporter.toMarkdown(ast);
        var mdBlob = new Blob([mdContent], { type: "text/markdown; charset=utf-8" });
        return {
          blob: mdBlob,
          filename: (options.filename || ast.title || "converted") + ".md",
          mimeType: "text/markdown; charset=utf-8",
        };
      }

      throw new Error("Unknown target format: " + targetFormat);
    },

    mergeAsts: function (asts, targetTitle) {
      var merged = createDocumentAst(targetTitle || "Merged Document");
      for (var i = 0; i < asts.length; i++) {
        var itemAst = asts[i];
        if (i > 0) {
          // Page break between documents
          merged.elements.push({ type: "page-break" });
        }
        for (var e = 0; e < itemAst.elements.length; e++) {
          merged.elements.push(itemAst.elements[e]);
        }
      }
      return merged;
    },

    bundleZip: async function (fileList, zipName) {
      if (!global.JSZip) throw new Error("JSZip is required to create ZIP bundles.");
      var zip = new global.JSZip();
      for (var i = 0; i < fileList.length; i++) {
        var f = fileList[i];
        zip.file(f.name, f.blob);
      }
      return await zip.generateAsync({ type: "blob" });
    },
  };

  global.KenEasyWordEngine = Object.freeze(engine);
})(typeof window !== "undefined" ? window : globalThis);
