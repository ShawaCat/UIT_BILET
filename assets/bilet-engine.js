/*
 * bilet-engine.js — генерация экзаменационных билетов прямо в браузере.
 *
 * Это перенос логики bilety.py на JavaScript: скрипт читает .docx-шаблон и базу
 * вопросов, копирует таблицу-шаблон нужное число раз, подставляет значения в
 * заполнители вида <DISCIP>, вставляет случайные вопросы и собирает новый .docx.
 *
 * Всё выполняется локально: файлы никуда не отправляются.
 * Требуется JSZip (window.JSZip).
 */
(function (global) {
  "use strict";

  var W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  var XML_NS = "http://www.w3.org/XML/1998/namespace";
  var W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

  /* Настройки вёрстки — те же значения, что в bilety.py */
  var SETTINGS = {
    questionNumbering: "{n}. {text}",
    questionAlignment: "left",
    questionSpaceAfterTwips: 60,
    questionHangingIndentTwips: 340,
    ticketGapTwips: 300,
    pageSlotPadTwips: 140,
    lastRowSafetyTwips: 60,
    heightSafetyFactor: 1.03,
    forceTableBorders: true,
    tasksPlaceholder: "<BILET_TASK>",
    numberPlaceholder: "NUM"
  };

  var PLACEHOLDER_SOURCE = "<[A-ZА-ЯЁ][A-ZА-ЯЁ0-9_]*>";
  var PLACEHOLDER_ALL = new RegExp(PLACEHOLDER_SOURCE, "g");
  var PLACEHOLDER_STICKY = new RegExp(PLACEHOLDER_SOURCE, "y");
  var HEADING_HINTS = ["heading", "заголовок", "title", "название"];
  var NUMBER_PREFIX = /^\s*(?:\d{1,3}[.)]|[-–—•])\s+/;

  var LINE_HEIGHT_FACTORS = {
    "calibri": 1.2207,
    "carlito": 1.2207,
    "times new roman": 1.1499,
    "liberation serif": 1.1499,
    "arial": 1.1499,
    "liberation sans": 1.1499,
    "cambria": 1.17
  };
  var DEFAULT_LINE_HEIGHT_FACTOR = 1.2;
  var DEFAULT_FONT_SIZE_PT = 12;

  var NARROW_CHARS = "ijl.,;:'!|()[]{}«»/\\-–—";
  var WIDE_CHARS = "mwMWШЩЖЮФМфжющ";
  var DIGITS = "0123456789";
  var UPPER = (function () {
    var out = "";
    for (var c = 65; c <= 90; c++) out += String.fromCharCode(c);      // A-Z
    for (var d = 1040; d <= 1071; d++) out += String.fromCharCode(d);  // А-Я
    return out;
  })();

  /* Порядок дочерних элементов по схеме OOXML */
  var PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct",
    "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing",
    "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId",
    "cnfStyle", "rPr", "sectPr", "pPrChange"];

  var TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
    "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing",
    "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription", "tblPrChange"];

  function indexMap(order) {
    var map = {};
    for (var i = 0; i < order.length; i++) map[order[i]] = i;
    return map;
  }
  var PPR_INDEX = indexMap(PPR_ORDER);
  var TBLPR_INDEX = indexMap(TBLPR_ORDER);
  var LAST_INDEX = 9999;

  /* ------------------------------------------------------------------ */
  /* Работа с DOM                                                        */
  /* ------------------------------------------------------------------ */

  function elementChildren(node) {
    var out = [];
    if (!node) return out;
    for (var child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === 1) out.push(child);
    }
    return out;
  }

  function kids(node, local) {
    return elementChildren(node).filter(function (child) {
      return child.namespaceURI === W && child.localName === local;
    });
  }

  function kid(node, local) {
    return kids(node, local)[0] || null;
  }

  function descendants(node, local) {
    return Array.prototype.slice.call(node.getElementsByTagNameNS(W, local));
  }

  function mkEl(doc, local, attrs) {
    var node = doc.createElementNS(W, "w:" + local);
    if (attrs) {
      Object.keys(attrs).forEach(function (name) {
        node.setAttributeNS(W, "w:" + name, String(attrs[name]));
      });
    }
    return node;
  }

  function val(node) {
    var value = node.getAttributeNS(W, "val");
    return value === null ? null : value;
  }

  function intVal(node, fallback) {
    var value = node ? val(node) : null;
    var parsed = parseInt(value, 10);
    return isFinite(parsed) ? parsed : fallback;
  }

  function insertOrdered(parent, node, map) {
    var position = Object.prototype.hasOwnProperty.call(map, node.localName)
      ? map[node.localName] : LAST_INDEX;
    var children = elementChildren(parent);
    for (var i = 0; i < children.length; i++) {
      var child = children[i];
      var childPosition = Object.prototype.hasOwnProperty.call(map, child.localName)
        ? map[child.localName] : LAST_INDEX;
      if (childPosition > position) {
        parent.insertBefore(node, child);
        return;
      }
    }
    parent.appendChild(node);
  }

  function ancestor(node, local) {
    var current = node;
    while (current) {
      if (current.namespaceURI === W && current.localName === local) return current;
      current = current.parentNode;
    }
    return null;
  }

  function paragraphText(paragraph) {
    return descendants(paragraph, "t").map(function (node) {
      return node.textContent || "";
    }).join("");
  }

  function normalizeValues(values) {
    var out = {};
    Object.keys(values || {}).forEach(function (key) {
      out[String(key).replace(/[<>]/g, "").trim().toUpperCase()] = String(values[key]);
    });
    return out;
  }

  /** Заменяет заполнители в абзаце, сохраняя разбиение на run'ы. */
  function replacePlaceholders(paragraph, values) {
    var runs = kids(paragraph, "r");
    if (!runs.length) return false;

    var texts = [];
    var owners = [];
    runs.forEach(function (run, index) {
      var node = kid(run, "t");
      var text = node && node.textContent ? node.textContent : "";
      texts.push(text);
      for (var i = 0; i < text.length; i++) owners.push(index);
    });

    var full = texts.join("");
    if (full.indexOf("<") < 0) return false;

    var outChars = [];
    var position = 0;
    var changed = false;
    while (position < full.length) {
      PLACEHOLDER_STICKY.lastIndex = position;
      var match = PLACEHOLDER_STICKY.exec(full);
      var key = match ? match[0].slice(1, -1).toUpperCase() : "";
      if (match && Object.prototype.hasOwnProperty.call(values, key)) {
        var replacement = Array.from(String(values[key]));
        for (var r = 0; r < replacement.length; r++) outChars.push([replacement[r], owners[position]]);
        position = PLACEHOLDER_STICKY.lastIndex;
        changed = true;
      } else {
        outChars.push([full.charAt(position), owners[position]]);
        position += 1;
      }
    }
    if (!changed) return false;

    var buckets = runs.map(function () { return []; });
    outChars.forEach(function (pair) { buckets[pair[1]].push(pair[0]); });

    runs.forEach(function (run, index) {
      var text = buckets[index].join("");
      var node = kid(run, "t");
      if (!node) {
        if (!text) return;
        node = mkEl(run.ownerDocument, "t");
        run.appendChild(node);
      }
      node.textContent = text;
      if (text !== text.trim()) node.setAttributeNS(XML_NS, "xml:space", "preserve");
    });
    return true;
  }

  function replacePlaceholdersText(text, values) {
    Object.keys(values).forEach(function (key) {
      text = text.split("<" + key + ">").join(String(values[key]));
    });
    return text;
  }

  function stripVolatileIds(root) {
    var all = Array.prototype.slice.call(root.getElementsByTagNameNS("*", "*"));
    all.push(root);
    all.forEach(function (node) {
      if (node.removeAttributeNS) {
        node.removeAttributeNS(W14, "paraId");
        node.removeAttributeNS(W14, "textId");
      }
    });
  }

  /* ------------------------------------------------------------------ */
  /* Вопросы                                                             */
  /* ------------------------------------------------------------------ */

  function questionParagraph(baseParagraph, text) {
    var doc = baseParagraph.ownerDocument;
    var paragraph = mkEl(doc, "p");

    var basePPr = kid(baseParagraph, "pPr");
    var pPr = basePPr ? basePPr.cloneNode(true) : mkEl(doc, "pPr");
    ["jc", "spacing", "ind"].forEach(function (tag) {
      var node = kid(pPr, tag);
      if (node) pPr.removeChild(node);
    });

    insertOrdered(pPr, mkEl(doc, "spacing", {
      before: 0,
      after: SETTINGS.questionSpaceAfterTwips
    }), PPR_INDEX);

    if (SETTINGS.questionHangingIndentTwips > 0) {
      insertOrdered(pPr, mkEl(doc, "ind", {
        left: SETTINGS.questionHangingIndentTwips,
        hanging: SETTINGS.questionHangingIndentTwips
      }), PPR_INDEX);
    }
    if (["left", "center", "right", "both"].indexOf(SETTINGS.questionAlignment) >= 0) {
      insertOrdered(pPr, mkEl(doc, "jc", { val: SETTINGS.questionAlignment }), PPR_INDEX);
    }
    paragraph.appendChild(pPr);

    var run = mkEl(doc, "r");
    var baseRun = kid(baseParagraph, "r");
    if (baseRun) {
      var baseRPr = kid(baseRun, "rPr");
      if (baseRPr) {
        var rPr = baseRPr.cloneNode(true);
        kids(rPr, "lang").forEach(function (node) { rPr.removeChild(node); });
        if (elementChildren(rPr).length) run.appendChild(rPr);
      }
    }
    var textNode = mkEl(doc, "t");
    textNode.setAttributeNS(XML_NS, "xml:space", "preserve");
    textNode.textContent = text;
    run.appendChild(textNode);
    paragraph.appendChild(run);
    return paragraph;
  }

  function insertQuestions(table, questions, values) {
    var placeholder = null;
    var paragraphs = descendants(table, "p");
    for (var i = 0; i < paragraphs.length; i++) {
      if (paragraphText(paragraphs[i]).indexOf(SETTINGS.tasksPlaceholder) >= 0) {
        placeholder = paragraphs[i];
        break;
      }
    }
    if (!placeholder) {
      throw new Error("В шаблоне не найден заполнитель " + SETTINGS.tasksPlaceholder +
        " — именно в это место подставляются вопросы.");
    }

    var row = ancestor(placeholder, "tr");
    questions.forEach(function (question, index) {
      var text = SETTINGS.questionNumbering
        .replace("{n}", String(index + 1))
        .replace("{text}", question);
      text = replacePlaceholdersText(text, values).split("\n").join(" ");
      placeholder.parentNode.insertBefore(questionParagraph(placeholder, text), placeholder);
    });
    placeholder.parentNode.removeChild(placeholder);
    return row;
  }

  /* ------------------------------------------------------------------ */
  /* Оценка высоты                                                       */
  /* ------------------------------------------------------------------ */

  function readStyleDefaults(stylesXml, themeXml) {
    var defaults = {
      font_size_pt: DEFAULT_FONT_SIZE_PT,
      space_before: 0,
      space_after: 160,
      line: 276,
      line_rule: "auto",
      line_factor: DEFAULT_LINE_HEIGHT_FACTOR
    };
    if (themeXml) {
      var minor = /<a:minorFont>[\s\S]*?<a:latin typeface="([^"]*)"/.exec(themeXml);
      if (minor) {
        var factor = LINE_HEIGHT_FACTORS[minor[1].trim().toLowerCase()];
        if (factor) defaults.line_factor = factor;
      }
    }
    if (!stylesXml) return defaults;

    var styles = new DOMParser().parseFromString(stylesXml, "application/xml");
    var docDefaults = styles.getElementsByTagNameNS(W, "docDefaults")[0];
    if (!docDefaults) return defaults;

    var rPrDefault = docDefaults.getElementsByTagNameNS(W, "rPrDefault")[0];
    if (rPrDefault) {
      var size = rPrDefault.getElementsByTagNameNS(W, "sz")[0];
      if (size) defaults.font_size_pt = intVal(size, 24) / 2;
    }
    var pPrDefault = docDefaults.getElementsByTagNameNS(W, "pPrDefault")[0];
    if (pPrDefault) {
      var spacing = pPrDefault.getElementsByTagNameNS(W, "spacing")[0];
      if (spacing) {
        defaults.space_before = parseFloat(spacing.getAttributeNS(W, "before") || 0) || 0;
        defaults.space_after = parseFloat(spacing.getAttributeNS(W, "after") || 0) || 0;
        defaults.line = parseFloat(spacing.getAttributeNS(W, "line") || 240) || 240;
        defaults.line_rule = spacing.getAttributeNS(W, "lineRule") || "auto";
      }
    }
    return defaults;
  }

  function charWidthEm(character) {
    if (character === " ") return 0.226;
    if (DIGITS.indexOf(character) >= 0) return 0.51;
    if (NARROW_CHARS.indexOf(character) >= 0) return 0.30;
    if (WIDE_CHARS.indexOf(character) >= 0) return 0.85;
    if (UPPER.indexOf(character) >= 0) return 0.63;
    return 0.50;
  }

  function countLines(text, widthTwips, fontSizePt) {
    if (!text) return 1;
    var emTwips = fontSizePt * 20;
    var available = Math.max(widthTwips * 0.97, emTwips);
    var lines = 1;
    var current = 0;
    var words = text.split(/\s+/).filter(function (word) { return word.length; });
    var spaceWidth = charWidthEm(" ") * emTwips;

    for (var i = 0; i < words.length; i++) {
      var wordWidth = 0;
      for (var c = 0; c < words[i].length; c++) wordWidth += charWidthEm(words[i].charAt(c)) * emTwips;
      if (current === 0) {
        current = wordWidth;
      } else if (current + spaceWidth + wordWidth <= available) {
        current += spaceWidth + wordWidth;
      } else {
        lines += 1;
        current = wordWidth;
      }
      while (current > available) {
        lines += 1;
        current -= available;
      }
    }
    return lines;
  }

  function spaceAfter(paragraph, defaults) {
    var pPr = kid(paragraph, "pPr");
    var spacing = pPr ? kid(pPr, "spacing") : null;
    if (!spacing) return defaults.space_after;
    return parseFloat(spacing.getAttributeNS(W, "after") || 0) || 0;
  }

  function paragraphHeight(paragraph, defaults, widthTwips) {
    var pPr = kid(paragraph, "pPr");
    var before = defaults.space_before;
    var after = defaults.space_after;
    var line = defaults.line;
    var rule = defaults.line_rule;
    var fontSize = defaults.font_size_pt;

    if (pPr) {
      var spacing = kid(pPr, "spacing");
      if (spacing) {
        var b = spacing.getAttributeNS(W, "before");
        var a = spacing.getAttributeNS(W, "after");
        var l = spacing.getAttributeNS(W, "line");
        var r = spacing.getAttributeNS(W, "lineRule");
        before = b === null ? before : (parseFloat(b) || 0);
        after = a === null ? after : (parseFloat(a) || 0);
        line = l === null ? line : (parseFloat(l) || 0);
        rule = r === null ? rule : r;
      }
    }

    var sizes = descendants(paragraph, "sz").map(function (node) {
      return intVal(node, 0) / 2;
    }).filter(function (size) { return size > 0; });
    if (sizes.length) fontSize = Math.max.apply(null, sizes);

    var single = fontSize * defaults.line_factor * 20;
    var lineHeight;
    if (rule === "exact") lineHeight = line;
    else if (rule === "atLeast") lineHeight = Math.max(line, single);
    else lineHeight = single * (line / 240);

    var lines = countLines(paragraphText(paragraph), widthTwips, fontSize);
    return before + lines * lineHeight + after;
  }

  function twipsAttr(node, name) {
    var value = node ? node.getAttributeNS(W, name) : null;
    var parsed = parseInt(value, 10);
    return isFinite(parsed) ? parsed : 0;
  }

  function gridWidths(table) {
    var grid = kid(table, "tblGrid");
    if (!grid) return [];
    return kids(grid, "gridCol").map(function (column) {
      return twipsAttr(column, "w");   // у gridCol ширина в w:w, а не в w:val
    });
  }

  function cellGridspan(cell) {
    var tcPr = kid(cell, "tcPr");
    var span = tcPr ? kid(tcPr, "gridSpan") : null;
    return Math.max(1, intVal(span, 1));
  }

  function vmergeState(cell) {
    var tcPr = kid(cell, "tcPr");
    var merge = tcPr ? kid(tcPr, "vMerge") : null;
    if (!merge) return null;
    return val(merge) === "restart" ? "restart" : "continue";
  }

  function cellTextWidth(cell, column, grid, span) {
    var width = 0;
    for (var i = column; i < column + span && i < grid.length; i++) width += grid[i];
    if (!width) width = 4000;
    return Math.max(600, width - 216);
  }

  function cellContentHeight(cell, defaults, textWidth) {
    var paragraphs = kids(cell, "p");
    if (!paragraphs.length) return 0;
    var total = 0;
    paragraphs.forEach(function (paragraph) {
      total += paragraphHeight(paragraph, defaults, textWidth);
    });
    return Math.max(0, total - spaceAfter(paragraphs[paragraphs.length - 1], defaults));
  }

  function findSpanEnd(rows, start, column) {
    var end = start;
    for (var index = start + 1; index < rows.length; index++) {
      var position = 0;
      var state = undefined;
      var cells = kids(rows[index], "tc");
      for (var c = 0; c < cells.length; c++) {
        var span = cellGridspan(cells[c]);
        if (position <= column && column < position + span) {
          state = vmergeState(cells[c]);
          break;
        }
        position += span;
      }
      if (state !== "continue") break;
      end = index;
    }
    return end;
  }

  function estimateRowsHeight(rows, defaults, grid) {
    var rowHeights = [];
    var spans = [];

    rows.forEach(function (row, index) {
      var trPr = kid(row, "trPr");
      var minimum = 0;
      if (trPr) {
        var height = kid(trPr, "trHeight");
        if (height) minimum = intVal(height, 0);
      }
      var best = 0;
      var column = 0;
      kids(row, "tc").forEach(function (cell) {
        var span = cellGridspan(cell);
        var state = vmergeState(cell);
        var width = cellTextWidth(cell, column, grid, span);
        var content = cellContentHeight(cell, defaults, width);
        if (state === "restart") {
          spans.push([index, findSpanEnd(rows, index, column), content]);
        } else if (state === null) {
          best = Math.max(best, content);
        }
        column += span;
      });
      rowHeights.push(Math.max(minimum, best));
    });

    spans.forEach(function (span) {
      var start = span[0], end = span[1], content = span[2];
      var total = 0;
      for (var i = start; i <= end; i++) total += rowHeights[i];
      if (content > total) rowHeights[end] += content - total;
    });
    return rowHeights;
  }

  /* ------------------------------------------------------------------ */
  /* Правка таблицы билета                                               */
  /* ------------------------------------------------------------------ */

  function pinTableGeometry(table) {
    var doc = table.ownerDocument;
    var tblPr = kid(table, "tblPr");
    if (!tblPr) {
      tblPr = mkEl(doc, "tblPr");
      table.insertBefore(tblPr, table.firstChild);
    }

    var grid = gridWidths(table);

    if (grid.length && !kid(tblPr, "tblW")) {
      var sum = grid.reduce(function (a, b) { return a + b; }, 0);
      insertOrdered(tblPr, mkEl(doc, "tblW", { w: sum, type: "dxa" }), TBLPR_INDEX);
    }

    var layout = kid(tblPr, "tblLayout");
    if (!layout) {
      layout = mkEl(doc, "tblLayout");
      insertOrdered(tblPr, layout, TBLPR_INDEX);
    }
    layout.setAttributeNS(W, "w:type", "fixed");

    if (SETTINGS.forceTableBorders && !kid(tblPr, "tblBorders")) {
      var borders = mkEl(doc, "tblBorders");
      ["top", "left", "bottom", "right", "insideH", "insideV"].forEach(function (edge) {
        borders.appendChild(mkEl(doc, edge, { val: "single", sz: 4, space: 0, color: "auto" }));
      });
      insertOrdered(tblPr, borders, TBLPR_INDEX);
    }

    kids(table, "tr").forEach(function (row) {
      var column = 0;
      kids(row, "tc").forEach(function (cell) {
        var span = cellGridspan(cell);
        var width = 0;
        for (var i = column; i < column + span && i < grid.length; i++) width += grid[i];
        column += span;
        if (!width) return;
        var tcPr = kid(cell, "tcPr");
        if (!tcPr) {
          tcPr = mkEl(doc, "tcPr");
          cell.insertBefore(tcPr, cell.firstChild);
        }
        var tcW = kid(tcPr, "tcW");
        if (!tcW) {
          tcW = mkEl(doc, "tcW");
          tcPr.insertBefore(tcW, tcPr.firstChild);
        }
        tcW.setAttributeNS(W, "w:w", String(width));
        tcW.setAttributeNS(W, "w:type", "dxa");
      });
    });
  }

  function setRowHeight(row, twips) {
    var doc = row.ownerDocument;
    var trPr = kid(row, "trPr");
    if (!trPr) {
      trPr = mkEl(doc, "trPr");
      row.insertBefore(trPr, row.firstChild);
    }
    var height = kid(trPr, "trHeight");
    if (!height) {
      height = mkEl(doc, "trHeight");
      trPr.appendChild(height);
    }
    height.setAttributeNS(W, "w:val", String(Math.round(twips)));
    height.setAttributeNS(W, "w:hRule", "atLeast");
  }

  function forbidRowSplit(row) {
    var doc = row.ownerDocument;
    var trPr = kid(row, "trPr");
    if (!trPr) {
      trPr = mkEl(doc, "trPr");
      row.insertBefore(trPr, row.firstChild);
    }
    if (kid(trPr, "cantSplit")) return;
    var cantSplit = mkEl(doc, "cantSplit");
    var height = kid(trPr, "trHeight");
    if (height) trPr.insertBefore(cantSplit, height);
    else trPr.appendChild(cantSplit);
  }

  function setCellValignTop(cell) {
    var doc = cell.ownerDocument;
    var tcPr = kid(cell, "tcPr");
    if (!tcPr) {
      tcPr = mkEl(doc, "tcPr");
      cell.insertBefore(tcPr, cell.firstChild);
    }
    var align = kid(tcPr, "vAlign");
    if (!align) {
      align = mkEl(doc, "vAlign");
      tcPr.appendChild(align);
    }
    align.setAttributeNS(W, "w:val", "top");
  }

  function setKeepNext(paragraph) {
    var doc = paragraph.ownerDocument;
    var pPr = kid(paragraph, "pPr");
    if (!pPr) {
      pPr = mkEl(doc, "pPr");
      paragraph.insertBefore(pPr, paragraph.firstChild);
    }
    if (kid(pPr, "keepNext")) return;
    insertOrdered(pPr, mkEl(doc, "keepNext"), PPR_INDEX);
  }

  function keepTicketTogether(table) {
    var rows = kids(table, "tr");
    rows.forEach(function (row, index) {
      forbidRowSplit(row);
      if (index === rows.length - 1) return;
      descendants(row, "p").forEach(setKeepNext);
    });
  }

  function stretchTicket(questionRow, targetHeight, headerHeight) {
    if (!questionRow) return;
    var wanted = targetHeight - headerHeight - SETTINGS.lastRowSafetyTwips - 20;
    if (wanted <= 0) return;

    var trPr = kid(questionRow, "trPr");
    var current = trPr ? intVal(kid(trPr, "trHeight"), 0) : 0;
    if (wanted > current) setRowHeight(questionRow, wanted);
    kids(questionRow, "tc").forEach(setCellValignTop);
  }

  function spacerParagraph(doc, heightTwips) {
    var paragraph = mkEl(doc, "p");
    var pPr = mkEl(doc, "pPr");
    pPr.appendChild(mkEl(doc, "spacing", {
      before: 0, after: 0, line: heightTwips, lineRule: "exact"
    }));
    var rPr = mkEl(doc, "rPr");
    rPr.appendChild(mkEl(doc, "sz", { val: 2 }));
    rPr.appendChild(mkEl(doc, "szCs", { val: 2 }));
    pPr.appendChild(rPr);
    paragraph.appendChild(pPr);
    return paragraph;
  }

  /* ------------------------------------------------------------------ */
  /* Случайная выборка                                                   */
  /* ------------------------------------------------------------------ */

  function mulberry32(seed) {
    var state = seed >>> 0;
    return function () {
      state = (state + 0x6D2B79F5) >>> 0;
      var t = state;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function sampleIndexes(rng, total, count) {
    var pool = [];
    for (var i = 0; i < total; i++) pool.push(i);
    for (var j = 0; j < count; j++) {
      var pick = j + Math.floor(rng() * (total - j));
      var tmp = pool[j]; pool[j] = pool[pick]; pool[pick] = tmp;
    }
    return pool.slice(0, count).sort(function (a, b) { return a - b; });
  }

  function combinations(n, k) {
    if (k > n) return 0;
    var result = 1;
    for (var i = 0; i < k; i++) result = result * (n - i) / (i + 1);
    return result;
  }

  function buildTickets(questions, ticketCount, perTicket, rng, unique) {
    if (ticketCount <= 0) throw new Error("Количество билетов должно быть больше нуля.");
    if (perTicket <= 0) throw new Error("Количество вопросов в билете должно быть больше нуля.");
    if (perTicket > questions.length) {
      throw new Error("В базе всего " + questions.length + " вопрос(ов), а в билете требуется " +
        perTicket + ". Уменьшите количество вопросов или дополните базу.");
    }

    var totalCombinations = combinations(questions.length, perTicket);
    var seen = Object.create(null);
    var tickets = [];
    var duplicated = 0;

    for (var number = 1; number <= ticketCount; number++) {
      var indexes = null;
      if (unique && totalCombinations > Object.keys(seen).length) {
        for (var attempt = 0; attempt < 300; attempt++) {
          var candidate = sampleIndexes(rng, questions.length, perTicket);
          var key = candidate.join(",");
          if (!seen[key]) { indexes = candidate; break; }
        }
      }
      if (!indexes) {
        indexes = sampleIndexes(rng, questions.length, perTicket);
        if (unique) duplicated += 1;
      }
      seen[indexes.join(",")] = true;
      tickets.push({
        number: number,
        questions: indexes.map(function (index) { return questions[index]; })
      });
    }
    return { tickets: tickets, duplicated: duplicated };
  }

  /* ------------------------------------------------------------------ */
  /* Чтение файлов                                                       */
  /* ------------------------------------------------------------------ */

  function parseXml(text, what) {
    var doc = new DOMParser().parseFromString(text, "application/xml");
    var error = doc.getElementsByTagName("parsererror")[0];
    if (error) throw new Error("Не удалось прочитать XML (" + what + ").");
    if (!doc.documentElement) throw new Error("Файл повреждён (" + what + ").");
    return doc;
  }

  function docxBody(doc) {
    var body = kids(doc.documentElement, "body")[0];
    if (!body) throw new Error("В файле нет тела документа Word.");
    return body;
  }

  function styleNames(stylesXml) {
    var map = {};
    if (!stylesXml) return map;
    try {
      var styles = new DOMParser().parseFromString(stylesXml, "application/xml");
      var nodes = styles.getElementsByTagNameNS(W, "style");
      for (var i = 0; i < nodes.length; i++) {
        var id = nodes[i].getAttributeNS(W, "styleId");
        var name = nodes[i].getElementsByTagNameNS(W, "name")[0];
        if (id) map[id] = name ? (name.getAttributeNS(W, "val") || "") : "";
      }
    } catch (e) { /* стили не критичны */ }
    return map;
  }

  function normalizeQuestionText(text) {
    var clean = String(text).replace(/\s+/g, " ").trim();
    clean = clean.replace(NUMBER_PREFIX, "");
    return clean;
  }

  /** Возвращает список вопросов из .docx (каждый непустой абзац — вопрос). */
  function extractQuestions(documentXml, stylesXml) {
    var doc = parseXml(documentXml, "вопросы");
    var body = docxBody(doc);
    var names = styleNames(stylesXml);
    var questions = [];
    var seen = Object.create(null);

    kids(body, "p").forEach(function (paragraph) {
      var pPr = kid(paragraph, "pPr");
      var styleNode = pPr ? kid(pPr, "pStyle") : null;
      var styleName = styleNode ? (names[val(styleNode)] || "").toLowerCase() : "";
      for (var i = 0; i < HEADING_HINTS.length; i++) {
        if (styleName.indexOf(HEADING_HINTS[i]) >= 0) return;
      }
      var text = normalizeQuestionText(paragraphText(paragraph));
      if (!text) return;
      var key = text.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      questions.push(text);
    });
    return questions;
  }

  function extractPlaceholders(documentXml) {
    var doc = parseXml(documentXml, "шаблон");
    var found = [];
    var seen = Object.create(null);
    descendants(docxBody(doc), "p").forEach(function (paragraph) {
      var matches = paragraphText(paragraph).match(PLACEHOLDER_ALL) || [];
      matches.forEach(function (token) {
        if (!seen[token]) { seen[token] = true; found.push(token); }
      });
    });
    return found;
  }

  function pageGeometry(body) {
    var sectPr = kid(body, "sectPr");
    var result = { usableHeight: 13958, width: 11906 };
    if (!sectPr) return result;
    var pgSz = kid(sectPr, "pgSz");
    var pgMar = kid(sectPr, "pgMar");
    var height = twipsAttr(pgSz, "h") || 16838;
    var top = pgMar ? twipsAttr(pgMar, "top") : 1440;
    var bottom = pgMar ? twipsAttr(pgMar, "bottom") : 1440;
    result.usableHeight = height - top - bottom;
    result.width = twipsAttr(pgSz, "w") || 11906;
    return result;
  }

  /* ------------------------------------------------------------------ */
  /* Сборка итогового файла                                              */
  /* ------------------------------------------------------------------ */

  async function readPackage(bytes) {
    if (typeof JSZip === "undefined" && global.JSZip) global.JSZip = global.JSZip;
    var zip = await global.JSZip.loadAsync(bytes);
    var documentFile = zip.file("word/document.xml");
    if (!documentFile) throw new Error("Это не документ Word: нет word/document.xml.");
    var stylesFile = zip.file("word/styles.xml");
    var themeFile = zip.file("word/theme/theme1.xml");
    return {
      zip: zip,
      documentXml: await documentFile.async("string"),
      stylesXml: stylesFile ? await stylesFile.async("string") : null,
      themeXml: themeFile ? await themeFile.async("string") : null
    };
  }

  /** Разбирает шаблон: заполнители, наличие <BILET_TASK>, геометрия страницы. */
  async function analyzeTemplate(bytes) {
    var pack = await readPackage(bytes);
    var doc = parseXml(pack.documentXml, "шаблон");
    var body = docxBody(doc);
    var table = descendants(body, "tbl")[0];
    if (!table) throw new Error("В шаблоне не найдена таблица билета.");
    var placeholders = extractPlaceholders(pack.documentXml);
    var geometry = pageGeometry(body);
    var defaults = readStyleDefaults(pack.stylesXml, pack.themeXml);
    return {
      placeholders: placeholders,
      hasTasks: placeholders.indexOf(SETTINGS.tasksPlaceholder) >= 0,
      hasNumber: placeholders.indexOf("<NUM>") >= 0 || placeholders.indexOf("<num>") >= 0,
      usableHeight: geometry.usableHeight,
      headerHeight: estimateHeaderHeight(table, defaults)
    };
  }

  async function analyzeQuestions(bytes) {
    var pack = await readPackage(bytes);
    return extractQuestions(pack.documentXml, pack.stylesXml);
  }

  function estimateHeaderHeight(table, defaults) {
    var rows = kids(table, "tr");
    if (rows.length < 2) return 0;
    var grid = gridWidths(table);
    var heights = estimateRowsHeight(rows.slice(0, rows.length - 1), defaults, grid);
    var total = heights.reduce(function (a, b) { return a + b; }, 0);
    return total * SETTINGS.heightSafetyFactor + 20 * heights.length;
  }

  /**
   * Генерирует .docx.
   * options: { templateBytes, questionsBytes, values, ticketCount,
   *            questionsPerTicket, maxTicketsPerPage, uniqueTickets, seed }
   */
  async function generate(options) {
    var templateCount = await readPackage(options.templateBytes);
    var questionsPack = await readPackage(options.questionsBytes);

    var questions = extractQuestions(questionsPack.documentXml, questionsPack.stylesXml);
    if (!questions.length) {
      throw new Error("В файле с вопросами не найдено ни одного вопроса: каждый вопрос " +
        "должен быть отдельным абзацем.");
    }

    var ticketCount = parseInt(options.ticketCount, 10);
    var perTicket = parseInt(options.questionsPerTicket, 10);
    var perPage = Math.max(1, parseInt(options.maxTicketsPerPage, 10) || 2);
    var unique = options.uniqueTickets !== false;
    var seed = (options.seed === null || options.seed === undefined || options.seed === "")
      ? null : parseInt(options.seed, 10);
    var rng = seed === null || !isFinite(seed) ? Math.random : mulberry32(seed);

    var built;
    if (Array.isArray(options.questionSets)) {
      // готовая разбивка (используется тестами и предпросмотром)
      built = {
        tickets: options.questionSets.map(function (set, index) {
          return {
            number: index + 1,
            questions: set.map(function (questionIndex) { return questions[questionIndex]; })
          };
        }),
        duplicated: 0
      };
    } else {
      built = buildTickets(questions, ticketCount, perTicket, rng, unique);
    }

    var doc = parseXml(templateCount.documentXml, "шаблон");
    var body = docxBody(doc);
    var templateTable = descendants(body, "tbl")[0];
    if (!templateTable) throw new Error("В шаблоне не найдена таблица билета.");
    if (paragraphTextOfTable(templateTable).indexOf(SETTINGS.tasksPlaceholder) < 0) {
      throw new Error("В шаблоне не найден заполнитель " + SETTINGS.tasksPlaceholder +
        " — именно в это место подставляются вопросы.");
    }

    var defaults = readStyleDefaults(templateCount.stylesXml, templateCount.themeXml);
    var geometry = pageGeometry(body);
    var gap = perPage > 1 ? SETTINGS.ticketGapTwips : 0;
    var slotHeight = (geometry.usableHeight - gap * (perPage - 1)) / perPage -
      SETTINGS.pageSlotPadTwips;
    if (slotHeight <= 500) throw new Error("Слишком маленькая страница в шаблоне — билет не помещается.");

    var grid = gridWidths(templateTable);
    var sectPr = kid(body, "sectPr");
    elementChildren(body).forEach(function (child) {
      if (child !== sectPr) body.removeChild(child);
    });

    var headerHeight = null;
    var naturalHeights = [];
    var elements = [];
    var usedQuestions = Object.create(null);
    var unknown = Object.create(null);

    built.tickets.forEach(function (ticket, index) {
      var values = normalizeValues(options.values || {});
      values[SETTINGS.numberPlaceholder] = String(ticket.number);

      var table = templateTable.cloneNode(true);
      stripVolatileIds(table);
      var questionRow = insertQuestions(table, ticket.questions, values);
      descendants(table, "p").forEach(function (paragraph) {
        replacePlaceholders(paragraph, values);
      });

      descendants(table, "p").forEach(function (paragraph) {
        var matches = paragraphText(paragraph).match(PLACEHOLDER_ALL) || [];
        matches.forEach(function (token) { unknown[token] = true; });
      });

      var rows = kids(table, "tr");
      if (headerHeight === null) {
        var questionRowIndex = rows.indexOf(questionRow);
        var headerRows = rows.slice(0, questionRowIndex >= 0 ? questionRowIndex : rows.length - 1);
        var headerSum = estimateRowsHeight(headerRows, defaults, grid)
          .reduce(function (a, b) { return a + b; }, 0);
        headerHeight = headerSum * SETTINGS.heightSafetyFactor + 20 * headerRows.length;
      }

      naturalHeights.push(estimateRowsHeight(rows, defaults, grid)
        .reduce(function (a, b) { return a + b; }, 0));

      pinTableGeometry(table);
      stretchTicket(questionRow, slotHeight, headerHeight);
      keepTicketTogether(table);

      if (index) elements.push(spacerParagraph(doc, gap));
      elements.push(table);

      ticket.questions.forEach(function (question) { usedQuestions[question] = true; });
    });

    elements.push(spacerParagraph(doc, 2));

    elements.forEach(function (element) {
      if (sectPr) body.insertBefore(element, sectPr);
      else body.appendChild(element);
    });

    var serializer = new XMLSerializer();
    // некоторые парсеры сами добавляют объявление XML — оставляем ровно одно
    var xml = serializer.serializeToString(doc)
      .replace(/^\uFEFF/, "")
      .replace(/^\s*<\?xml[^>]*\?>\s*/, "");
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
    templateCount.zip.file("word/document.xml", xml);
    var bytes = await templateCount.zip.generateAsync({
      type: "uint8array",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    var maxNatural = naturalHeights.length ? Math.max.apply(null, naturalHeights) : 0;
    return {
      bytes: bytes,
      fileName: "Билеты.docx",
      questions: questions,
      tickets: built.tickets,
      report: {
        ticketCount: built.tickets.length,
        perTicket: perTicket,
        questionsInBase: questions.length,
        usedQuestions: Object.keys(usedQuestions).length,
        perPage: perPage,
        usableHeight: geometry.usableHeight,
        slotHeight: slotHeight,
        headerHeight: headerHeight || 0,
        naturalHeight: maxNatural,
        duplicated: built.duplicated,
        unknownPlaceholders: Object.keys(unknown)
      }
    };
  }

  function paragraphTextOfTable(table) {
    return descendants(table, "p").map(paragraphText).join("\n");
  }

  var api = {
    settings: SETTINGS,
    analyzeTemplate: analyzeTemplate,
    analyzeQuestions: analyzeQuestions,
    extractQuestions: extractQuestions,
    extractPlaceholders: extractPlaceholders,
    generate: generate
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  global.BiletEngine = api;
})(typeof window !== "undefined" ? window : globalThis);
