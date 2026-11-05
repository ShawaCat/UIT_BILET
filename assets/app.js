/*
 * app.js — логика страницы: чтение файлов, поля заполнителей, запуск генерации и скачивание.
 */
(function () {
  "use strict";

  var engine = window.BiletEngine;

  var DEFAULT_FILES = { template: "Макет.docx", questions: "Вопросы.docx" };
  var STORAGE_KEY = "bilet-generator-settings";

  /* Человеческие подписи для известных заполнителей */
  var FIELD_LABELS = {
    DISCIP: "Дисциплина",
    DATE: "Дата",
    LECTOR_NAME: "Лектор",
    KAFEDRA: "Кафедра",
    FACULTET: "Факультет",
    GROUP: "Группа",
    UNIVERS: "Учебное заведение",
    ZAVKAF: "Зав. кафедрой"
  };

  var MONTHS_RU = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
    "августа", "сентября", "октября", "ноября", "декабря"];

  var DEFAULT_VALUES = {
    DISCIP: "ТАУ",
    DATE: "",
    LECTOR_NAME: "Сидорова Е.Ю."
  };

  var state = {
    templateBytes: null,
    questionsBytes: null,
    templateName: "",
    questionsName: "",
    questionCount: 0,
    hasTasks: false,
    placeholders: [],
    lastBlobUrl: null,
    lastFileName: "",
    busy: false
  };

  var el = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    [
      "template-input", "questions-input", "template-state", "questions-state",
      "repo-files", "repo-hint", "fields", "ini-input", "ini-apply",
      "count", "per-ticket", "per-page", "seed", "unique", "params-hint",
      "generate", "status", "report", "preview", "preview-wrap"
    ].forEach(function (id) { el[id] = document.getElementById(id); });

    if (!engine) {
      showMessage("error", "Не загрузился движок генерации (assets/bilet-engine.js). Обновите страницу.");
      return;
    }

    DEFAULT_VALUES.DATE = todayRu();
    restoreSettings();

    el["template-input"].addEventListener("change", onTemplateChange);
    el["questions-input"].addEventListener("change", onQuestionsChange);
    el["repo-files"].addEventListener("click", onRepoFiles);
    el["ini-apply"].addEventListener("click", onIniApply);
    el["generate"].addEventListener("click", onGenerate);
    ["count", "per-ticket", "per-page", "seed", "unique"].forEach(function (id) {
      el[id].addEventListener("change", function () { saveSettings(); updateParamsHint(); });
    });
    el.fields.addEventListener("input", saveSettings);
    updateParamsHint();
    updateReadyState();
  }

  /* ------------------------------------------------------------------ */
  /* Файлы                                                              */
  /* ------------------------------------------------------------------ */

  function fileToBytes(file) {
    if (file.arrayBuffer) {
      return file.arrayBuffer().then(function (buffer) { return new Uint8Array(buffer); });
    }
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(new Uint8Array(reader.result)); };
      reader.onerror = function () { reject(new Error("Не удалось прочитать файл.")); };
      reader.readAsArrayBuffer(file);
    });
  }

  function setFileState(which, text, kind) {
    var node = el[which + "-state"];
    if (!node) return;
    node.textContent = text;
    node.setAttribute("data-state", kind || "empty");
  }

  async function onTemplateChange(event) {
    var file = event.target.files && event.target.files[0];
    if (!file) return;
    setFileState("template", "читаю «" + file.name + "»…", "empty");
    try {
      await applyTemplate(await fileToBytes(file), file.name);
    } catch (e) {
      state.templateBytes = null;
      setFileState("template", "не удалось прочитать файл: " + (e.message || e), "error");
      showMessage("error", "Шаблон не загружен: " + (e.message || e));
    }
    updateReadyState();
  }

  async function applyTemplate(bytes, fileName) {
    try {
      var info = await engine.analyzeTemplate(bytes);
      state.templateBytes = bytes;
      state.templateName = fileName;
      state.placeholders = info.placeholders;
      state.hasTasks = info.hasTasks;
      setFileState("template",
        "«" + fileName + "» · заполнители: " + (info.placeholders.join(" ") || "нет"),
        info.hasTasks ? "ok" : "error");
      if (!info.hasTasks) {
        showMessage("error", "В шаблоне нет заполнителя " + engine.settings.tasksPlaceholder +
          " — непонятно, куда вставлять вопросы.");
      }
      renderFields();
      saveSettings();
    } catch (e) {
      state.templateBytes = null;
      setFileState("template", "ошибка: " + (e.message || e), "error");
      throw e;
    }
  }

  async function onQuestionsChange(event) {
    var file = event.target.files && event.target.files[0];
    if (!file) return;
    setFileState("questions", "читаю «" + file.name + "»…", "empty");
    try {
      await applyQuestions(await fileToBytes(file), file.name);
    } catch (e) {
      state.questionsBytes = null;
      state.questionCount = 0;
      setFileState("questions", "ошибка: " + (e.message || e), "error");
      showMessage("error", "База вопросов не загружена: " + (e.message || e));
    }
    updateReadyState();
  }

  async function applyQuestions(bytes, fileName) {
    var questions = await engine.analyzeQuestions(bytes);
    state.questionsBytes = bytes;
    state.questionsName = fileName;
    state.questionCount = questions.length;
    setFileState("questions", "«" + fileName + "» · вопросов: " + questions.length,
      questions.length ? "ok" : "error");
    if (!questions.length) {
      showMessage("error", "В файле вопросов не найдено ни одного вопроса.");
    }
    updateParamsHint();
    saveSettings();
  }

  async function onRepoFiles() {
    if (state.busy) return;
    el["repo-files"].disabled = true;
    showMessage("info", "Загружаю файлы с сайта…");
    try {
      var templateBytes = await fetchBytes(DEFAULT_FILES.template);
      var questionsBytes = await fetchBytes(DEFAULT_FILES.questions);
      await applyTemplate(templateBytes, DEFAULT_FILES.template);
      await applyQuestions(questionsBytes, DEFAULT_FILES.questions);
      showMessage("ok", "Файлы из репозитория загружены: " + DEFAULT_FILES.template + ", " +
        DEFAULT_FILES.questions + ".");
      updateParamsHint();
    } catch (e) {
      showMessage("error", "Не получилось загрузить файлы с сайта (" + (e.message || e) + "). " +
        "Скорее всего, страница открыта локально (file://) — выберите файлы вручную.");
    } finally {
      el["repo-files"].disabled = false;
      updateReadyState();
    }
  }

  function fetchBytes(name) {
    return fetch(encodeURIComponent(name)).then(function (response) {
      if (!response.ok) throw new Error("HTTP " + response.status);
      return response.arrayBuffer();
    }).then(function (buffer) { return new Uint8Array(buffer); });
  }

  /* ------------------------------------------------------------------ */
  /* Поля заполнителей                                                  */
  /* ------------------------------------------------------------------ */

  function answerKey(token) {
    return token.replace(/[<>]/g, "").toUpperCase();
  }

  function isAutoPlaceholder(token) {
    var key = answerKey(token);
    return key === engine.settings.numberPlaceholder.toUpperCase() ||
      key === answerKey(engine.settings.tasksPlaceholder);
  }

  function currentValues() {
    var values = {};
    el.fields.querySelectorAll("input[data-key]").forEach(function (input) {
      values[input.getAttribute("data-key")] = input.value;
    });
    return values;
  }

  function renderFields() {
    var previous = currentValues();
    var tokens = (state.placeholders || []).filter(function (token) { return !isAutoPlaceholder(token); });
    el.fields.innerHTML = "";

    if (!tokens.length) {
      var hint = document.createElement("p");
      hint.className = "hint";
      hint.textContent = state.templateBytes
        ? "В шаблоне нет заполнителей — заполнять нечего."
        : "Поля появляются после выбора шаблона — по заполнителям в нём.";
      el.fields.appendChild(hint);
      return;
    }

    tokens.forEach(function (token) {
      var key = answerKey(token);
      var label = document.createElement("label");
      label.className = "field";

      var title = document.createElement("span");
      title.className = "field-name";
      title.textContent = FIELD_LABELS[key] || key;
      var code = document.createElement("code");
      code.textContent = " " + token;
      title.appendChild(code);

      var input = document.createElement("input");
      input.type = "text";
      input.setAttribute("data-key", key);
      input.value = previous[key] !== undefined && previous[key] !== ""
        ? previous[key]
        : (DEFAULT_VALUES[key] || "");

      label.appendChild(title);
      label.appendChild(input);
      el.fields.appendChild(label);
    });
  }

  function onIniApply() {
    var text = el["ini-input"].value || "";
    var values = {};
    text.split(/\r?\n/).forEach(function (line) {
      var trimmed = line.trim();
      if (!trimmed || trimmed.charAt(0) === "#" || trimmed.charAt(0) === ";") return;
      var position = trimmed.indexOf("=");
      if (position < 0) return;
      var key = trimmed.slice(0, position).trim().replace(/[<>]/g, "").toUpperCase();
      var value = trimmed.slice(position + 1).trim().replace(/^["']|["']$/g, "");
      if (key) values[key] = value;
    });

    var applied = 0;
    el.fields.querySelectorAll("input[data-key]").forEach(function (input) {
      var key = input.getAttribute("data-key");
      if (Object.prototype.hasOwnProperty.call(values, key)) {
        input.value = values[key];
        applied += 1;
      }
    });
    showMessage(applied ? "ok" : "warn", applied
      ? "Заполнено полей: " + applied + "."
      : "В тексте не нашлось ни одного знакомого заполнителя.");
    saveSettings();
  }

  /* ------------------------------------------------------------------ */
  /* Параметры                                                          */
  /* ------------------------------------------------------------------ */

  function readParams() {
    return {
      ticketCount: parseInt(el.count.value, 10),
      questionsPerTicket: parseInt(el["per-ticket"].value, 10),
      perPage: parseInt(el["per-page"].value, 10),
      seed: el.seed.value === "" ? null : parseInt(el.seed.value, 10),
      unique: el.unique.checked
    };
  }

  function updateParamsHint() {
    var params = readParams();
    if (!state.questionCount) {
      el["params-hint"].textContent = "";
      return;
    }
    if (isFinite(params.questionsPerTicket) && params.questionsPerTicket > state.questionCount) {
      el["params-hint"].textContent = "В базе " + state.questionCount +
        " вопрос(ов) — уменьшите M, иначе сгенерировать не получится.";
    } else if (isFinite(params.ticketCount) && isFinite(params.questionsPerTicket)) {
      el["params-hint"].textContent = "Будет создано " + params.ticketCount + " билет(ов) по " +
        params.questionsPerTicket + " вопрос(ов) из " + state.questionCount + ".";
    } else {
      el["params-hint"].textContent = "";
    }
  }

  function updateReadyState() {
    var params = readParams();
    var ready = Boolean(state.templateBytes && state.questionsBytes) && !state.busy &&
      isFinite(params.ticketCount) && params.ticketCount > 0 &&
      isFinite(params.questionsPerTicket) && params.questionsPerTicket > 0;
    el.generate.disabled = !ready;
  }

  /* ------------------------------------------------------------------ */
  /* Генерация                                                          */
  /* ------------------------------------------------------------------ */

  async function onGenerate() {
    if (state.busy) return;
    var params = readParams();

    if (!state.templateBytes || !state.questionsBytes) {
      showMessage("error", "Сначала выберите шаблон и базу вопросов.");
      return;
    }
    if (!isFinite(params.ticketCount) || params.ticketCount <= 0) {
      showMessage("error", "Количество билетов должно быть больше нуля.");
      return;
    }
    if (!isFinite(params.questionsPerTicket) || params.questionsPerTicket <= 0) {
      showMessage("error", "Количество вопросов в билете должно быть больше нуля.");
      return;
    }

    state.busy = true;
    el.generate.disabled = true;
    el.generate.textContent = "Генерирую…";
    clearMessages();

    try {
      var result = await engine.generate({
        templateBytes: state.templateBytes,
        questionsBytes: state.questionsBytes,
        values: currentValues(),
        ticketCount: params.ticketCount,
        questionsPerTicket: params.questionsPerTicket,
        maxTicketsPerPage: params.perPage,
        uniqueTickets: params.unique,
        seed: params.seed
      });

      var fileName = "Билеты_" + params.ticketCount + "x" + params.questionsPerTicket + ".docx";
      downloadBytes(result.bytes, fileName);

      showMessage("ok", "Готово: " + params.ticketCount + " билет(ов) по " +
        params.questionsPerTicket + " вопрос(ов). Файл скачивается: " + fileName);
      renderReport(result.report);
      renderPreview(result.tickets);
      renderDownloadAgain(fileName);
      saveSettings();
    } catch (e) {
      showMessage("error", e && e.message ? e.message : String(e));
    } finally {
      state.busy = false;
      el.generate.textContent = "Сгенерировать и скачать .docx";
      updateReadyState();
    }
  }

  function downloadBytes(bytes, fileName) {
    var blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
    if (state.lastBlobUrl) URL.revokeObjectURL(state.lastBlobUrl);
    state.lastBlobUrl = URL.createObjectURL(blob);
    state.lastFileName = fileName;

    var link = document.createElement("a");
    link.href = state.lastBlobUrl;
    link.download = fileName;
    link.rel = "noopener";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  function renderReport(report) {
    el.report.hidden = false;
    el.report.innerHTML = "";

    var list = document.createElement("dl");
    var rows = [
      ["Билетов", report.ticketCount],
      ["Вопросов в билете", report.perTicket],
      ["Вопросов в базе", report.questionsInBase + " (использовано " + report.usedQuestions + ")"],
      ["Билетов на странице", report.perPage === 1 ? "1" : "до 2 (если влезают)"],
      ["Место под билет", Math.round(report.slotHeight) + " twips, шапка ≈ " + Math.round(report.headerHeight)],
      ["Оценка высоты билета", Math.round(report.naturalHeight) + " twips"]
    ];
    rows.forEach(function (row) {
      var dt = document.createElement("dt");
      dt.textContent = row[0];
      var dd = document.createElement("dd");
      dd.textContent = String(row[1]);
      list.appendChild(dt);
      list.appendChild(dd);
    });
    el.report.appendChild(list);

    if (report.perPage > 1 && report.naturalHeight > report.slotHeight) {
      showMessage("warn", "Билет выше отведённого места — на страницах будет по одному билету. " +
        "Уменьшите количество вопросов в билете.");
    } else if (report.perPage > 1) {
      showMessage("info", "На каждой странице помещается по два билета.");
    }
    if (report.duplicated) {
      showMessage("warn", "Вариантов наборов меньше, чем билетов: " + report.duplicated +
        " билет(ов) повторяют состав вопросов других.");
    }
    if (report.unknownPlaceholders && report.unknownPlaceholders.length) {
      showMessage("warn", "Не заполнены заполнители: " + report.unknownPlaceholders.join(" ") +
        " — добавьте для них поля или значения.");
    }
  }

  function renderPreview(tickets) {
    el["preview-wrap"].hidden = false;
    el.preview.innerHTML = "";
    tickets.slice(0, 5).forEach(function (ticket) {
      var block = document.createElement("div");
      block.className = "ticket";
      var title = document.createElement("h3");
      title.textContent = "Билет №" + ticket.number;
      var list = document.createElement("ol");
      ticket.questions.forEach(function (question) {
        var item = document.createElement("li");
        item.textContent = question;
        list.appendChild(item);
      });
      block.appendChild(title);
      block.appendChild(list);
      el.preview.appendChild(block);
    });
    if (tickets.length > 5) {
      var rest = document.createElement("p");
      rest.className = "hint";
      rest.textContent = "…и ещё " + (tickets.length - 5) + " билет(ов) в файле.";
      el.preview.appendChild(rest);
    }
  }

  function renderDownloadAgain(fileName) {
    var existing = el.report.querySelector(".download-again");
    if (existing) existing.remove();

    var wrapper = document.createElement("p");
    wrapper.className = "download-again";
    wrapper.innerHTML = "Если браузер не сохранил файл автоматически — ";
    var link = document.createElement("a");
    link.href = state.lastBlobUrl;
    link.download = fileName;
    link.textContent = "скачать ещё раз";
    wrapper.appendChild(link);
    el.report.appendChild(wrapper);
  }

  /* ------------------------------------------------------------------ */
  /* Сообщения и сохранение настроек                                    */
  /* ------------------------------------------------------------------ */

  function showMessage(kind, text) {
    var node = document.createElement("div");
    node.className = "msg " + kind;
    node.textContent = text;
    el.status.appendChild(node);
  }

  function clearMessages() {
    el.status.innerHTML = "";
    el.report.hidden = true;
    el.report.innerHTML = "";
    el["preview-wrap"].hidden = true;
    el.preview.innerHTML = "";
  }

  function saveSettings() {
    try {
      var values = currentValues();
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        values: values,
        count: el.count.value,
        perTicket: el["per-ticket"].value,
        perPage: el["per-page"].value,
        seed: el.seed.value,
        unique: el.unique.checked
      }));
    } catch (e) { /* приватный режим — молча пропускаем */ }
  }

  function restoreSettings() {
    var saved = null;
    try {
      saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    } catch (e) { saved = null; }
    if (!saved) return;
    if (saved.values) {
      Object.keys(saved.values).forEach(function (key) {
        DEFAULT_VALUES[key] = saved.values[key];
      });
    }
    if (saved.count) el.count.value = saved.count;
    if (saved.perTicket) el["per-ticket"].value = saved.perTicket;
    if (saved.perPage) el["per-page"].value = saved.perPage;
    if (saved.seed) el.seed.value = saved.seed;
    if (typeof saved.unique === "boolean") el.unique.checked = saved.unique;
  }

  function todayRu() {
    var now = new Date();
    var day = now.getDate() < 10 ? "0" + now.getDate() : String(now.getDate());
    return day + " " + MONTHS_RU[now.getMonth()] + " " + now.getFullYear();
  }
})();
