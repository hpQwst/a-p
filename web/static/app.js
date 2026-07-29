(function () {
  var deckSummary = null;

  /* ---------- tema claro/escuro ---------- */

  function toggleTheme() {
    var root = document.documentElement;
    var next = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
    root.setAttribute("data-theme", next);
    try {
      localStorage.setItem("qwst-theme", next);
    } catch (err) {
      /* armazenamento indisponivel: tema vale so para esta pagina */
    }
  }

  /* ---------- dropzones de arquivo ---------- */

  function updateFileCard(input) {
    var card = input.closest("[data-dropzone]");
    if (!card) {
      return;
    }
    var label = card.querySelector("[data-file-name]");
    var files = input.files;
    var count = files ? files.length : 0;
    card.classList.toggle("has-file", count > 0);
    if (label) {
      if (count === 0) {
        label.textContent = "";
      } else if (count === 1) {
        label.textContent = files[0].name + " (" + formatBytes(files[0].size) + ")";
      } else {
        var total = 0;
        for (var i = 0; i < count; i += 1) {
          total += files[i].size;
        }
        label.textContent = count + " arquivos (" + formatBytes(total) + ")";
      }
    }
  }

  function formatBytes(size) {
    if (!size && size !== 0) {
      return "";
    }
    var units = ["B", "KB", "MB", "GB"];
    var value = size;
    var unit = 0;
    while (value >= 1024 && unit < units.length - 1) {
      value = value / 1024;
      unit += 1;
    }
    return (unit === 0 ? value : value.toFixed(1)) + " " + units[unit];
  }

  function updateCombinedUploadSize() {
    var form = document.querySelector("[data-combined-upload-warning-mb]");
    var summary = document.querySelector("[data-upload-size-summary]");
    if (!form || !summary) {
      return;
    }
    var total = 0;
    form.querySelectorAll("input[type='file']").forEach(function (input) {
      Array.prototype.forEach.call(input.files || [], function (file) {
        total += file.size || 0;
      });
    });
    var thresholdMb = parseInt(form.getAttribute("data-combined-upload-warning-mb") || "250", 10);
    var threshold = thresholdMb * 1024 * 1024;
    summary.classList.toggle("warn", total > threshold);
    if (!total) {
      summary.textContent = "Selecione os arquivos para ver o tamanho combinado.";
    } else if (total > threshold) {
      summary.textContent = "Arquivos: " + formatBytes(total) + ". Caso grande: o processamento pode levar alguns minutos. O tamanho é só uma estimativa; a ferramenta também verifica slides e objetos.";
    } else {
      summary.textContent = "Tamanho combinado: " + formatBytes(total) + ".";
    }
  }

  function bindDropzones() {
    document.querySelectorAll("[data-dropzone]").forEach(function (card) {
      var input = card.querySelector("input[type='file']");
      if (!input) {
        return;
      }
      ["dragenter", "dragover"].forEach(function (name) {
        card.addEventListener(name, function (event) {
          event.preventDefault();
          card.classList.add("dragover");
        });
      });
      ["dragleave", "drop"].forEach(function (name) {
        card.addEventListener(name, function (event) {
          event.preventDefault();
          card.classList.remove("dragover");
        });
      });
      card.addEventListener("drop", function (event) {
        if (event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files.length) {
          input.files = event.dataTransfer.files;
          input.dispatchEvent(new Event("change", { bubbles: true }));
        }
      });
    });
  }

  function showProgress(message) {
    var overlay = document.getElementById("progress-overlay");
    var label = document.getElementById("progress-message");
    var steps = document.getElementById("progress-steps");
    var realProgress = document.getElementById("real-progress");
    if (!overlay) {
      return;
    }
    if (label) {
      label.textContent = message || "Processando...";
    }
    if (steps) {
      renderProgressSteps(steps, progressStepsForMessage(message || ""));
    }
    if (realProgress) {
      realProgress.hidden = true;
    }
    overlay.hidden = false;
  }

  function setRealProgress(progress, countSuffix) {
    var container = document.getElementById("real-progress");
    var bar = document.getElementById("progress-bar");
    var count = document.getElementById("progress-count");
    if (!container || !bar || !progress) {
      return;
    }
    var total = parseInt(progress.total || 0, 10);
    var completed = parseInt(progress.completed || 0, 10);
    var percent = Math.max(0, Math.min(100, parseInt(progress.percent || 0, 10)));
    container.hidden = total <= 0;
    bar.value = percent;
    bar.textContent = percent + "%";
    if (count) {
      count.textContent = completed + " de " + total + " " + (countSuffix || "objetos") + " · " + percent + "%";
    }
  }

  function progressStepsForMessage(message) {
    var text = (message || "").toLowerCase();
    if (text.indexOf("preview") !== -1 || text.indexOf("analisando") !== -1) {
      return ["Lendo PPTX e ZIP", "Extraindo contexto dos XLSX", "Mapeando targets", "Executando IA seletiva", "Montando preview"];
    }
    if (text.indexOf("ppt") !== -1 || text.indexOf("download") !== -1 || text.indexOf("gerando") !== -1) {
      return ["Validando revisoes", "Aplicando matrizes", "Atualizando graficos/tabelas", "Gerando arquivo"];
    }
    if (text.indexOf("ia") !== -1) {
      return ["Preparando contexto", "Enviando somente pendencias", "Validando resposta", "Atualizando preview"];
    }
    return ["Preparando dados", "Processando", "Atualizando tela"];
  }

  function renderProgressSteps(container, labels) {
    container.innerHTML = "";
    labels.forEach(function (label, index) {
      var item = document.createElement("li");
      item.textContent = label;
      if (index === 0) {
        item.className = "active";
      }
      container.appendChild(item);
    });
    var activeIndex = 0;
    window.clearInterval(container._progressTimer);
    container._progressTimer = window.setInterval(function () {
      var items = container.querySelectorAll("li");
      if (!items.length) {
        return;
      }
      items[activeIndex].classList.remove("active");
      items[activeIndex].classList.add("done");
      activeIndex = Math.min(activeIndex + 1, items.length - 1);
      items[activeIndex].classList.add("active");
      if (activeIndex === items.length - 1) {
        window.clearInterval(container._progressTimer);
      }
    }, 1800);
  }

  function activeSquad() {
    var projectSelect = document.querySelector("[data-project-select]");
    var squadSelect = document.querySelector("[data-squad-select]");
    if (projectSelect && projectSelect.value) {
      var selectedProject = projectSelect.options[projectSelect.selectedIndex];
      return selectedProject ? selectedProject.getAttribute("data-squad") || "" : "";
    }
    return squadSelect ? squadSelect.value : "";
  }

  function syncMappingTemplateOptions() {
    var mappingSelect = document.querySelector("[data-mapping-template-select]");
    if (!mappingSelect) {
      return;
    }
    var squad = activeSquad();
    Array.prototype.forEach.call(mappingSelect.options, function (option) {
      var optionSquad = option.getAttribute("data-squad") || "";
      var visible = !option.value || !optionSquad || optionSquad === squad;
      option.hidden = !visible;
      option.disabled = !visible;
    });
    if (mappingSelect.selectedOptions.length && mappingSelect.selectedOptions[0].disabled) {
      mappingSelect.value = "";
    }
  }

  function previewFilterState() {
    var search = document.querySelector("[data-preview-search]");
    var activeChip = document.querySelector("[data-status-filter].active");
    return {
      text: search ? search.value.trim().toLowerCase() : "",
      status: activeChip ? activeChip.getAttribute("data-status-filter") || "all" : "all",
    };
  }

  function applyPreviewFilters() {
    var cards = document.querySelectorAll("[data-target-card]");
    if (!cards.length) {
      return;
    }
    var state = previewFilterState();
    cards.forEach(function (card) {
      var status = card.getAttribute("data-status") || "";
      var search = card.getAttribute("data-search") || "";
      var statusMatch = state.status === "all" || status === state.status;
      var textMatch = !state.text || search.indexOf(state.text) !== -1;
      card.hidden = !(statusMatch && textMatch);
    });

    document.querySelectorAll("[data-slide-section]").forEach(function (section) {
      var visibleCards = section.querySelectorAll("[data-target-card]:not([hidden])");
      var hasVisibleCards = visibleCards.length > 0;
      section.hidden = !hasVisibleCards;
      if (hasVisibleCards && (state.text || state.status !== "all")) {
        section.open = true;
      }
    });
  }

  function setAllDetails(open) {
    document.querySelectorAll("[data-slide-section], [data-target-card]").forEach(function (detail) {
      detail.open = open;
    });
  }

  function startPreviewPolling() {
    var shell = document.querySelector("[data-preview-processing='1']");
    if (!shell) {
      return;
    }
    var url = shell.getAttribute("data-preview-status-url");
    if (!url) {
      return;
    }
    var previewUrl = shell.getAttribute("data-preview-url") || url.replace(/\/processing-status$/, "/preview");
    var attempts = 0;
    var timer = window.setInterval(function () {
      attempts += 1;
      fetch(url, { cache: "no-store" })
        .then(function (response) {
          if (!response.ok) {
            throw new Error("Falha ao consultar status.");
          }
          return response.json();
        })
        .then(function (state) {
          updateProcessingSlides(state);
          if (state.status === "complete" || state.status === "error" || state.active === false) {
            window.clearInterval(timer);
            window.location.replace(state.preview_url || previewUrl);
          }
        })
        .catch(function () {
          if (attempts > 8) {
            window.clearInterval(timer);
          }
        });
    }, 1800);
  }

  function updateProcessingSlides(state) {
    updatePreviewObjectProgress(state && state.progress ? state.progress : null);
    var slides = state && state.slides ? state.slides : {};
    Object.keys(slides).forEach(function (key) {
      var item = slides[key] || {};
      var slide = item.slide || key;
      var section = document.querySelector("[data-processing-slide='" + slide + "']");
      if (!section) {
        return;
      }
      setText(section, "[data-processing-slide-count]", (item.target_count || "...") + " target(s)");
      setText(section, "[data-processing-slide-status]", item.status || "");
      setText(section, "[data-processing-slide-stage]", stageLabel(item.stage || item.status || ""));
      setText(section, "[data-processing-slide-message]", item.message || "");
      setText(section, "[data-processing-slide-mapped]", item.mapped_count || 0);
      setText(section, "[data-processing-slide-targets]", item.target_count || 0);
    });
  }

  function updatePreviewObjectProgress(progress) {
    var container = document.querySelector("[data-object-progress]");
    if (!container || !progress) {
      return;
    }
    var total = parseInt(progress.total || 0, 10);
    var completed = parseInt(progress.completed || 0, 10);
    var percent = Math.max(0, Math.min(100, parseInt(progress.percent || 0, 10)));
    var bar = container.querySelector("[data-object-progress-bar]");
    container.hidden = total <= 0;
    setText(container, "[data-object-progress-label]", progress.message || "Analisando objetos do preview.");
    setText(container, "[data-object-progress-count]", completed + " de " + total + " objetos · " + percent + "%");
    if (bar) {
      bar.value = percent;
      bar.textContent = percent + "%";
    }
  }

  function setText(root, selector, value) {
    var el = root.querySelector(selector);
    if (el) {
      el.textContent = value;
    }
  }

  function stageLabel(stage) {
    if (stage === "queued") {
      return "Na fila";
    }
    if (stage === "analysis") {
      return "Analisando PPT e XLSX";
    }
    if (stage === "analysis_done") {
      return "Mapeamento pronto";
    }
    if (stage === "ai_match") {
      return "IA enxuta";
    }
    if (stage === "complete" || stage === "done") {
      return "Cards prontos";
    }
    if (stage === "error") {
      return "Erro no processamento";
    }
    return "Processando";
  }

  function parseSlideScope(value) {
    var text = (value || "").trim();
    if (!text) {
      return [];
    }
    var slides = {};
    var parts = text.split(/[,;\s]+/);
    for (var i = 0; i < parts.length; i += 1) {
      var part = parts[i].trim();
      var range = part.match(/^(\d+)[-:](\d+)$/);
      if (range) {
        var left = parseInt(range[1], 10);
        var right = parseInt(range[2], 10);
        var start = Math.min(left, right);
        var end = Math.max(left, right);
        for (var slide = start; slide <= end; slide += 1) {
          if (slide > 0) {
            slides[slide] = true;
          }
        }
      } else if (/^\d+$/.test(part)) {
        var number = parseInt(part, 10);
        if (number > 0) {
          slides[number] = true;
        }
      }
    }
    return Object.keys(slides);
  }

  function deckScopeSize() {
    var scopeInput = document.querySelector("[data-slide-scope-input]");
    var selected = parseSlideScope(scopeInput ? scopeInput.value : "");
    if (selected.length) {
      return selected.length;
    }
    return deckSummary ? deckSummary.slide_count || 0 : 0;
  }

  function updateDeckInspector() {
    var inspector = document.querySelector("[data-deck-inspector]");
    var title = document.querySelector("[data-deck-title]");
    var meta = document.querySelector("[data-deck-meta]");
    var confirm = document.querySelector("[data-large-deck-confirm]");
    var checkbox = document.querySelector("[data-large-deck-checkbox]");
    if (!inspector) {
      return;
    }
    if (!deckSummary) {
      inspector.hidden = true;
      return;
    }
    var scopeSize = deckScopeSize();
    var threshold = deckSummary.large_slide_threshold || 10;
    var requiresConfirmation = scopeSize > threshold;
    inspector.hidden = false;
    if (title) {
      title.textContent = requiresConfirmation ? "PPT grande detectado" : "PPT pronto para analise";
    }
    if (meta) {
      meta.textContent = deckSummary.slide_count + " slides, " + deckSummary.target_count + " targets (" + deckSummary.chart_count + " graficos e " + deckSummary.table_count + " tabelas). Escopo atual: " + scopeSize + " slide(s).";
    }
    if (confirm) {
      confirm.hidden = !requiresConfirmation;
    }
    if (!requiresConfirmation && checkbox) {
      checkbox.checked = false;
    }
  }

  function inspectPpt(file) {
    var inspector = document.querySelector("[data-deck-inspector]");
    var title = document.querySelector("[data-deck-title]");
    var meta = document.querySelector("[data-deck-meta]");
    deckSummary = null;
    if (!file) {
      updateDeckInspector();
      return;
    }
    if (inspector) {
      inspector.hidden = false;
    }
    if (title) {
      title.textContent = "Lendo PPT...";
    }
    if (meta) {
      meta.textContent = "Contando slides, graficos e tabelas.";
    }
    var formData = new FormData();
    formData.append("pptx", file);
    fetch("/ppt-summary", {
      method: "POST",
      body: formData,
    })
      .then(function (response) {
        if (!response.ok) {
          throw new Error("Nao consegui ler este PPT.");
        }
        return response.json();
      })
      .then(function (data) {
        deckSummary = data;
        updateDeckInspector();
      })
      .catch(function (error) {
        if (title) {
          title.textContent = "Nao consegui avaliar o PPT";
        }
        if (meta) {
          meta.textContent = error.message || "Tente selecionar o arquivo novamente.";
        }
      });
  }

  document.addEventListener("change", function (event) {
    var target = event.target;
    if (!target || !target.matches) {
      return;
    }
    if (target.matches("[data-project-select], [data-squad-select]")) {
      syncMappingTemplateOptions();
    }
    if (target.matches("[data-admin-squad-switch]") && target.form) {
      target.form.submit();
    }
    if (target.matches("[data-preview-search]")) {
      applyPreviewFilters();
    }
    if (target.matches("[data-pptx-input]")) {
      inspectPpt(target.files && target.files[0]);
    }
    if (target.matches("[data-dropzone] input[type='file']")) {
      updateFileCard(target);
      updateCombinedUploadSize();
    }
  });

  document.addEventListener("input", function (event) {
    var target = event.target;
    if (target && target.matches && target.matches("[data-preview-search]")) {
      applyPreviewFilters();
    }
    if (target && target.matches && target.matches("[data-slide-scope-input]")) {
      updateDeckInspector();
    }
  });

  document.addEventListener("submit", function (event) {
    var form = event.target;
    if (!form || !form.getAttribute) {
      return;
    }
    if (form.matches && form.matches("[data-large-deck-threshold]") && deckSummary) {
      var threshold = deckSummary.large_slide_threshold || parseInt(form.getAttribute("data-large-deck-threshold"), 10) || 10;
      var checkbox = form.querySelector("[data-large-deck-checkbox]");
      if (deckScopeSize() > threshold && checkbox && !checkbox.checked) {
        event.preventDefault();
        updateDeckInspector();
        checkbox.focus();
        return;
      }
    }
    showProgress(form.getAttribute("data-progress-message") || "Processando projeto...");
  });

  document.addEventListener("click", function (event) {
    var link = event.target && event.target.closest ? event.target.closest("a[data-progress-message]") : null;
    if (!link) {
      return;
    }
    showProgress(link.getAttribute("data-progress-message"));
    if (link.href.indexOf("/download") !== -1) {
      window.setTimeout(function () {
        var overlay = document.getElementById("progress-overlay");
        if (overlay) {
          overlay.hidden = true;
        }
      }, 20000);
    }
  });

  document.addEventListener("click", function (event) {
    var saveButton = event.target && event.target.closest ? event.target.closest("[data-save-checkpoint]") : null;
    if (saveButton) {
      event.preventDefault();
      var saveStatus = document.querySelector("[data-save-status]");
      saveButton.disabled = true;
      if (saveStatus) {
        saveStatus.textContent = "Salvando...";
      }
      fetch(saveButton.getAttribute("data-save-url"), {
        method: "POST",
        headers: { Accept: "application/json" },
      })
        .then(function (response) {
          if (!response.ok) {
            throw new Error("Não consegui salvar agora.");
          }
          return response.json();
        })
        .then(function (payload) {
          if (saveStatus) {
            saveStatus.textContent = payload.message || "Trabalho salvo.";
          }
        })
        .catch(function (error) {
          if (saveStatus) {
            saveStatus.textContent = error.message || "Falha ao salvar.";
          }
        })
        .finally(function () {
          saveButton.disabled = false;
        });
      return;
    }
    var asyncDownload = event.target && event.target.closest ? event.target.closest("[data-async-download]") : null;
    if (asyncDownload) {
      event.preventDefault();
      asyncDownload.disabled = true;
      var startedAt = Date.now();
      showProgress("Preparando a geração...");
      fetch(asyncDownload.getAttribute("data-generate-url"), { method: "POST" })
        .then(function (response) {
          if (!response.ok) {
            throw new Error("Nao foi possivel iniciar a geracao.");
          }
          return response.json();
        })
        .then(function (payload) {
          if (payload.download_url) {
            window.location.assign(payload.download_url);
            return;
          }
          var statusUrl = payload.status_url;
          if (!statusUrl) {
            throw new Error("Nao recebi a URL de acompanhamento.");
          }
          var timer = window.setInterval(function () {
            updateGenerationProgress(startedAt, {});
            fetch(statusUrl, { cache: "no-store" })
              .then(function (response) { return response.json(); })
              .then(function (state) {
                if (state.status === "complete" && state.download_url) {
                  window.clearInterval(timer);
                  window.location.assign(state.download_url);
                } else if (state.status === "error") {
                  window.clearInterval(timer);
                  asyncDownload.disabled = false;
                  var overlay = document.getElementById("progress-overlay");
                  if (overlay) { overlay.hidden = true; }
                  window.alert(state.message || "A geracao do PPT falhou.");
                } else {
                  updateGenerationProgress(startedAt, state);
                }
              })
              .catch(function () {});
          }, 1800);
        })
        .catch(function (error) {
          asyncDownload.disabled = false;
          var overlay = document.getElementById("progress-overlay");
          if (overlay) { overlay.hidden = true; }
          window.alert(error.message || "Nao foi possivel iniciar a geracao.");
        });
      return;
    }
    if (event.target && event.target.closest && event.target.closest("[data-theme-toggle]")) {
      toggleTheme();
      return;
    }
    if (event.target && event.target.closest && event.target.closest("[data-advanced-toggle]")) {
      applyAdvancedMode(!document.body.classList.contains("show-advanced"));
      return;
    }
    var chip = event.target && event.target.closest ? event.target.closest("[data-status-filter]") : null;
    if (chip) {
      document.querySelectorAll("[data-status-filter]").forEach(function (item) {
        item.classList.toggle("active", item === chip);
      });
      applyPreviewFilters();
      return;
    }
    if (event.target && event.target.closest && event.target.closest("[data-open-all]")) {
      setAllDetails(true);
      return;
    }
    if (event.target && event.target.closest && event.target.closest("[data-close-all]")) {
      setAllDetails(false);
    }
  });

  function updateGenerationProgress(startedAt, state) {
    var target = document.getElementById("progress-message");
    if (!target) {
      return;
    }
    var seconds = Math.round((Date.now() - startedAt) / 1000);
    var elapsed = seconds < 60 ? seconds + "s" : Math.floor(seconds / 60) + "min " + (seconds % 60) + "s";
    var base = (state && state.message) || "Gerando o PowerPoint...";
    setRealProgress(state && state.progress ? state.progress : null, "objetos");
    var hint = seconds > 45 ? " Decks grandes levam alguns minutos — pode deixar aberto." : "";
    target.textContent = base + " (" + elapsed + ")" + hint;
  }

  function applyAdvancedMode(on) {
    document.body.classList.toggle("show-advanced", on);
    var btn = document.querySelector("[data-advanced-toggle]");
    if (btn) {
      btn.setAttribute("aria-pressed", on ? "true" : "false");
      btn.textContent = on ? "Modo simples" : "Modo avançado";
    }
    try {
      localStorage.setItem("qwst-advanced", on ? "1" : "0");
    } catch (error) {}
  }

  function initAdvancedMode() {
    var stored = "0";
    try {
      stored = localStorage.getItem("qwst-advanced") || "0";
    } catch (error) {}
    applyAdvancedMode(stored === "1");
  }

  var CHART_FALLBACK = ["#1f7a5c", "#cc5a2a", "#2c6a94", "#96610a", "#175d46", "#b4402e"];

  function chartNum(value) {
    if (typeof value === "number") {
      return value;
    }
    if (value === null || value === undefined) {
      return NaN;
    }
    var text = String(value).trim().replace(/%/g, "").replace(/\s/g, "");
    if (!text) {
      return NaN;
    }
    if (/,\d+$/.test(text) && text.indexOf(".") !== -1) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, ".");
    }
    return parseFloat(text);
  }

  function esc(value) {
    return String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function chartLegend(names, colors) {
    return '<div class="viz-legend">' + names.map(function (name, i) {
      return '<span class="viz-key"><i style="background:' + colors[i] + '"></i>' + esc(name) + "</span>";
    }).join("") + "</div>";
  }

  function buildChartSvg(payload) {
    var kind = payload.kind || "";
    var headers = payload.headers || [];
    var rows = payload.rows || [];
    if (!kind || !rows.length) {
      return "";
    }
    var cats = rows.slice(0, 14).map(function (r) { return r[0]; });
    var names = headers.slice(1);
    var S = names.length;
    var C = cats.length;
    if (S < 1 || C < 1) {
      return "";
    }
    var vals = [];
    var s;
    var c;
    for (s = 0; s < S; s += 1) {
      vals[s] = [];
      for (c = 0; c < C; c += 1) {
        var n = chartNum(rows[c][s + 1]);
        vals[s][c] = isNaN(n) ? 0 : n;
      }
    }
    var colors = [];
    for (s = 0; s < S; s += 1) {
      var xml = (payload.colors || [])[s];
      colors[s] = xml && xml.length ? xml : CHART_FALLBACK[s % CHART_FALLBACK.length];
    }
    var legend = S > 1 ? chartLegend(names, colors) : "";
    var W = 680;
    var H = 250;
    var svg;
    if (kind === "pie") {
      svg = pieSvg(cats, vals[0], W, H);
      return legend + svg;
    }
    if (kind === "line" || kind === "area") {
      svg = lineSvg(kind, cats, vals, colors, W, H);
      return legend + svg;
    }
    svg = barSvg(kind, cats, vals, colors, W, H);
    return legend + svg;
  }

  function niceMax(value) {
    if (value <= 0) {
      return 1;
    }
    var pow = Math.pow(10, Math.floor(Math.log10(value)));
    var n = value / pow;
    var step = n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10;
    return step * pow;
  }

  function catLabels(cats, x0, plotW, y, horizontal) {
    return cats.map(function (cat, i) {
      var text = String(cat === null || cat === undefined ? "" : cat);
      if (text.length > 16) {
        text = text.slice(0, 15) + "…";
      }
      var cx = x0 + (plotW / cats.length) * (i + 0.5);
      return '<text x="' + cx.toFixed(1) + '" y="' + y + '" text-anchor="middle" class="chart-cat">' + esc(text) + "</text>";
    }).join("");
  }

  function barSvg(kind, cats, vals, colors, W, H) {
    var horizontal = kind.indexOf("bar") === 0;
    var stacked = kind.indexOf("stacked") !== -1;
    var S = vals.length;
    var C = cats.length;
    var padL = horizontal ? 4 : 6;
    var padR = 6;
    var padT = 8;
    var padB = horizontal ? 8 : 40;
    var plotW = W - padL - padR;
    var plotH = H - padT - padB;
    var maxV;
    if (stacked) {
      maxV = 0;
      for (var ci = 0; ci < C; ci += 1) {
        var sum = 0;
        for (var si = 0; si < S; si += 1) { sum += Math.max(0, vals[si][ci]); }
        maxV = Math.max(maxV, sum);
      }
    } else {
      maxV = 0;
      for (var s2 = 0; s2 < S; s2 += 1) {
        for (var c2 = 0; c2 < C; c2 += 1) { maxV = Math.max(maxV, Math.abs(vals[s2][c2])); }
      }
    }
    maxV = niceMax(maxV);
    var parts = [];
    if (!horizontal) {
      var groupW = plotW / C;
      var baseY = padT + plotH;
      parts.push('<line x1="' + padL + '" y1="' + baseY + '" x2="' + (padL + plotW) + '" y2="' + baseY + '" class="chart-axis"/>');
      for (var g = 0; g < C; g += 1) {
        var gx = padL + groupW * g;
        if (stacked) {
          var acc = 0;
          for (var st = 0; st < S; st += 1) {
            var v = Math.max(0, vals[st][g]);
            var h = (v / maxV) * plotH;
            var yTop = baseY - (acc + v) / maxV * plotH;
            var bw = groupW * 0.6;
            parts.push('<rect x="' + (gx + groupW * 0.2).toFixed(1) + '" y="' + yTop.toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + h.toFixed(1) + '" fill="' + colors[st] + '"/>');
            acc += v;
          }
        } else {
          var bwEach = (groupW * 0.72) / S;
          for (var sc = 0; sc < S; sc += 1) {
            var vv = Math.abs(vals[sc][g]);
            var hh = (vv / maxV) * plotH;
            var bx = gx + groupW * 0.14 + bwEach * sc;
            parts.push('<rect x="' + bx.toFixed(1) + '" y="' + (baseY - hh).toFixed(1) + '" width="' + (bwEach * 0.86).toFixed(1) + '" height="' + hh.toFixed(1) + '" fill="' + colors[sc] + '"/>');
          }
        }
      }
      parts.push(catLabels(cats, padL, plotW, H - 22, false));
    } else {
      var groupH = plotH / C;
      var baseX = padL;
      parts.push('<line x1="' + baseX + '" y1="' + padT + '" x2="' + baseX + '" y2="' + (padT + plotH) + '" class="chart-axis"/>');
      for (var gr = 0; gr < C; gr += 1) {
        var gy = padT + groupH * gr;
        var bhEach = (groupH * 0.72) / (stacked ? 1 : S);
        if (stacked) {
          var accx = 0;
          for (var stk = 0; stk < S; stk += 1) {
            var vs = Math.max(0, vals[stk][gr]);
            var ws = (vs / maxV) * plotW;
            parts.push('<rect x="' + (baseX + accx / maxV * plotW).toFixed(1) + '" y="' + (gy + groupH * 0.14).toFixed(1) + '" width="' + ws.toFixed(1) + '" height="' + (groupH * 0.72).toFixed(1) + '" fill="' + colors[stk] + '"/>');
            accx += vs;
          }
        } else {
          for (var scb = 0; scb < S; scb += 1) {
            var vb = Math.abs(vals[scb][gr]);
            var wb = (vb / maxV) * plotW;
            var by = gy + groupH * 0.14 + bhEach * scb;
            parts.push('<rect x="' + baseX + '" y="' + by.toFixed(1) + '" width="' + wb.toFixed(1) + '" height="' + (bhEach * 0.86).toFixed(1) + '" fill="' + colors[scb] + '"/>');
          }
        }
      }
    }
    return '<svg class="chart-svg" viewBox="0 0 ' + W + " " + H + '" preserveAspectRatio="xMidYMid meet" role="img">' + parts.join("") + "</svg>";
  }

  function lineSvg(kind, cats, vals, colors, W, H) {
    var S = vals.length;
    var C = cats.length;
    var padL = 6;
    var padR = 6;
    var padT = 8;
    var padB = 40;
    var plotW = W - padL - padR;
    var plotH = H - padT - padB;
    var maxV = 0;
    for (var s = 0; s < S; s += 1) {
      for (var c = 0; c < C; c += 1) { maxV = Math.max(maxV, Math.abs(vals[s][c])); }
    }
    maxV = niceMax(maxV);
    var baseY = padT + plotH;
    var stepX = C > 1 ? plotW / (C - 1) : plotW;
    var parts = ['<line x1="' + padL + '" y1="' + baseY + '" x2="' + (padL + plotW) + '" y2="' + baseY + '" class="chart-axis"/>'];
    for (var si = 0; si < S; si += 1) {
      var pts = [];
      for (var ci = 0; ci < C; ci += 1) {
        var x = padL + (C > 1 ? stepX * ci : plotW / 2);
        var y = baseY - (Math.abs(vals[si][ci]) / maxV) * plotH;
        pts.push(x.toFixed(1) + "," + y.toFixed(1));
      }
      if (kind === "area") {
        parts.push('<polygon points="' + padL + "," + baseY + " " + pts.join(" ") + " " + (padL + plotW) + "," + baseY + '" fill="' + colors[si] + '" fill-opacity="0.18"/>');
      }
      parts.push('<polyline points="' + pts.join(" ") + '" fill="none" stroke="' + colors[si] + '" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>');
      for (var pi = 0; pi < pts.length; pi += 1) {
        var xy = pts[pi].split(",");
        parts.push('<circle cx="' + xy[0] + '" cy="' + xy[1] + '" r="2.6" fill="' + colors[si] + '"/>');
      }
    }
    parts.push(catLabels(cats, padL, plotW, H - 22, false));
    return '<svg class="chart-svg" viewBox="0 0 ' + W + " " + H + '" preserveAspectRatio="xMidYMid meet" role="img">' + parts.join("") + "</svg>";
  }

  function pieSvg(cats, series, W, H) {
    var total = 0;
    series.forEach(function (v) { total += Math.max(0, v); });
    if (total <= 0) {
      return "";
    }
    var cx = H / 2;
    var cy = H / 2;
    var r = H / 2 - 12;
    var angle = -Math.PI / 2;
    var parts = [];
    var legend = [];
    for (var i = 0; i < cats.length; i += 1) {
      var v = Math.max(0, series[i]);
      if (v <= 0) { continue; }
      var frac = v / total;
      var end = angle + frac * Math.PI * 2;
      var color = CHART_FALLBACK[i % CHART_FALLBACK.length];
      var large = frac > 0.5 ? 1 : 0;
      var x1 = cx + r * Math.cos(angle);
      var y1 = cy + r * Math.sin(angle);
      var x2 = cx + r * Math.cos(end);
      var y2 = cy + r * Math.sin(end);
      parts.push('<path d="M' + cx + " " + cy + " L" + x1.toFixed(1) + " " + y1.toFixed(1) + " A" + r + " " + r + " 0 " + large + " 1 " + x2.toFixed(1) + " " + y2.toFixed(1) + ' Z" fill="' + color + '"/>');
      var label = String(cats[i] === null || cats[i] === undefined ? "" : cats[i]);
      if (label.length > 18) { label = label.slice(0, 17) + "…"; }
      legend.push('<span class="viz-key"><i style="background:' + color + '"></i>' + esc(label) + "</span>");
      angle = end;
    }
    var svg = '<svg class="chart-svg" viewBox="0 0 ' + W + " " + H + '" preserveAspectRatio="xMidYMid meet" role="img"><g>' + parts.join("") + "</g></svg>";
    return '<div class="viz-legend">' + legend.join("") + "</div>" + svg;
  }

  function renderCharts() {
    document.querySelectorAll("[data-chart]").forEach(function (chart) {
      if (chart.getAttribute("data-chart-done") === "1") {
        return;
      }
      var dataEl = chart.querySelector("[data-chart-data]");
      var canvas = chart.querySelector("[data-chart-canvas]");
      if (!dataEl || !canvas) {
        return;
      }
      var payload;
      try {
        payload = JSON.parse(dataEl.textContent || "{}");
      } catch (error) {
        return;
      }
      var markup = buildChartSvg(payload);
      if (!markup) {
        chart.hidden = true;
        return;
      }
      canvas.innerHTML = markup;
      chart.setAttribute("data-chart-done", "1");
    });
  }

  syncMappingTemplateOptions();
  applyPreviewFilters();
  startPreviewPolling();
  bindDropzones();
  initAdvancedMode();
  renderCharts();
})();
