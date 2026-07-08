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
    var file = input.files && input.files[0];
    card.classList.toggle("has-file", !!file);
    if (label) {
      label.textContent = file ? file.name + " (" + formatBytes(file.size) + ")" : "";
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
    if (!overlay) {
      return;
    }
    if (label) {
      label.textContent = message || "Processando...";
    }
    if (steps) {
      renderProgressSteps(steps, progressStepsForMessage(message || ""));
    }
    overlay.hidden = false;
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
    if (target.matches("[data-preview-search]")) {
      applyPreviewFilters();
    }
    if (target.matches("[data-pptx-input]")) {
      inspectPpt(target.files && target.files[0]);
    }
    if (target.matches("[data-dropzone] input[type='file']")) {
      updateFileCard(target);
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
    if (event.target && event.target.closest && event.target.closest("[data-theme-toggle]")) {
      toggleTheme();
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

  syncMappingTemplateOptions();
  applyPreviewFilters();
  startPreviewPolling();
  bindDropzones();
})();
