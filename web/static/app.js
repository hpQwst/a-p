(function () {
  var deckSummary = null;

  function showProgress(message) {
    var overlay = document.getElementById("progress-overlay");
    var label = document.getElementById("progress-message");
    if (!overlay) {
      return;
    }
    if (label) {
      label.textContent = message || "Processando...";
    }
    overlay.hidden = false;
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
})();
