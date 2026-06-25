(function () {
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

  document.addEventListener("change", function (event) {
    var target = event.target;
    if (!target || !target.matches) {
      return;
    }
    if (target.matches("[data-project-select], [data-squad-select]")) {
      syncMappingTemplateOptions();
    }
  });

  document.addEventListener("submit", function (event) {
    var form = event.target;
    if (!form || !form.getAttribute) {
      return;
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

  syncMappingTemplateOptions();
})();
