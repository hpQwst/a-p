(function () {
  "use strict";

  var root = document.querySelector("[data-model-studio]");
  if (!root) return;

  var rows = Array.prototype.slice.call(root.querySelectorAll("[data-studio-row]"));
  var status = root.querySelector("[data-studio-save-status]");
  var saveButton = root.querySelector("[data-studio-save]");
  var timer = null;
  var saving = false;
  var dirty = false;

  function rowPayload(row) {
    var payload = { id_objeto: row.getAttribute("data-target-id") };
    row.querySelectorAll("[data-field]").forEach(function (field) {
      payload[field.getAttribute("data-field")] = field.value;
    });
    return payload;
  }

  function setStatus(text, kind) {
    if (!status) return;
    status.textContent = text;
    status.classList.toggle("error", kind === "error");
  }

  function save() {
    if (saving) return Promise.resolve(false);
    saving = true;
    setStatus("Salvando...", "");
    return fetch(root.getAttribute("data-save-url"), {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      body: JSON.stringify({ objects: rows.map(rowPayload) })
    }).then(function (response) {
      return response.json().then(function (data) {
        if (!response.ok || !data.ok) throw new Error(data.error || "Falha ao salvar.");
        dirty = false;
        setStatus("Salvo automaticamente", "");
        return true;
      });
    }).catch(function (error) {
      setStatus(error.message || "Falha ao salvar", "error");
      return false;
    }).finally(function () {
      saving = false;
    });
  }

  function scheduleSave() {
    dirty = true;
    setStatus("Alteracoes pendentes", "");
    window.clearTimeout(timer);
    timer = window.setTimeout(save, 1200);
  }

  function selectTarget(targetId) {
    rows.forEach(function (row) {
      row.classList.toggle("selected", row.getAttribute("data-target-id") === targetId);
    });
    root.querySelectorAll("[data-wire-target]").forEach(function (wire) {
      wire.classList.toggle("selected", wire.getAttribute("data-wire-target") === targetId);
    });
  }

  root.addEventListener("input", function (event) {
    if (event.target && event.target.matches("[data-field]")) scheduleSave();
  });
  root.addEventListener("change", function (event) {
    if (event.target && event.target.matches("[data-field]")) {
      var row = event.target.closest("[data-studio-row]");
      if (row && event.target.getAttribute("data-field") === "ativo") {
        row.classList.toggle("is-active", event.target.value === "1");
      }
      scheduleSave();
    }
  });
  if (saveButton) saveButton.addEventListener("click", save);

  root.addEventListener("click", function (event) {
    var wire = event.target.closest("[data-wire-target]");
    if (wire) {
      var targetId = wire.getAttribute("data-wire-target");
      selectTarget(targetId);
      var row = root.querySelector('[data-studio-row][data-target-id="' + CSS.escape(targetId) + '"]');
      if (row) row.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
      return;
    }
    var row = event.target.closest("[data-studio-row]");
    if (row) selectTarget(row.getAttribute("data-target-id"));
  });

  var search = root.querySelector("[data-studio-search]");
  var currentFilter = "all";
  function applyFilter() {
    var term = search ? search.value.trim().toLowerCase() : "";
    rows.forEach(function (row) {
      var active = row.querySelector('[data-field="ativo"]').value === "1";
      var file = row.querySelector('[data-field="arquivo_xlsx"]').value.trim();
      var filterMatch = currentFilter === "all" ||
        (currentFilter === "active" && active) ||
        (currentFilter === "missing" && active && !file);
      var searchMatch = !term || (row.getAttribute("data-search-text") || "").toLowerCase().indexOf(term) >= 0;
      row.hidden = !(filterMatch && searchMatch);
    });
  }
  if (search) search.addEventListener("input", applyFilter);
  root.querySelectorAll("[data-studio-filter]").forEach(function (button) {
    button.addEventListener("click", function () {
      currentFilter = button.getAttribute("data-studio-filter");
      root.querySelectorAll("[data-studio-filter]").forEach(function (item) {
        item.classList.toggle("active", item === button);
      });
      applyFilter();
    });
  });

  root.querySelectorAll("[data-save-before-download]").forEach(function (link) {
    link.addEventListener("click", function (event) {
      if (!dirty) return;
      event.preventDefault();
      var href = link.href;
      save().then(function (ok) { if (ok) window.location.href = href; });
    });
  });

  var importForm = root.querySelector("[data-studio-import]");
  if (importForm) {
    importForm.addEventListener("submit", function (event) {
      if (!dirty) return;
      event.preventDefault();
      save().then(function (ok) {
        if (ok) importForm.requestSubmit();
      });
    });
  }
})();
