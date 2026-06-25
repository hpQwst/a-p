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
})();
