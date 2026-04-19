(function () {
  try {
    var stored = localStorage.getItem("urnik.theme.v1");
    if (stored !== "light" && stored !== "dark") {
      var prefersDark =
        window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
      stored = prefersDark ? "dark" : "light";
    }
    var root = document.documentElement;
    root.dataset.theme = stored;
    if (stored === "dark") root.classList.add("is-dark");
  } catch (_) {}
})();
