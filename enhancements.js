(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const options = $("docOptions");
  const updateRecipeDefaults = () => {
    if (!options) return;
    const inputs = [...options.querySelectorAll("input[type=checkbox]")];
    if (!inputs.length) return;
    inputs.forEach(input => {
      const label = input.closest("label");
      input.checked = /receta/i.test(label?.textContent || input.value);
    });
    const count = $("docCount");
    if (count) count.textContent = `${inputs.filter(input => input.checked).length} de ${inputs.length} seleccionados`;
  };
  if (options) {
    new MutationObserver(updateRecipeDefaults).observe(options, { childList: true, subtree: true });
    updateRecipeDefaults();
  }

  const table = $("resultsTable");
  const getRows = () => [...(table?.tBodies[0]?.rows || [])];

  const enhanceRows = () => {
    if (!table) return;

    const header = table.tHead?.rows[0];
    if (header && !header.dataset.columnsReduced) {
      // Elimina # correlativo y Código interno; conserva Código como dato para el buscador.
      while (header.cells.length > 6) header.deleteCell(0);
      header.dataset.columnsReduced = "true";
    }

    getRows().forEach(row => {
      if (!row.dataset.columnsReduced) {
        if (row.cells.length >= 8) {
          row.dataset.productCode = row.cells[1].textContent.trim();
        }
        while (row.cells.length > 6) row.deleteCell(0);
        row.dataset.columnsReduced = "true";
      }

      const article = row.cells[0];
      if (article) {
        article.style.whiteSpace = "normal";
        article.style.overflowWrap = "anywhere";
        const length = article.textContent.trim().length;
        article.style.fontSize = length > 75 ? "6.2pt" : length > 55 ? "6.7pt" : "7pt";
        article.style.lineHeight = "1.05";
      }

      const availability = row.cells[4];
      if (availability) {
        availability.classList.remove("availability-low", "availability-medium", "availability-high");
        const value = Number(
          availability.textContent
            .replace("%", "")
            .replace(/\./g, "")
            .replace(",", ".")
            .trim()
        );
        if (Number.isFinite(value)) {
          if (value < 50) availability.classList.add("availability-low");
          else if (value > 50 && value < 100) availability.classList.add("availability-medium");
          else if (value > 100) availability.classList.add("availability-high");
        }
      }
    });

    applyProductFilter();
  };

  if (table) {
    new MutationObserver(enhanceRows).observe(table, { childList: true, subtree: true });
    enhanceRows();
  }

  const reportTab = $("reportTab");
  const searchTab = $("productSearchTab");
  const searchPanel = $("productSearchPanel");
  const searchInput = $("productSearchInput");
  let searching = false;

  function applyProductFilter() {
    if (!table) return;
    const term = (searching ? searchInput?.value : "")?.trim().toLocaleLowerLowerCase?.("es-CL") || "";
    let visible = 0;
    getRows().forEach(row => {
      const product = row.cells[0]?.textContent || "";
      const code = row.dataset.productCode || "";
      const match = !term || `${product} ${code}`.toLocaleLowerCase("es-CL").includes(term);
      row.hidden = !match;
      if (match) visible++;
    });
    const message = $("productSearchMessage");
    if (message && searching) {
      message.textContent = term ? `${visible} producto(s) encontrado(s).` : `${visible} productos en el informe.`;
    }
  }

  function setSearchMode(enabled) {
    searching = enabled;
    searchPanel?.classList.toggle("hidden", !enabled);
    reportTab?.classList.toggle("active", !enabled);
    searchTab?.classList.toggle("active", enabled);
    reportTab?.setAttribute("aria-selected", String(!enabled));
    searchTab?.setAttribute("aria-selected", String(enabled));
    if (!enabled && searchInput) searchInput.value = "";
    applyProductFilter();
    if (enabled) searchInput?.focus();
  }

  reportTab?.addEventListener("click", () => setSearchMode(false));
  searchTab?.addEventListener("click", () => setSearchMode(true));
  searchInput?.addEventListener("input", applyProductFilter);
})();
