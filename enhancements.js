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
  if (options) new MutationObserver(updateRecipeDefaults).observe(options, { childList: true, subtree: true });

  const table = $("resultsTable");
  const getRows = () => [...(table?.tBodies[0]?.rows || [])];
  const enhanceRows = () => {
    if (!table) return;
    const header = table.tHead?.rows[0];
    if (header && header.cells.length > 6) {
      header.deleteCell(0); // correlativo
      header.deleteCell(0); // código interno de cruce
    }
    getRows().forEach(row => {
      if (!row.dataset.productCode && row.cells.length >= 8) row.dataset.productCode = row.cells[1].textContent.trim();
      while (row.cells.length > 6) row.deleteCell(0);
      const availability = row.cells[4];
      if (!availability) return;
      availability.classList.remove("availability-low", "availability-medium", "availability-high");
      const value = Number(availability.textContent.replace("%", "").replace(/\./g, "").replace(",", ".").trim());
      if (!Number.isFinite(value)) return;
      if (value < 50) availability.classList.add("availability-low");
      else if (value > 50 && value < 100) availability.classList.add("availability-medium");
      else if (value > 100) availability.classList.add("availability-high");
    });
    applyProductFilter();
  };
  if (table) new MutationObserver(enhanceRows).observe(table, { childList: true, subtree: true });

  const reportTab = $("reportTab"), searchTab = $("productSearchTab"), searchPanel = $("productSearchPanel"), searchInput = $("productSearchInput");
  let searching = false;
  function applyProductFilter() {
    if (!table) return;
    const term = (searching ? searchInput?.value : "")?.trim().toLocaleLowerCase("es-CL") || "";
    let visible = 0;
    getRows().forEach(row => {
      const product = row.cells[0]?.textContent || "";
      const code = row.dataset.productCode || "";
      const match = !term || `${product} ${code}`.toLocaleLowerCase("es-CL").includes(term);
      row.hidden = !match;
      if (match) visible++;
    });
    const message = $("productSearchMessage");
    if (message && searching) message.textContent = term ? `${visible} producto(s) encontrado(s).` : `${visible} productos en el informe.`;
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
