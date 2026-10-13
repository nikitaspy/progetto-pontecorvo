(() => {
  const selectAll = document.querySelector("#select-all-components");
  const checkboxes = Array.from(document.querySelectorAll(".component-checkbox"));
  const count = document.querySelector("#selection-count");

  if (!selectAll || !count || checkboxes.length === 0) return;

  const refresh = () => {
    const selected = checkboxes.filter((checkbox) => checkbox.checked).length;
    selectAll.checked = selected === checkboxes.length;
    selectAll.indeterminate = selected > 0 && selected < checkboxes.length;
    count.textContent = `${selected} ${selected === 1 ? "selezionato" : "selezionati"}`;
  };

  selectAll.addEventListener("change", () => {
    checkboxes.forEach((checkbox) => {
      checkbox.checked = selectAll.checked;
    });
    refresh();
  });

  checkboxes.forEach((checkbox) => checkbox.addEventListener("change", refresh));
  refresh();
})();
