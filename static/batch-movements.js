const batchLines = document.querySelector("#batch-lines");
const lineTemplate = document.querySelector("#batch-line-template");
const addLineButton = document.querySelector("#add-batch-line");
const batchForm = document.querySelector("#batch-movement-form");
const submitButton = document.querySelector("#submit-batch-movement");

function refreshLineNumbers() {
  const lines = [...batchLines.querySelectorAll(".batch-line")];
  lines.forEach((line, index) => {
    line.querySelector(".batch-line-number").textContent = String(index + 1);
    line.querySelector(".batch-remove-line").disabled = lines.length === 1;
  });
}

addLineButton?.addEventListener("click", () => {
  const fragment = lineTemplate.content.cloneNode(true);
  batchLines.append(fragment);
  refreshLineNumbers();
  const newLine = batchLines.lastElementChild;
  newLine?.querySelector("select")?.focus();
});

batchLines?.addEventListener("click", (event) => {
  const removeButton = event.target.closest(".batch-remove-line");
  if (!removeButton || removeButton.disabled) return;
  removeButton.closest(".batch-line").remove();
  refreshLineNumbers();
});

batchForm?.addEventListener("submit", () => {
  window.setTimeout(() => {
    submitButton.disabled = true;
    submitButton.textContent = "Registrazione…";
  }, 0);
});

refreshLineNumbers();
