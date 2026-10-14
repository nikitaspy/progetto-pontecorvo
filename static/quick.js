const quantityInput = document.querySelector("#quick-quantity");
const form = document.querySelector("#quick-movement-form");
const primaryButton = document.querySelector(".quick-primary");

function normalizedQuantity() {
  return Math.max(1, Number.parseInt(quantityInput.value, 10) || 1);
}

function updateQuantity(value) {
  quantityInput.value = Math.max(1, value);
  if (primaryButton && !primaryButton.disabled) {
    const quantity = normalizedQuantity();
    primaryButton.textContent = `Preleva ${quantity} ${quantity === 1 ? "pezzo" : "pezzi"}`;
  }
}

document.querySelectorAll("[data-step]").forEach((button) => {
  button.addEventListener("click", () => {
    updateQuantity(normalizedQuantity() + Number.parseInt(button.dataset.step, 10));
  });
});

document.querySelectorAll("[data-quantity]").forEach((button) => {
  button.addEventListener("click", () => updateQuantity(Number.parseInt(button.dataset.quantity, 10)));
});

quantityInput?.addEventListener("input", () => updateQuantity(normalizedQuantity()));

form?.addEventListener("submit", () => {
  window.setTimeout(() => {
    form.querySelectorAll("button[type='submit']").forEach((button) => {
      button.disabled = true;
    });
  }, 0);
});
