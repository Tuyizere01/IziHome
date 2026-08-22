const header = document.querySelector(".header");
const nav = document.querySelector(".nav");
const toggle = document.querySelector(".menu-toggle");
const year = document.querySelector("#year");
const toast = document.querySelector(".toast");
const form = document.querySelector("#booking-form");
const preview = document.querySelector("#selected-preview");
const packageSelect = document.querySelector("#package-select");
const ingredientsSelect = document.querySelector("#ingredients-select");

if (year) year.textContent = new Date().getFullYear();

const syncHeader = () => {
  header?.classList.toggle("is-scrolled", window.scrollY > 20);
};

syncHeader();
window.addEventListener("scroll", syncHeader, { passive: true });

toggle?.addEventListener("click", () => {
  nav?.classList.toggle("open");
  header?.classList.toggle("is-open", nav?.classList.contains("open"));
});
document.querySelectorAll(".nav a").forEach((link) => {
  link.addEventListener("click", () => nav?.classList.remove("open"));
});

const getSelected = () => JSON.parse(localStorage.getItem("homefreshDishes") || "[]");
const getPackage = () => localStorage.getItem("homefreshPackage") || "";
const getIngredients = () => localStorage.getItem("homefreshIngredients") || "have";

const names = {
  solo: "Fresh Solo",
  duo: "Fresh Duo",
  family: "Fresh Family",
  cook: "Cook With Me",
  train: "Train My Home Helper"
};

if (preview && typeof DISHES !== "undefined") {
  const selected = getSelected();
  const dishes = DISHES.filter((dish) => selected.includes(dish.id)).map((dish) => dish.name);
  preview.innerHTML = dishes.length
    ? `Selected dishes: <strong>${dishes.join(", ")}</strong>. <a href="menu.html">Edit menu</a>.`
    : `No dishes selected yet. <a href="menu.html">Browse the menu</a>.`;
}

if (packageSelect && getPackage()) packageSelect.value = getPackage();
if (ingredientsSelect) ingredientsSelect.value = getIngredients();

document.querySelectorAll(".choice").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".choice").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    localStorage.setItem("homefreshIngredients", button.dataset.ingredients);
    if (ingredientsSelect) ingredientsSelect.value = button.dataset.ingredients;
  });
});

document.querySelectorAll('a[href^="menu.html?package="]').forEach((link) => {
  link.addEventListener("click", () => {
    const pkg = new URL(link.href, window.location.href).searchParams.get("package");
    if (pkg) localStorage.setItem("homefreshPackage", pkg);
  });
});

form?.addEventListener("submit", (event) => {
  event.preventDefault();
  localStorage.removeItem("homefreshDishes");
  form.reset();
  toast?.classList.add("show");
  setTimeout(() => toast?.classList.remove("show"), 3200);
});
