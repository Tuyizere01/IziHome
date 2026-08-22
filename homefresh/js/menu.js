const grid = document.querySelector("#dish-grid");
const cartList = document.querySelector("#cart-list");
const cartEmpty = document.querySelector("#cart-empty");
const cartCount = document.querySelector("#cart-count");
const search = document.querySelector("#search");
const packageNote = document.querySelector("#package-note");

const params = new URLSearchParams(window.location.search);
if (params.get("package")) {
  localStorage.setItem("homefreshPackage", params.get("package"));
}

const packageLabels = {
  solo: "Fresh Solo — pick about 3 dishes",
  duo: "Fresh Duo — pick about 4 dishes",
  family: "Fresh Family — pick about 5 dishes"
};

const savedPackage = localStorage.getItem("homefreshPackage");
if (packageNote && savedPackage && packageLabels[savedPackage]) {
  packageNote.textContent = packageLabels[savedPackage];
}

let filter = "all";
let query = "";
let selected = JSON.parse(localStorage.getItem("homefreshDishes") || "[]");

const save = () => localStorage.setItem("homefreshDishes", JSON.stringify(selected));

const matches = (dish) => {
  const haystack = `${dish.name} ${dish.region} ${dish.ingredients} ${dish.description}`.toLowerCase();
  const searchOk = haystack.includes(query);
  if (!searchOk) return false;
  if (filter === "all") return true;
  if (filter === "african" || filter === "world") return dish.cuisine === filter;
  return dish.region === filter;
};

const renderCart = () => {
  const items = DISHES.filter((dish) => selected.includes(dish.id));
  cartCount.textContent = String(items.length);
  cartEmpty.style.display = items.length ? "none" : "block";
  cartList.innerHTML = items
    .map((dish) => `<li>${dish.name} <button type="button" data-remove="${dish.id}">Remove</button></li>`)
    .join("");
};

const renderGrid = () => {
  const dishes = DISHES.filter(matches);
  grid.innerHTML = dishes
    .map((dish) => {
      const on = selected.includes(dish.id);
      return `
        <article class="dish ${on ? "selected" : ""}" data-id="${dish.id}">
          <img src="${dish.image}" alt="${dish.name}">
          <div class="dish-body">
            <span class="badge ${dish.cuisine === "world" ? "world" : ""}">${dish.region}</span>
            <h3>${dish.name}</h3>
            <p>${dish.description}</p>
            <p class="ingredients"><strong>Ingredients:</strong> ${dish.ingredients}</p>
            <button class="btn ${on ? "btn-outline" : "btn-orange"}" type="button" data-toggle="${dish.id}">
              ${on ? "Added to my menu" : "Add to my menu"}
            </button>
          </div>
        </article>`;
    })
    .join("");
};

const toggleDish = (id) => {
  selected = selected.includes(id) ? selected.filter((item) => item !== id) : [...selected, id];
  save();
  renderGrid();
  renderCart();
};

document.querySelector("#filters")?.addEventListener("click", (event) => {
  const button = event.target.closest(".filter-btn");
  if (!button) return;
  filter = button.dataset.filter;
  document.querySelectorAll(".filter-btn").forEach((item) => item.classList.toggle("active", item === button));
  renderGrid();
});

search?.addEventListener("input", () => {
  query = search.value.trim().toLowerCase();
  renderGrid();
});

grid?.addEventListener("click", (event) => {
  const id = event.target.dataset.toggle;
  if (id) toggleDish(id);
});

cartList?.addEventListener("click", (event) => {
  const id = event.target.dataset.remove;
  if (id) toggleDish(id);
});

renderGrid();
renderCart();
