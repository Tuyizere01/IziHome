const nav = document.querySelector(".nav");
const toggle = document.querySelector(".menu-toggle");
const year = document.querySelector("#year");
const form = document.querySelector("#booking-form");
const toast = document.querySelector(".toast");

if (year) year.textContent = new Date().getFullYear();

toggle?.addEventListener("click", () => nav?.classList.toggle("open"));
document.querySelectorAll(".nav a").forEach((link) => {
  link.addEventListener("click", () => nav?.classList.remove("open"));
});

form?.addEventListener("submit", (event) => {
  event.preventDefault();
  form.reset();
  toast?.classList.add("show");
  setTimeout(() => toast?.classList.remove("show"), 3200);
});
