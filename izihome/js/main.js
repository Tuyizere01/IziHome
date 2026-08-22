const nav = document.querySelector(".nav");
const toggle = document.querySelector(".menu-toggle");
const year = document.querySelector("#year");
if (year) year.textContent = new Date().getFullYear();
toggle?.addEventListener("click", () => nav?.classList.toggle("open"));
document.querySelectorAll(".nav a").forEach((link) => {
  link.addEventListener("click", () => nav?.classList.remove("open"));
});
