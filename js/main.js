const nav = document.querySelector(".nav");
const toggle = document.querySelector(".menu-toggle");
const form = document.querySelector("#booking-form");
const toast = document.querySelector(".toast");
const year = document.querySelector("#year");

if (year) {
  year.textContent = new Date().getFullYear();
}

toggle?.addEventListener("click", () => {
  nav?.classList.toggle("open");
});

document.querySelectorAll('.nav a').forEach((link) => {
  link.addEventListener("click", () => nav?.classList.remove("open"));
});

const sections = document.querySelectorAll("section[id], article[id]");
const navLinks = document.querySelectorAll(".nav a");

const setActiveLink = () => {
  let current = "home";
  sections.forEach((section) => {
    if (window.scrollY >= section.offsetTop - 140) {
      current = section.id;
    }
  });
  navLinks.forEach((link) => {
    link.classList.toggle("active", link.getAttribute("href") === `#${current}`);
  });
};

if (!location.pathname.includes("15-day")) {
  window.addEventListener("scroll", setActiveLink);
  setActiveLink();
}

form?.addEventListener("submit", (event) => {
  event.preventDefault();
  form.reset();
  toast?.classList.add("show");
  setTimeout(() => toast?.classList.remove("show"), 3200);
});
