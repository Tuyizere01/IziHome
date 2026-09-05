(function () {
  const slides = Array.from(document.querySelectorAll(".slide"));
  const dots = document.getElementById("dots");
  const counter = document.getElementById("counter");
  const deck = document.getElementById("deck");
  let i = 0;

  slides.forEach((slide, n) => {
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.title = slide.dataset.label || "Slide " + (n + 1);
    btn.addEventListener("click", () => go(n));
    li.appendChild(btn);
    dots.appendChild(li);
  });

  function go(n) {
    i = (n + slides.length) % slides.length;
    slides.forEach((s, k) => s.classList.toggle("is-on", k === i));
    dots.querySelectorAll("button").forEach((b, k) => b.classList.toggle("on", k === i));
    counter.textContent = i + 1 + " / " + slides.length;
  }

  document.getElementById("prev").addEventListener("click", () => go(i - 1));
  document.getElementById("next").addEventListener("click", () => go(i + 1));
  document.getElementById("fs").addEventListener("click", () => {
    if (!document.fullscreenElement) document.documentElement.requestFullscreen();
    else document.exitFullscreen();
  });

  document.addEventListener("keydown", (e) => {
    if (["ArrowRight", "ArrowDown", "PageDown", " "].includes(e.key)) {
      e.preventDefault();
      go(i + 1);
    }
    if (["ArrowLeft", "ArrowUp", "PageUp"].includes(e.key)) {
      e.preventDefault();
      go(i - 1);
    }
    if (e.key === "Home") go(0);
    if (e.key === "End") go(slides.length - 1);
    if (e.key === "f") document.getElementById("fs").click();
  });

  let touchX = 0;
  deck.addEventListener("touchstart", (e) => { touchX = e.changedTouches[0].screenX; }, { passive: true });
  deck.addEventListener("touchend", (e) => {
    const dx = e.changedTouches[0].screenX - touchX;
    if (dx < -40) go(i + 1);
    if (dx > 40) go(i - 1);
  });

  go(0);
  deck.focus();
})();
