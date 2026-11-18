/* "Back to top" button: appears after scrolling down, scrolls smoothly back to the top. */
(function () {
  const css = document.createElement("style");
  css.textContent =
    ".totop{position:fixed;right:16px;bottom:16px;z-index:900;width:52px;height:52px;border-radius:50%;border:3px solid #fff;background:var(--blue,#2a3bff);color:#fff;" +
    "display:grid;place-items:center;cursor:pointer;padding:0;box-shadow:0 4px 14px rgba(42,59,255,.35);opacity:0;transform:translateY(12px);pointer-events:none;transition:opacity .2s,transform .2s}" +
    ".totop.show{opacity:1;transform:none;pointer-events:auto}.totop:hover{transform:translateY(-2px)}" +
    ".totop:focus-visible{outline:3px solid var(--blue,#2a3bff);outline-offset:3px}" +
    "@media (prefers-reduced-motion:reduce){.totop{transition:none}}";
  document.head.append(css);
  const b = document.createElement("button");
  b.type = "button"; b.className = "totop"; b.setAttribute("aria-label", "Back to top");
  b.innerHTML = '<svg viewBox="0 0 20 20" width="22" height="22" aria-hidden="true"><path d="M10 16V4M4 9l6-6 6 6" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  b.onclick = () => {
    const calm = matchMedia("(prefers-reduced-motion: reduce)").matches;
    window.scrollTo({ top: 0, behavior: calm ? "auto" : "smooth" });
  };
  const upd = () => b.classList.toggle("show", scrollY > 700);
  addEventListener("scroll", upd, { passive: true });
  const add = () => { document.body.append(b); upd(); };
  document.body ? add() : addEventListener("DOMContentLoaded", add);
})();
