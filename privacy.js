/* PRIVACY POP-UP. Loads before ratings.js. Decides whether covers may be looked up at Apple.
   Nothing is stored or requested until the visitor chooses. */
(function () {
  const THEME = "be";          // "be" = sun & sepia (Common – Be)   |   "donuts" = glaze pink (J Dilla – Donuts)
  const CONTACT = "";          // optional: e.g. "Questions? Mail us at name@example.com"  (shown in the pop-up and the footer)
  const KEY = "privacyChoice.v1";
  let mem = null;
  const read = () => { try { return localStorage.getItem(KEY); } catch (e) { return mem; } };
  const write = v => { mem = v; try { localStorage.setItem(KEY, v); } catch (e) {} };
  const P = window.PRIVACY = { covers: read() === "yes", decided: read() === "yes" || read() === "no", open: null };

  const el = (t, c, h) => { const e = document.createElement(t); if (c) e.className = c; if (h != null) e.innerHTML = h; return e; };
  let box;
  function build() {
    box = el("section", "pv"); box.dataset.theme = THEME; box.hidden = true;
    box.setAttribute("role", "dialog"); box.setAttribute("aria-labelledby", "pv-title"); box.setAttribute("aria-modal", "false");
    const big = THEME === "donuts" ? "Data" : "Be", sub = THEME === "donuts" ? "donuts: what’s in the box?" : "aware of your data";
    box.innerHTML =
      '<div class="pv-art">' + (THEME === "be" ? '<img src="be-cover.jpg" alt="" onerror="this.remove()">' : '') +
      '<svg class="pv-donut" viewBox="0 0 80 80" aria-hidden="true"><circle cx="40" cy="40" r="26" fill="#ff8fb5" stroke="#3a1626" stroke-width="3"/><circle cx="40" cy="40" r="9" fill="#f7b6cc" stroke="#3a1626" stroke-width="3"/>' +
      '<g stroke-width="3.2" stroke-linecap="round"><path d="M22 28l5-3" stroke="#fff"/><path d="M52 22l5 3" stroke="#3a1626"/><path d="M60 45l5-1" stroke="#fff"/><path d="M50 60l3 5" stroke="#3a1626"/><path d="M27 58l-4 3" stroke="#fff"/><path d="M18 44l-1 5" stroke="#3a1626"/></g></svg>' +
      '<h2 id="pv-title"><span class="pv-big"></span><span class="pv-sub"></span></h2></div>' +
      '<div class="pv-body"><p class="pv-lead"><b>We collect no data.</b> No analytics, no cookies.</p>' +
      '<div class="pv-opt"><button class="pv-b pv-yes" type="button" data-v="yes">Yes, show covers</button><p>You see the album covers. They are loaded from Apple or the Cover Art Archive, so they see your IP address and which covers you load. Your browser keeps them for next time, only on your device.</p></div>' +
      '<div class="pv-opt"><button class="pv-b pv-no" type="button" data-v="no">No covers</button><p>You don’t see album covers. Those services do not see your IP address. You get coloured tiles. The rest works the same.</p></div>' +
      '<p class="pv-sp">Spotify links send data only if you click them.</p>' +
      '<p class="pv-note">Change this anytime at the bottom of the site. <details><summary>More detail</summary>' +
      '<span><b>Your device:</b> your choice (<code>privacyChoice.v1</code>), the covers your browser keeps like on any website, and, for brand-new albums, the cover links found (<code>coverCache.v1</code>). No deletes that memory. We never receive any of it.</span>' +
      '<span><b>Covers:</b> most come from Apple. A few come from the Cover Art Archive, which passes you on to the Internet Archive (archive.org). Both see your IP address. We tell them nothing about which page you are on. For albums that are not on our ready-made list yet, your browser may also search Apple by album and artist name.</span>' +
      '<span><b>Spotify:</b> if you click “Listen on Spotify”, your browser opens Spotify in a new tab. Spotify then receives your IP address, browser and device details, and a search for that artist and album. If you are logged in there, it can link the visit to your account. We send no referrer, so Spotify is not told which page you came from. Spotify’s privacy policy applies. We see none of it.</span>' +
      '<span><b>Hosting:</b> this site runs on GitHub Pages, which can log visits whatever you choose here.</span> <span class="pv-contact"></span></details></p></div>';
    box.querySelector(".pv-big").textContent = big; box.querySelector(".pv-sub").textContent = sub;
    if (CONTACT) box.querySelector(".pv-contact").textContent = CONTACT; else box.querySelector(".pv-contact").remove();
    const x = el("button", "pv-x", '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M2 2l8 8M10 2l-8 8" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/></svg>'); x.type = "button"; x.setAttribute("aria-label", "Close without changing anything"); x.hidden = !P.decided;
    box.querySelector(".pv-art").append(x); x.onclick = () => { box.hidden = true; };
    box.addEventListener("click", e => { const v = e.target.closest("[data-v]"); if (v) choose(v.dataset.v === "yes"); });
    box.addEventListener("keydown", e => { if (e.key === "Escape" && P.decided) box.hidden = true; });
    document.body.prepend(box);
  }
  function choose(yes) {
    P.covers = yes; P.decided = true; write(yes ? "yes" : "no");
    if (!yes) { try { localStorage.removeItem("coverCache.v1"); } catch (e) {} }
    box.hidden = true; box.querySelector(".pv-x").hidden = false;
    window.dispatchEvent(new CustomEvent("privacy-change"));
    if (P.reopened && P.footerBtn) P.footerBtn.focus({ preventScroll: true });   // never jump the page
  }
  P.open = () => { P.reopened = true; box.hidden = false; box.querySelector("#pv-title").setAttribute("tabindex", "-1"); box.querySelector("#pv-title").focus({ preventScroll: true }); };

  function init() {
    build();
    const f = document.querySelector("footer");
    if (f) {
      const s = el("span", "pv-small", "");
      const b = el("button", "pv-link", "Privacy &amp; covers"); b.type = "button"; b.onclick = P.open; P.footerBtn = b;
      s.append(b); if (CONTACT) s.append(" · " + CONTACT); f.append(s);
    }
    if (!P.decided) box.hidden = false;
  }
  document.readyState === "loading" ? document.addEventListener("DOMContentLoaded", init) : init();
})();
