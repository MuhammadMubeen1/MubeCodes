(() => {
  const year = document.getElementById("year");
  if (year) year.textContent = String(new Date().getFullYear());

  const nav = document.querySelector(".nav");
  const toggle = document.querySelector(".nav__toggle");
  const mobile = document.querySelector(".nav-mobile");
  const floatCta = document.querySelector(".float-cta");

  const onScroll = () => {
    const y = window.scrollY;
    if (nav) nav.classList.toggle("is-scrolled", y > 24);
    if (floatCta) floatCta.classList.toggle("is-visible", y > 520);
  };
  onScroll();
  window.addEventListener("scroll", onScroll, { passive: true });

  if (toggle && mobile) {
    toggle.addEventListener("click", () => {
      const open = toggle.getAttribute("aria-expanded") === "true";
      toggle.setAttribute("aria-expanded", String(!open));
      mobile.hidden = open;
    });

    mobile.querySelectorAll("a").forEach((link) => {
      link.addEventListener("click", () => {
        toggle.setAttribute("aria-expanded", "false");
        mobile.hidden = true;
      });
    });
  }

  const reveals = document.querySelectorAll(".reveal");
  if ("IntersectionObserver" in window) {
    const io = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("is-visible");
            io.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12, rootMargin: "0px 0px -40px 0px" }
    );
    reveals.forEach((el, i) => {
      el.style.transitionDelay = `${Math.min(i % 6, 5) * 0.06}s`;
      io.observe(el);
    });
  } else {
    reveals.forEach((el) => el.classList.add("is-visible"));
  }

  // Accordion: only one FAQ open at a time (agency UX pattern)
  const faqs = document.querySelectorAll(".faq__item");
  faqs.forEach((item) => {
    item.addEventListener("toggle", () => {
      if (!item.open) return;
      faqs.forEach((other) => {
        if (other !== item) other.open = false;
      });
    });
  });

  // Portfolio 5-Tab Switching Logic with Keyboard Accessibility
  const tabs = Array.from(document.querySelectorAll(".portfolio__tab"));
  const panels = Array.from(document.querySelectorAll(".portfolio__panel"));

  if (tabs.length && panels.length) {
    const activateTab = (selectedTab) => {
      const targetId = selectedTab.getAttribute("aria-controls");

      tabs.forEach((tab) => {
        const isMatch = tab === selectedTab;
        tab.classList.toggle("is-active", isMatch);
        tab.setAttribute("aria-selected", String(isMatch));
        tab.setAttribute("tabindex", isMatch ? "0" : "-1");
      });

      panels.forEach((panel) => {
        const isMatch = panel.id === targetId;
        panel.hidden = !isMatch;
        if (isMatch) {
          // Instantly reveal all elements in the activated panel
          panel.querySelectorAll(".reveal").forEach((el) => {
            el.classList.add("is-visible");
          });
        }
      });
    };

    tabs.forEach((tab, index) => {
      tab.addEventListener("click", () => activateTab(tab));

      // Keyboard arrow navigation
      tab.addEventListener("keydown", (e) => {
        let newIndex = null;
        if (e.key === "ArrowRight") {
          newIndex = (index + 1) % tabs.length;
        } else if (e.key === "ArrowLeft") {
          newIndex = (index - 1 + tabs.length) % tabs.length;
        } else if (e.key === "Home") {
          newIndex = 0;
        } else if (e.key === "End") {
          newIndex = tabs.length - 1;
        }

        if (newIndex !== null) {
          e.preventDefault();
          tabs[newIndex].focus();
          activateTab(tabs[newIndex]);
        }
      });
    });

    // Deep linking: activate matching tab if hash is present in URL
    const activateFromHash = () => {
      const hash = (window.location.hash || "").replace("#", "").toLowerCase();
      if (!hash) return;
      const targetTab = tabs.find((t) => {
        const tabId = t.id.toLowerCase();
        const ctrlId = (t.getAttribute("aria-controls") || "").toLowerCase();
        return tabId === hash || ctrlId === hash || ctrlId === `panel-${hash}` || tabId === `tab-${hash}`;
      });
      if (targetTab) {
        activateTab(targetTab);
      }
    };

    activateFromHash();
    window.addEventListener("hashchange", activateFromHash);
  }

  const form = document.getElementById("contact-form");
  const note = document.getElementById("form-note");
  if (form && note) {
    form.addEventListener("submit", (e) => {
      e.preventDefault();
      const data = new FormData(form);
      const subject = encodeURIComponent(
        `MubeCodes inquiry — ${data.get("need") || "Growth call"}`
      );
      const body = encodeURIComponent(
        [
          `Name: ${data.get("name") || ""}`,
          `Email: ${data.get("email") || ""}`,
          `Business: ${data.get("business") || ""}`,
          `Website: ${data.get("website") || ""}`,
          `Need: ${data.get("need") || ""}`,
          "",
          data.get("message") || "",
        ].join("\n")
      );
      window.location.href = `mailto:mubecodes@gmail.com?subject=${subject}&body=${body}`;
      note.hidden = false;
      form.reset();
    });
  }
})();
