(function () {
  "use strict";

  if (window.lucide) {
    window.lucide.createIcons();
  }

  var revealItems = Array.prototype.slice.call(document.querySelectorAll(".reveal"));
  var metricsSection = document.querySelector("#metrics");

  function animateCounter(el) {
    var target = Number(el.getAttribute("data-count")) || 0;
    var prefix = el.getAttribute("data-prefix") || "";
    var suffix = el.getAttribute("data-suffix") || "";
    var duration = 900;
    var start = null;

    function step(timestamp) {
      if (!start) {
        start = timestamp;
      }
      var progress = Math.min((timestamp - start) / duration, 1);
      var eased = 1 - Math.pow(1 - progress, 3);
      var value = Math.round(target * eased);
      el.textContent = prefix + value + suffix;
      if (progress < 1) {
        requestAnimationFrame(step);
      }
    }

    requestAnimationFrame(step);
  }

  function initScrollProgress() {
    var bar = document.querySelector("#progress-bar");
    if (!bar) {
      return;
    }
    function update() {
      var scrollTop = window.scrollY || document.documentElement.scrollTop;
      var height = document.documentElement.scrollHeight - window.innerHeight;
      var progress = height > 0 ? Math.min((scrollTop / height) * 100, 100) : 0;
      bar.style.width = progress + "%";
    }
    window.addEventListener("scroll", update, { passive: true });
    update();
  }

  function initConsole() {
    var lines = Array.prototype.slice.call(document.querySelectorAll(".console-body p"));
    if (!lines.length) {
      return;
    }
    lines.forEach(function (line, index) {
      line.style.opacity = "0";
      line.style.transition = "opacity 0.45s ease";
      line.style.transitionDelay = index * 450 + "ms";
      window.setTimeout(function () {
        line.style.opacity = "1";
      }, 250 + index * 450);
    });
  }

  function initCopyLink() {
    var button = document.querySelector("[data-copy-link]");
    if (!button) {
      return;
    }
    button.addEventListener("click", function () {
      var url = window.location.href;
      var label = button.querySelector("span");
      function done() {
        if (label) {
          var original = label.textContent;
          label.textContent = "已复制";
          window.setTimeout(function () {
            label.textContent = original;
          }, 1600);
        }
      }
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(url).then(done, function () {
          fallbackCopy(url, done);
        });
      } else {
        fallbackCopy(url, done);
      }
    });
  }

  function fallbackCopy(text, done) {
    var area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.select();
    try {
      document.execCommand("copy");
    } catch (e) {
      /* ignore */
    }
    document.body.removeChild(area);
    done();
  }

  if ("IntersectionObserver" in window) {
    var revealObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add("in-view");
            revealObserver.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12 }
    );

    revealItems.forEach(function (item, index) {
      item.style.setProperty("--delay", (index % 8) * 60 + "ms");
      revealObserver.observe(item);
    });

    var counterStarted = false;
    var counterObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting && !counterStarted) {
            counterStarted = true;
            document.querySelectorAll(".metric-value[data-count]").forEach(animateCounter);
            counterObserver.disconnect();
          }
        });
      },
      { threshold: 0.3 }
    );

    if (metricsSection) {
      counterObserver.observe(metricsSection);
    }
  } else {
    revealItems.forEach(function (item) {
      item.classList.add("in-view");
    });
    document.querySelectorAll(".metric-value[data-count]").forEach(function (el) {
      var target = el.getAttribute("data-count");
      var prefix = el.getAttribute("data-prefix") || "";
      var suffix = el.getAttribute("data-suffix") || "";
      el.textContent = prefix + target + suffix;
    });
  }

  var printButton = document.querySelector("[data-print]");
  if (printButton) {
    printButton.addEventListener("click", function () {
      window.print();
    });
  }

  initScrollProgress();
  initConsole();
  initCopyLink();
})();
