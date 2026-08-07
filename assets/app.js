(function () {
  "use strict";

  if (window.lucide) {
    window.lucide.createIcons();
  }

  var revealItems = document.querySelectorAll(".reveal");
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

    revealItems.forEach(function (item) {
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
})();
