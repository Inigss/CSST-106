const menuToggle = document.querySelector(".menu-toggle");
const sidebar = document.querySelector("#sidebar");
const currentPage = document.querySelector("#current-page");
const dashboardLink = document.querySelector(
  '.primary-nav > .nav-link[href="#dashboard"]',
);
const documentsLink = document.querySelector(
  '.primary-nav > .nav-link[href="#lab-documents"]',
);

menuToggle.addEventListener("click", () => {
  const isOpen = sidebar.classList.toggle("is-open");
  menuToggle.setAttribute("aria-expanded", String(isOpen));
  menuToggle.setAttribute(
    "aria-label",
    isOpen ? "Close navigation menu" : "Open navigation menu",
  );
});

function updateCurrentPage() {
  const isDocumentSection = window.location.hash === "#lab-documents";
  currentPage.textContent = isDocumentSection ? "Lab handouts" : "Dashboard";

  dashboardLink.classList.toggle("is-active", !isDocumentSection);
  documentsLink.classList.toggle("is-active", isDocumentSection);

  if (isDocumentSection) {
    documentsLink.setAttribute("aria-current", "page");
    dashboardLink.removeAttribute("aria-current");
  } else {
    dashboardLink.setAttribute("aria-current", "page");
    documentsLink.removeAttribute("aria-current");
  }
}

sidebar.addEventListener("click", (event) => {
  if (!event.target.closest(".nav-link")) {
    return;
  }

  sidebar.classList.remove("is-open");
  menuToggle.setAttribute("aria-expanded", "false");
  menuToggle.setAttribute("aria-label", "Open navigation menu");
});

window.addEventListener("hashchange", updateCurrentPage);
updateCurrentPage();
