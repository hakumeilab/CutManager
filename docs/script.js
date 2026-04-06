const revealElements = document.querySelectorAll(".reveal");
const releaseLinks = document.querySelectorAll("[data-release-link]");
const releaseVersion = document.querySelector("[data-release-version]");
const releaseVersionHeading = document.querySelector("[data-release-version-heading]");
const releaseDate = document.querySelector("[data-release-date]");
const releaseSummary = document.querySelector("[data-release-summary]");
const releaseNotes = document.querySelector("[data-release-notes]");

const releaseMetadataUrl = "./release.json";
const releasesPageUrl = "https://github.com/hakumeilab/CutManager/releases/latest";

const showElement = (element) => {
  element.classList.add("is-visible");
};

const normalizeVersion = (value) => String(value || "").trim().replace(/^[^0-9]+/, "");

const formatDate = (value) => {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";

  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
};

const renderReleaseNotes = (notes) => {
  if (!releaseNotes) return;

  if (!notes.length) {
    releaseNotes.textContent = "最新の更新内容は GitHub Releases で確認できます。";
    return;
  }

  const list = document.createElement("ul");
  notes.forEach((note) => {
    const item = document.createElement("li");
    item.textContent = note;
    list.appendChild(item);
  });

  releaseNotes.replaceChildren(list);
};

const updateReleaseUi = async () => {
  try {
    const response = await fetch(releaseMetadataUrl, { cache: "no-store" });

    if (!response.ok) throw new Error(`HTTP ${response.status}`);

    const release = await response.json();
    const version = normalizeVersion(release.tag_name || release.name || "");
    const publishedDate = formatDate(release.published_at);
    const downloadUrl = release.download_url || release.html_url || releasesPageUrl;
    const notes = Array.isArray(release.notes) ? release.notes : [];

    if (version && releaseVersion) {
      releaseVersion.textContent = `最新版 v${version}`;
    }

    if (version && releaseVersionHeading) {
      releaseVersionHeading.textContent = `v${version}`;
    }

    if (releaseDate) {
      releaseDate.textContent = publishedDate ? `公開日 ${publishedDate}` : "";
    }

    if (releaseSummary) {
      const parts = [];
      if (version) parts.push(`最新版は v${version}`);
      if (publishedDate) parts.push(`公開日 ${publishedDate}`);
      releaseSummary.textContent = parts.length > 0
        ? `${parts.join(" / ")}。`
        : "GitHub Releases から最新版を取得できます。";
    }

    renderReleaseNotes(notes);

    releaseLinks.forEach((link) => {
      link.setAttribute("href", downloadUrl);
      if (version) {
        link.setAttribute("aria-label", `CutManager v${version} をダウンロード`);
      }
    });
  } catch {
    if (releaseVersion) {
      releaseVersion.textContent = "最新版は GitHub Releases で確認";
    }

    if (releaseVersionHeading) {
      releaseVersionHeading.textContent = "GitHub Releases";
    }

    if (releaseDate) {
      releaseDate.textContent = "";
    }

    if (releaseSummary) {
      releaseSummary.textContent = "GitHub Releases から最新版を取得できます。";
    }

    renderReleaseNotes([]);

    releaseLinks.forEach((link) => link.setAttribute("href", releasesPageUrl));
  }
};

if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
  revealElements.forEach(showElement);
} else {
  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        showElement(entry.target);
        observer.unobserve(entry.target);
      });
    },
    {
      threshold: 0.14,
      rootMargin: "0px 0px -8% 0px",
    }
  );

  revealElements.forEach((element) => observer.observe(element));
}

updateReleaseUi();
