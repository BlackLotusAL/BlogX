const DEFAULT_WORKBOOK = "./data/articles.xlsx";
const IGNORED_SHEETS = new Set(["说明", "字段说明", "README", "Readme"]);

const state = {
  workbook: null,
  months: [],
  activeMonth: "",
};

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  wireInteractions();
  loadWorkbookFromUrl(DEFAULT_WORKBOOK);
});

function cacheElements() {
  els.monthNav = document.querySelector("#monthNav");
  els.countText = document.querySelector("#countText");
  els.articleList = document.querySelector("#articleList");
  els.emptyState = document.querySelector("#emptyState");
}

function wireInteractions() {
  window.addEventListener("hashchange", () => {
    const month = decodeURIComponent(window.location.hash.slice(1));
    if (state.months.some((item) => item.name === month)) {
      state.activeMonth = month;
      renderAll();
    }
  });
}

async function loadWorkbookFromUrl(url) {
  try {
    const response = await fetch(url, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`${response.status} ${response.statusText}`);
    }
    const buffer = await response.arrayBuffer();
    openWorkbook(buffer);
  } catch (error) {
    renderError(`默认数据加载失败：${error.message}`);
  }
}

function openWorkbook(buffer) {
  state.workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    dense: false,
  });

  state.months = state.workbook.SheetNames.map((name) => ({
    name,
    rows: getRowsFromSheet(name),
  })).filter((sheet) => !IGNORED_SHEETS.has(sheet.name) && sheet.rows.length > 0);

  if (state.months.length === 0) {
    state.activeMonth = "";
    renderEmpty("Excel 中没有找到文章记录");
    return;
  }

  const hashMonth = decodeURIComponent(window.location.hash.slice(1));
  state.activeMonth = state.months.some((item) => item.name === hashMonth)
    ? hashMonth
    : state.months[0].name;

  renderAll();
}

function getRowsFromSheet(sheetName) {
  const sheet = state.workbook.Sheets[sheetName];
  if (!sheet || !sheet["!ref"]) return [];

  const rawRows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false,
  });

  return rawRows
    .map(normalizeArticle)
    .filter((article) => article.title && article.visible)
    .sort(sortArticles);
}

function normalizeArticle(row) {
  const title = readCell(row, ["标题", "Title"]);
  const tags = splitList(readCell(row, ["标签", "Tags"]));
  const tagLinks = parseTagLinks(tags, readCell(row, ["标签链接", "标签网址", "Tag Links"]));

  return {
    title,
    summary: readCell(row, ["摘要", "简介", "Summary"]),
    category: readCell(row, ["分类", "栏目", "Category"]),
    date: normalizeDate(readCell(row, ["发布日期", "日期", "Date"])),
    readingTime: normalizeReadingTime(readCell(row, ["阅读时间", "阅读分钟", "Reading Time"])),
    views: normalizeViews(readCell(row, ["浏览量", "阅读量", "Views"])),
    url: normalizeUrl(readCell(row, ["文章链接", "链接", "URL", "Article URL"])),
    author: readCell(row, ["作者", "Author"]),
    order: toNumber(readCell(row, ["排序", "Order"])),
    visible: isVisible(readCell(row, ["是否显示", "显示", "Visible"])),
    tags: tags.map((label, index) => ({
      label,
      url: tagLinks.byLabel.get(label) || tagLinks.byIndex.get(index) || "",
    })),
  };
}

function readCell(row, acceptedKeys) {
  for (const [key, value] of Object.entries(row)) {
    if (acceptedKeys.includes(String(key).trim())) {
      return cleanText(value);
    }
  }
  return "";
}

function cleanText(value) {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function splitList(value) {
  return cleanText(value)
    .split(/[;；|,，\n]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseTagLinks(tags, value) {
  const byLabel = new Map();
  const byIndex = new Map();
  const segments = cleanText(value)
    .split(/[;；|\n]+/)
    .map((item) => item.trim())
    .filter(Boolean);
  const unkeyed = [];

  for (const segment of segments) {
    const eqIndex = segment.indexOf("=");
    if (eqIndex > 0) {
      const label = segment.slice(0, eqIndex).trim();
      const url = normalizeUrl(segment.slice(eqIndex + 1));
      if (label && url) byLabel.set(label, url);
    } else {
      unkeyed.push(normalizeUrl(segment));
    }
  }

  unkeyed.forEach((url, index) => {
    if (url) byIndex.set(index, url);
  });

  tags.forEach((tag, index) => {
    if (!byLabel.has(tag) && byIndex.has(index)) {
      byLabel.set(tag, byIndex.get(index));
    }
  });

  return { byLabel, byIndex };
}

function normalizeUrl(value) {
  const url = cleanText(value);
  if (!url) return "";
  if (/^[a-z][a-z0-9+.-]*:/i.test(url) || url.startsWith("/") || url.startsWith("#")) {
    return url;
  }
  if (/^www\./i.test(url)) {
    return `https://${url}`;
  }
  return url;
}

function normalizeDate(value) {
  const text = cleanText(value);
  if (!text) return "";

  const excelSerial = Number(text);
  if (Number.isFinite(excelSerial) && excelSerial > 20000 && excelSerial < 80000) {
    const date = XLSX.SSF.parse_date_code(excelSerial);
    if (date) {
      return `${date.y}-${pad(date.m)}-${pad(date.d)}`;
    }
  }

  return text.replace(/\//g, "-");
}

function normalizeReadingTime(value) {
  const text = cleanText(value);
  if (!text) return "";
  return /^\d+(\.\d+)?$/.test(text) ? `${text} 分钟` : text;
}

function normalizeViews(value) {
  const text = cleanText(value);
  if (!text) return "";
  const numericText = text.replace(/,/g, "");
  if (!/^\d+(\.\d+)?$/.test(numericText)) return text;
  const count = Number(numericText);
  if (!Number.isFinite(count)) return text;
  return `${count.toLocaleString("zh-CN")} 浏览`;
}

function toNumber(value) {
  const number = Number(cleanText(value));
  return Number.isFinite(number) ? number : null;
}

function isVisible(value) {
  const text = cleanText(value).toLowerCase();
  return !["否", "no", "false", "0", "隐藏", "hide"].includes(text);
}

function sortArticles(a, b) {
  if (a.order !== null || b.order !== null) {
    return (a.order ?? 9999) - (b.order ?? 9999);
  }
  return dateValue(b.date) - dateValue(a.date);
}

function dateValue(value) {
  const timestamp = Date.parse(value);
  return Number.isFinite(timestamp) ? timestamp : 0;
}

function renderAll() {
  renderMonthNav();
  renderArticles();
}

function renderMonthNav() {
  els.monthNav.textContent = "";

  for (const month of state.months) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "month-tab";
    button.textContent = month.name;
    button.setAttribute("aria-current", String(month.name === state.activeMonth));
    button.addEventListener("click", () => {
      state.activeMonth = month.name;
      history.replaceState(null, "", `#${encodeURIComponent(month.name)}`);
      renderAll();
    });
    els.monthNav.append(button);
  }
}

function renderArticles() {
  const current = state.months.find((item) => item.name === state.activeMonth);
  const articles = current?.rows ?? [];

  els.articleList.textContent = "";
  els.emptyState.hidden = articles.length > 0;

  if (!current) {
    els.countText.innerHTML = "未选择月份";
    return;
  }

  els.countText.innerHTML = `<strong>${current.name}</strong> · 共 <strong>${articles.length}</strong> 篇文章`;

  if (articles.length === 0) {
    els.emptyState.hidden = false;
    return;
  }

  for (const article of articles) {
    els.articleList.append(createArticleCard(article));
  }
}

function createArticleCard(article) {
  const card = document.createElement("article");
  card.className = "article-card";
  card.dataset.clickable = String(Boolean(article.url));

  if (article.url) {
    card.tabIndex = 0;
    card.setAttribute("role", "link");
    card.setAttribute("aria-label", `打开文章：${article.title}`);
    card.addEventListener("click", (event) => {
      if (event.target.closest("a, button")) return;
      window.open(article.url, "_blank", "noopener,noreferrer");
    });
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        window.open(article.url, "_blank", "noopener,noreferrer");
      }
    });
  }

  const head = document.createElement("div");
  head.className = "article-head";

  const meta = document.createElement("div");
  meta.className = "article-meta";
  if (article.category) {
    const category = document.createElement("span");
    category.className = "category-badge";
    category.textContent = article.category;
    meta.append(category);
  }
  if (article.date) {
    meta.append(createMetaItem("calendar", article.date));
  }
  if (article.readingTime) {
    meta.append(createMetaItem("clock", article.readingTime));
  }
  if (article.views) {
    meta.append(createMetaItem("eye", article.views));
  }
  if (article.author) {
    meta.append(createMetaItem("user", article.author));
  }
  head.append(meta);

  const title = document.createElement("h2");
  title.className = "article-title";
  if (article.url) {
    const link = document.createElement("a");
    link.href = article.url;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.textContent = article.title;
    title.append(link);
  } else {
    title.textContent = article.title;
  }
  head.append(title);
  card.append(head);

  if (article.summary) {
    const summary = document.createElement("p");
    summary.className = "article-summary";
    summary.textContent = article.summary;
    card.append(summary);
  }

  if (article.tags.length > 0) {
    const tagRow = document.createElement("div");
    tagRow.className = "tag-row";
    tagRow.append(icon("tag", "tag-prefix"));

    for (const tag of article.tags) {
      if (tag.url) {
        const tagLink = document.createElement("a");
        tagLink.className = "tag-link";
        tagLink.href = tag.url;
        tagLink.target = "_blank";
        tagLink.rel = "noopener noreferrer";
        tagLink.textContent = tag.label;
        tagRow.append(tagLink);
      } else {
        const pill = document.createElement("span");
        pill.className = "tag-pill";
        pill.textContent = tag.label;
        tagRow.append(pill);
      }
    }
    card.append(tagRow);
  }

  if (article.url) {
    const external = document.createElement("a");
    external.className = "external-link";
    external.href = article.url;
    external.target = "_blank";
    external.rel = "noopener noreferrer";
    external.setAttribute("aria-label", "打开文章");
    external.append(icon("external", "external-icon"));
    card.append(external);
  }

  return card;
}

function createMetaItem(type, text) {
  const item = document.createElement("span");
  item.className = "meta-item";
  item.append(icon(type, "meta-icon"), document.createTextNode(text));
  return item;
}

function icon(type, className) {
  const span = document.createElement("span");
  span.className = className;
  span.setAttribute("aria-hidden", "true");
  const paths = {
    calendar:
      '<path d="M8 2v4M16 2v4M3 10h18"/><rect x="3" y="4" width="18" height="18" rx="2"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    eye: '<path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
    tag: '<path d="M20 13 11 22l-8-8V4h10l7 7z"/><path d="M7.5 7.5h.01"/>',
    external: '<path d="M14 3h7v7"/><path d="M10 14 21 3"/><path d="M21 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5"/>',
    user: '<path d="M20 21a8 8 0 0 0-16 0"/><circle cx="12" cy="7" r="4"/>',
  };
  span.innerHTML = `<svg viewBox="0 0 24 24">${paths[type] ?? ""}</svg>`;
  return span;
}

function renderEmpty(message) {
  els.monthNav.textContent = "";
  els.articleList.textContent = "";
  els.countText.textContent = message;
  els.emptyState.hidden = false;
}

function renderError(message) {
  els.monthNav.textContent = "";
  els.articleList.textContent = "";
  els.countText.innerHTML = `<span class="error-text">${message}</span>`;
  els.emptyState.hidden = false;
}

function pad(value) {
  return String(value).padStart(2, "0");
}
