/**
 * 全站共用導航列：三部互動專區都用同一個版面。
 * 各章節的 js/nav.js 先設定 window.SITE_NAV，再載入本檔：
 *   { title, color, items: [{ title, url, external? }], prev?: { title, url, hint }, next?, github }
 */
(function() {
  function link(item, isActive, color) {
    const target = item.external ? ' target="_blank" rel="noopener"' : "";
    return `
      <a href="${item.url}"${target} style="
        text-decoration: none;
        font-size: 0.85rem;
        font-weight: 700;
        padding: 6px 8px;
        border-radius: 6px;
        white-space: nowrap;
        transition: all 0.2s ease;
        color: ${isActive ? "#FFFFFF" : "#475569"};
        background: ${isActive ? color : "transparent"};
        border: 1px solid ${isActive ? color : "transparent"};
      ">${item.title}</a>`;
  }

  function chapterLink(item) {
    return `
      <a href="${item.url}" title="${item.hint || ""}" style="
        text-decoration: none;
        font-size: 0.82rem;
        font-weight: 700;
        padding: 5px 8px;
        border-radius: 6px;
        white-space: nowrap;
        color: #334155;
        border: 1px dashed #CBD5E1;
      ">${item.title}</a>`;
  }

  function renderNavbar() {
    const config = window.SITE_NAV;
    if (!config) return;
    const currentPath = window.location.pathname.split("/").pop() || "index.html";
    const chapters = [config.prev, config.next].filter(Boolean);

    const header = document.createElement("header");
    header.style.cssText = "background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(12px); border-bottom: 1px solid #E2E8F0; position: sticky; top: 0; z-index: 100;";
    header.innerHTML = `
      <div style="max-width: 1200px; margin: 0 auto; padding: 12px 16px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
        <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
          <a href="../index.html" title="回到全課程總首頁" style="font-size: 0.85rem; font-weight: 800; color: #0F172A; text-decoration: none; padding: 6px 12px; border-radius: 8px; background: #FFFFFF; border: 1.5px solid #94A3B8; box-shadow: 0 1px 2px rgba(0,0,0,0.06); display: flex; align-items: center; gap: 4px; white-space: nowrap;">
            <span>🏠 全課程總首頁</span>
          </a>
          <span style="color: #CBD5E1;">/</span>
          <a href="./index.html" style="display: flex; align-items: center; gap: 6px; font-weight: 800; font-size: 1.05rem; color: #0F172A; text-decoration: none;">
            <span>${config.title}</span>
          </a>
        </div>
        <nav aria-label="章節導航" style="width: 100%; display: flex; gap: 4px; align-items: center; flex-wrap: wrap;">
          ${config.items.map(item => {
            const isActive = !item.external && (currentPath === item.url.replace("./", "") || (currentPath === "" && item.url === "./index.html"));
            return link(item, isActive, config.color);
          }).join("")}
          ${chapters.length ? `<span style="width: 1px; height: 20px; background: #CBD5E1; margin: 0 4px;"></span>${chapters.map(chapterLink).join("")}` : ""}
          <a href="${config.github}" target="_blank" rel="noopener" style="
            text-decoration: none;
            font-size: 0.85rem;
            font-weight: 600;
            padding: 5px 8px;
            color: #64748B;
            border-left: 1px solid #CBD5E1;
            margin-left: 2px;
            white-space: nowrap;
          ">GitHub 教材 ↗</a>
        </nav>
      </div>
    `;

    const oldHeader = document.querySelector("header");
    if (oldHeader) {
      oldHeader.replaceWith(header);
    } else {
      document.body.insertBefore(header, document.body.firstChild);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", renderNavbar);
  } else {
    renderNavbar();
  }
})();
