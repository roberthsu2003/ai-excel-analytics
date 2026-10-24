/**
 * 【用 AI 生成 Excel】章節導航設定
 */
window.SITE_NAV = {
  title: "📗 用 AI 生成 Excel",
  color: "#059669",
  items: [
    { title: "章節大廳", url: "./index.html" },
    { title: "📘 Excel 名詞圖鑑", url: "./excel-terms.html" },
    { title: "🔗 公式與參照動畫", url: "./formula-reference.html" },
    { title: "📐 規範書流程", url: "./spec-workflow.html" },
    { title: "🛡️ 防呆與警示", url: "./validation-format.html" }
  ],
  next: { title: "第二部 ➡️", url: "../chart-selection/index.html", hint: "看懂資料與選對圖表" },
  github: "https://github.com/roberthsu2003/ai-excel-analytics/tree/main/01_用AI生成Excel"
};

// 載入全站共用導航列（docs/shared/js/site-nav.js）
(function() {
  const script = document.createElement("script");
  script.src = new URL("../../shared/js/site-nav.js", document.currentScript.src).href;
  document.head.appendChild(script);
})();
