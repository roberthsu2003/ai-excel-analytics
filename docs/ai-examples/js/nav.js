/**
 * 【AI 協作資料分析】章節導航設定
 */
window.SITE_NAV = {
  title: "🤖 AI 協作資料分析",
  color: "#2563EB",
  items: [
    { title: "章節大廳", url: "./index.html" },
    { title: "🧯 常見問題排解 ↗", url: "https://github.com/roberthsu2003/ai-excel-analytics/blob/main/03_AI%E5%8D%94%E4%BD%9C%E8%B3%87%E6%96%99%E5%88%86%E6%9E%90/%E5%B8%B8%E8%A6%8B%E5%95%8F%E9%A1%8C%E6%8E%92%E8%A7%A3.md", external: true }
  ],
  prev: { title: "⬅️ 第二部", url: "../chart-selection/index.html", hint: "看懂資料與選對圖表" },
  github: "https://github.com/roberthsu2003/ai-excel-analytics/tree/main/03_AI協作資料分析"
};

// 載入全站共用導航列（docs/shared/js/site-nav.js）
(function() {
  const script = document.createElement("script");
  script.src = new URL("../../shared/js/site-nav.js", document.currentScript.src).href;
  document.head.appendChild(script);
})();
