/**
 * 【看懂資料與選對圖表】章節導航設定
 */
window.SITE_NAV = {
  title: "📊 看懂資料與選對圖表",
  color: "#2563EB",
  items: [
    { title: "章節大廳", url: "./index.html" },
    { title: "💡 數據提問思考規則", url: "./question-rules.html" },
    { title: "🎯 五大問題與圖表動畫", url: "./question-types.html" },
    { title: "📐 圖表名稱與坐標軸", url: "./chart-vocabulary.html" },
    { title: "📊 資料條件挑圖指南", url: "./chart-matrix.html" },
    { title: "⚡ 欄位組合速查", url: "./field-combinations.html" }
  ],
  prev: { title: "⬅️ 第一部", url: "../excel-generation/index.html", hint: "用 AI 生成 Excel" },
  next: { title: "第三部 ➡️", url: "../ai-examples/index.html", hint: "AI 協作資料分析" },
  github: "https://github.com/roberthsu2003/ai-excel-analytics/tree/main/02_看懂資料與選對圖表"
};

// 載入全站共用導航列（docs/shared/js/site-nav.js）
(function() {
  const script = document.createElement("script");
  script.src = new URL("../../shared/js/site-nav.js", document.currentScript.src).href;
  document.head.appendChild(script);
})();
