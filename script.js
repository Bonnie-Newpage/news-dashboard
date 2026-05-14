// ================================================
// 每日財經新聞自動整理 v6
// ================================================
const SHEET_ID   = “1-Gf6ZlVxMsbb5tGo6-uNUi9DQPAC6hab_tBHSGdFBdg”;
const GEMINI_KEY = “AIzaSyDN3LiPmOrrkHVKNLonCxjI0yo4J9v1s7Y”;

// ================================================
// 主程式
// ================================================
function fetchDailyNews() {
const ss  = SpreadsheetApp.openById(SHEET_ID);
const now = Utilities.formatDate(new Date(), “Asia/Taipei”, “yyyy-MM-dd HH:mm”);
const intlRows = fetchIntlNews(ss, now);
const twRows   = fetchTwNews(ss, now);
generateAndSaveSummary(ss, intlRows, twRows, now);
Logger.log(“全部完成！” + now);
}

// ================================================
// 國際財經新聞
// ================================================
function fetchIntlNews(ss, now) {
const sheet = ss.getSheetByName(“國際新聞”);
sheet.clearContents();
sheet.getRange(1,1,1,5).setValues([[“更新時間”,“標題”,“中文標題”,“來源”,“連結”]]);
sheet.getRange(1,1,1,5).setBackground(”#1a1a2e”).setFontColor(”#22c55e”).setFontWeight(“bold”);

const feeds = [
{ source:“Reuters”,         url:“https://news.google.com/rss/search?q=when:24h+reuters+stock+market&ceid=US:en&hl=en-US&gl=US” },
{ source:“CNBC Finance”,    url:“https://www.cnbc.com/id/10001147/device/rss/rss.html” },
{ source:“Bloomberg Markets”, url:“https://feeds.bloomberg.com/markets/news.rss” },
{ source:“WSJ Markets”,     url:“https://feeds.content.dowjones.io/public/rss/mw_realtimeheadlines” },
{ source:“Financial Times”, url:“https://www.ft.com/rss/home/uk” }
];

const rows = parseFeeds(feeds, now, true);
writeToSheet(sheet, rows);
Logger.log(“國際新聞共 “ + rows.length + “ 則”);
return rows;
}

// ================================================
// 台灣財經新聞
// ================================================
function fetchTwNews(ss, now) {
const sheet = ss.getSheetByName(“台灣新聞”);
sheet.clearContents();
sheet.getRange(1,1,1,5).setValues([[“更新時間”,“標題”,“中文標題”,“來源”,“連結”]]);
sheet.getRange(1,1,1,5).setBackground(”#1a1a2e”).setFontColor(”#f97316”).setFontWeight(“bold”);

const feeds = [
{ source:“Yahoo奇摩財經”,  url:“https://tw.stock.yahoo.com/rss?category=market-news” },
{ source:“經濟日報”,        url:“https://money.udn.com/rssfeed/news/1001/5590?ch=money” },
{ source:“工商時報”,        url:“https://news.google.com/rss/search?q=%E5%B7%A5%E5%95%86%E6%99%82%E5%A0%B1+%E8%B2%A1%E7%B6%93+%E5%8F%B0%E8%82%A1&hl=zh-TW&gl=TW&ceid=TW:zh-Hant” },
{ source:“鉅亨網”,          url:“https://news.google.com/rss/search?q=%E9%89%85%E4%BA%A8%E7%B6%B2+%E5%8F%B0%E8%82%A1+%E8%B2%A1%E7%B6%93&hl=zh-TW&gl=TW&ceid=TW:zh-Hant” },
{ source:“MoneyDJ”,        url:“https://news.google.com/rss/search?q=MoneyDJ+%E5%8F%B0%E8%82%A1&hl=zh-TW&gl=TW&ceid=TW:zh-Hant” },
{ source:“台灣財經綜合”,    url:“https://news.google.com/rss/search?q=%E5%8F%B0%E8%82%A1+%E8%B2%A1%E7%B6%93+%E4%BB%8A%E6%97%A5&hl=zh-TW&gl=TW&ceid=TW:zh-Hant” }
];

const rows = parseFeeds(feeds, now, false);
writeToSheet(sheet, rows);
Logger.log(“台灣新聞共 “ + rows.length + “ 則”);
return rows;
}

// ================================================
// 生成 AI 摘要並存入試算表
// ================================================
function generateAndSaveSummary(ss, intlRows, twRows, now) {
var sheet = ss.getSheetByName(“每日摘要”);
if (!sheet) sheet = ss.insertSheet(“每日摘要”);
sheet.clearContents();
sheet.getRange(1,1,1,3).setValues([[“更新時間”,“日期”,“摘要內容”]]);
sheet.getRange(1,1,1,3).setBackground(”#1a1a2e”).setFontColor(”#a78bfa”).setFontWeight(“bold”);

var intlTitles = intlRows.map(function(r, i) {
return (i+1) + “. [” + r[3] + “] “ + (r[2] || r[1]);
}).join(”\n”);

var twTitles = twRows.map(function(r, i) {
return (i+1) + “. [” + r[3] + “] “ + r[1];
}).join(”\n”);

var today = Utilities.formatDate(new Date(), “Asia/Taipei”, “yyyy-MM-dd”);
var prompt = today + “ 財經新聞，請用繁體中文寫300字市場分析摘要，直接輸出內容：\n\n國際：\n” + intlTitles + “\n\n台灣：\n” + twTitles;

var models = [“gemini-2.0-flash”, “gemini-2.0-flash-lite”, “gemini-1.5-flash-001”];

for (var m = 0; m < models.length; m++) {
var model = models[m];
try {
var url = “https://generativelanguage.googleapis.com/v1/models/” + model + “:generateContent?key=” + GEMINI_KEY;
var payload = {
contents: [{ parts: [{ text: prompt }] }],
generationConfig: { maxOutputTokens: 600, temperature: 0.7 }
};
var resp = UrlFetchApp.fetch(url, {
method: “post”,
contentType: “application/json”,
payload: JSON.stringify(payload),
muteHttpExceptions: true
});

```
  var json = JSON.parse(resp.getContentText());

  if (json.error) {
    Logger.log("模型 " + model + " 錯誤：" + json.error.code + " " + json.error.message);
    continue;
  }

  var summary = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0] ? json.candidates[0].content.parts[0].text : "";

  if (summary) {
    sheet.getRange(2,1,1,3).setValues([[now, today, summary]]);
    Logger.log("摘要生成成功（" + model + "），共 " + summary.length + " 字");
    return;
  }
} catch(e) {
  Logger.log("模型 " + model + " 例外：" + e.message);
}
```

}

Logger.log(“所有模型都失敗”);
sheet.getRange(2,1,1,3).setValues([[now, today, “今日摘要暫時無法生成，請稍後重試。”]]);
}

// ================================================
// 共用：解析 RSS
// ================================================
function parseFeeds(feeds, now, needTranslate) {
var rows = [];
for (var f = 0; f < feeds.length; f++) {
var feed = feeds[f];
try {
var resp = UrlFetchApp.fetch(feed.url, {
muteHttpExceptions: true,
headers: { “User-Agent”: “Mozilla/5.0” }
});
var xml  = XmlService.parse(resp.getContentText());
var root = xml.getRootElement();
var items = [];

```
  if (root.getName() === "rss") {
    var ch = root.getChild("channel");
    if (ch) items = ch.getChildren("item");
  } else if (root.getName() === "feed") {
    var ns = XmlService.getNamespace("http://www.w3.org/2005/Atom");
    items = root.getChildren("entry", ns);
    if (!items || items.length === 0) items = root.getChildren("entry");
  } else {
    var ch2 = root.getChild("channel");
    if (ch2) items = ch2.getChildren("item");
  }

  if (items.length > 2) items = items.slice(0, 2);

  for (var i = 0; i < items.length; i++) {
    var item  = items[i];
    var title = item.getChildText("title") || "";
    var link  = item.getChildText("link")  || "";
    if (!link) {
      var el = item.getChild("link");
      if (el) link = el.getAttribute("href") ? el.getAttribute("href").getValue() : "";
    }
    if (!title) continue;

    var zhTitle = "";
    if (needTranslate) {
      try {
        zhTitle = LanguageApp.translate(title, "en", "zh-TW");
        Utilities.sleep(200);
      } catch(te) { zhTitle = ""; }
    } else {
      zhTitle = title;
    }
    rows.push([now, title, zhTitle, feed.source, link]);
  }
  Logger.log(feed.source + "：成功 " + items.length + " 則");
} catch(e) {
  Logger.log(feed.source + " 錯誤：" + e.message);
}
```

}
return rows;
}

// ================================================
// 共用：寫入試算表
// ================================================
function writeToSheet(sheet, rows) {
if (rows.length === 0) return;
sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

// ================================================
// 每天 8:00 自動觸發（只需執行一次）
// ================================================
function setDailyTrigger() {
var triggers = ScriptApp.getProjectTriggers();
for (var i = 0; i < triggers.length; i++) {
ScriptApp.deleteTrigger(triggers[i]);
}
ScriptApp.newTrigger(“fetchDailyNews”)
.timeBased().everyDays(1).atHour(8)
.inTimezone(“Asia/Taipei”).create();
Logger.log(“已設定每天早上 8:00 自動執行！”);
}

// ================================================
// Web App API
// ================================================
function doGet(e) {
var sheetName = (e && e.parameter && e.parameter.sheet) ? e.parameter.sheet : “國際新聞”;
var ss  = SpreadsheetApp.openById(SHEET_ID);
var tab = ss.getSheetByName(sheetName);

if (!tab) {
return ContentService
.createTextOutput(JSON.stringify({ error: “找不到工作表” }))
.setMimeType(ContentService.MimeType.JSON);
}

var data    = tab.getDataRange().getValues();
var headers = data[0];
var rows    = [];

for (var i = 1; i < data.length; i++) {
var obj = {};
for (var j = 0; j < headers.length; j++) {
obj[headers[j]] = String(data[i][j] || “”);
}
var hasTitle   = obj[“標題”]   && obj[“標題”]   !== “標題”;
var hasSummary = obj[“摘要內容”] && obj[“摘要內容”] !== “摘要內容”;
if (hasTitle || hasSummary) rows.push(obj);
}

return ContentService
.createTextOutput(JSON.stringify({ ok: true, data: rows }))
.setMimeType(ContentService.MimeType.JSON);
}
