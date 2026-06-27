const scriptProps = PropertiesService.getScriptProperties();

const CONFIG = {
 SPREADSHEET_ID: scriptProps.getProperty("SPREADSHEET_ID") || "",
 GEMINI_API_KEY: scriptProps.getProperty("GEMINI_API_KEY") || "",
 GCP_PROJECT_ID: scriptProps.getProperty("GCP_PROJECT_ID") || "",
 VERTEX_REGION: scriptProps.getProperty("VERTEX_REGION") || "",
 VERTEX_MODEL: scriptProps.getProperty("VERTEX_MODEL") || "",
 SHARED_PASSWORD: scriptProps.getProperty("SHARED_PASSWORD") || "kopi@2026",
 
 // Mengambil daftar anggota tim secara dinamis dari Properti Skrip
 get TEAM_MEMBERS() {
  try {
   const raw = scriptProps.getProperty("TEAM_MEMBERS");
   if (raw) return JSON.parse(raw);
  } catch(err) {
   Logger.log("Error parsing TEAM_MEMBERS property: " + err.message);
  }
  // Data contoh (sebagai fallback jika properti belum disetel)
  return [
   { email: "john@example.com", name: "John", team: "Customer Service" },
   { email: "brad@example.com", name: "Brad", team: "Operator" }
  ];
 },

 // Mengambil email yang diizinkan langsung dari daftar anggota tim
 get ALLOWED_EMAILS() {
  return this.TEAM_MEMBERS.map(m => m.email.toLowerCase().trim());
 },

 TG_BOT_TOKEN: scriptProps.getProperty("TG_BOT_TOKEN") || "",
 TG_CHAT_ID: scriptProps.getProperty("TG_CHAT_ID") || "",
 MANAGER_EMAIL: scriptProps.getProperty("MANAGER_EMAIL") || "john@example.com",
 REPORT_HOUR_WIB: parseInt(scriptProps.getProperty("REPORT_HOUR_WIB")) || 16,
};

const SHEETS = { CHECKIN: "CheckIn", ANALYSIS: "AI_Analysis", TG_LOG: "TG_Log" };

function R(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
 try { const data = JSON.parse(e.postData.contents); saveCheckin(data); return R({ status:"ok", message:"Check-in tersimpan!" }); } 
 catch(err) { return R({ status:"error", message:err.message }); }
}

// ✅ FUNGSI UTAMA: MENANGANI API DAN MENYERVIS HTML UNTUK GITHUB
function doGet(e) {
 const action = (e && e.parameter && e.parameter.action) || "";
 
 // Jika diminta HTML oleh GitHub Pages
 if (action === "getHtml") {
  return HtmlService.createHtmlOutputFromFile('index')
   .setTitle('AI Daily Check-in')
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
 }
 
 // Jika ada action API lainnya
 if (action !== "") {
  try {
   switch(action) {
    case "ping": return R({ status:"ok", pong:true, ts:new Date().toISOString() });
    case "getToday": return R(getTodayCheckins());
    case "getAll": return R(getAllCheckins(parseInt(e.parameter.days)||30));
    case "getStats": return R(getStats());
    case "getLastAnalysis": return R(getLastAnalysis());
    case "getLastTelegram": return R(getLastTelegram());
    case "getMissing": return R(getMissingToday());
    case "login": return R(handleLogin(e.parameter.email, e.parameter.pass));
    case "getMembers": return R(getMembers());
    case "save": return R(saveFromGet(e));
    default: return R({ status:"error", message:"Unknown action" });
   }
  } catch(err) { return R({ status:"error", message:err.message }); }
 }

 // Jika langsung buka URL Web App (bukan dari GitHub)
 return HtmlService.createHtmlOutputFromFile('index')
  .setTitle('AI Daily Check-in — Operation & Service Squad')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() { if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); return SpreadsheetApp.getActiveSpreadsheet(); }

function normalizeDate(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "string" && /^\d{4}-\d{2}-\d{2}$/.test(val.trim())) return val.trim();
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM-dd");
  if (typeof val === "string") {
   const slashMatch = val.trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
   if (slashMatch) { const d = new Date(parseInt(slashMatch[3]), parseInt(slashMatch[2])-1, parseInt(slashMatch[1])); if (!isNaN(d.getTime())) return Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd"); }
   const parsed = new Date(val.trim()); if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, "Asia/Jakarta", "yyyy-MM-dd");
  }
  if (typeof val === "number") { const d = new Date((val - 25569) * 86400 * 1000); if (!isNaN(d.getTime())) return Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd"); }
  return String(val).trim();
}

function getSheet(name, headers, headerColor) {
 const ss = getSpreadsheet(); let sh = ss.getSheetByName(name);
 if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground(headerColor||"#1a237e").setFontColor("#fff"); sh.setFrozenRows(1); }
 return sh;
}

function sheetRows(name) {
 const ss = getSpreadsheet(); const sh = ss.getSheetByName(name);
 if (!sh || sh.getLastRow() < 2) return [];
 const vals = sh.getDataRange().getValues(); const hdrs = vals[0].map(h => String(h).trim());
 return vals.slice(1).map(row => { const o = {}; hdrs.forEach((h,i) => o[h] = row[i]); return o; });
}

function todayWIB() { return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd"); }

function handleLogin(email, pass) {
 if (!email || !pass) return { valid:false, message:"Email dan password wajib diisi." };
 const emailNorm = email.toLowerCase().trim();
 if (!CONFIG.ALLOWED_EMAILS.map(e => e.toLowerCase().trim()).includes(emailNorm)) return { valid:false, message:"Email tidak terdaftar dalam tim." };
 if (pass !== CONFIG.SHARED_PASSWORD) return { valid:false, message:"Password salah." };
 const today = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
 return { valid:true, token:Utilities.base64Encode(emailNorm + "|" + today + "|ops"), email:emailNorm };
}

function getMembers() { return { members: CONFIG.TEAM_MEMBERS.map(m => ({ name:m.name, team:m.team })) }; }

function saveCheckin(data) {
  const hdrs = ["ID", "Tanggal", "Waktu", "Nama", "Tim", "AI Tool", "Prompt", "Kategori", "Streak", "Login Email"];
  const sh = getSheet(SHEETS.CHECKIN, hdrs, "#1a237e");
  const now = new Date();
  sh.appendRow([
   data.id || now.getTime(),
   data.date || Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd"),
   data.time || Utilities.formatDate(now, "Asia/Jakarta", "HH:mm"),
   data.name || "", data.team || "",
   Array.isArray(data.tools) ? data.tools.join(", ") : (data.tools || ""),
   data.prompt || "", data.useCase || "-", data.streak || 1, data.by || ""
  ]);
  const lastRow = sh.getLastRow();
  sh.getRange(lastRow, 2).setNumberFormat("@");
  sh.getRange(lastRow, 3).setNumberFormat("@");
}

function saveFromGet(e) { try { saveCheckin(JSON.parse(e.parameter.data)); return { status:"ok", message:"Check-in tersimpan ke Sheets!" }; } catch(err) { return { status:"error", message:err.message }; } }
function getTodayCheckins() { const td = todayWIB(); return sheetRows(SHEETS.CHECKIN).filter(r => normalizeDate(r["Tanggal"]) === td); }
function getAllCheckins(days) { const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - days); const cs = Utilities.formatDate(cutoff,"Asia/Jakarta","yyyy-MM-dd"); return sheetRows(SHEETS.CHECKIN).filter(r => normalizeDate(r["Tanggal"]) >= cs); }

function getStats() {
  const all = sheetRows(SHEETS.CHECKIN); const td = todayWIB(); const month = td.slice(0, 7);
  const todayRows = all.filter(r => normalizeDate(r["Tanggal"]) === td);
  const monthRows = all.filter(r => normalizeDate(r["Tanggal"]).startsWith(month));
  const counts = {}; monthRows.forEach(r => { const n = r["Nama"]; if (n) counts[n] = (counts[n] || 0) + 1; });
  const leaderboard = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([name, count]) => ({ name, count }));
  const teamBreakdown = {}; todayRows.forEach(r => { const t = r["Tim"]; if (t) teamBreakdown[t] = (teamBreakdown[t] || 0) + 1; });
  return { today: todayRows.length, total: all.length, totalMembers: CONFIG.TEAM_MEMBERS.length, thisMonth: monthRows.length, leaderboard, teamBreakdown, checkedInToday: todayRows.map(r => r["Nama"]).filter(Boolean) };
}

function getMissingToday() {
 const td = getTodayCheckins(); const chk = new Set(td.map(c => String(c["Nama"]||"").toLowerCase().trim()));
 const missing = CONFIG.TEAM_MEMBERS.filter(m => !chk.has(m.name.toLowerCase().trim()));
 const byTeam = {}; missing.forEach(m => { (byTeam[m.team]=byTeam[m.team]||[]).push(m.name); });
 return { missing, byTeam, count:missing.length };
}

function getLastAnalysis() { const rows = sheetRows(SHEETS.ANALYSIS); if (!rows.length) return { status:"empty" }; const last = rows[rows.length-1]; return { status:"ok", date:last["Tanggal"]||"", time:last["Jam"]||"", checkinCount:last["Check-in"]||0, missingCount:last["Missing"]||0, text:last["Analisa"]||"" }; }
function getLastTelegram() { const rows = sheetRows(SHEETS.TG_LOG); if (!rows.length) return { status:"empty", text:"" }; const last = rows[rows.length-1]; return { status:"ok", date:last["Tanggal"]||"", time:last["Jam"]||"", text:last["Pesan"]||"", sent:last["Status"]||"" }; }

function callVertexGemini(prompt) {
 const endpoint = "https://" + CONFIG.VERTEX_REGION + "-aiplatform.googleapis.com/v1/projects/" + CONFIG.GCP_PROJECT_ID + "/locations/" + CONFIG.VERTEX_REGION + "/publishers/google/models/" + CONFIG.VERTEX_MODEL + ":generateContent?key=" + CONFIG.GEMINI_API_KEY;
 const payload = { contents: [{ role:"user", parts:[{ text: prompt }] }], generationConfig: { temperature: 1.0, maxOutputTokens: 8192, topP: 0.7 }, safetySettings: [{ category:"HARM_CATEGORY_HARASSMENT", threshold:"BLOCK_MEDIUM_AND_ABOVE" }, { category:"HARM_CATEGORY_HATE_SPEECH", threshold:"BLOCK_MEDIUM_AND_ABOVE" }, { category:"HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold:"BLOCK_MEDIUM_AND_ABOVE" }, { category:"HARM_CATEGORY_DANGEROUS_CONTENT", threshold:"BLOCK_MEDIUM_AND_ABOVE" }] };
 const response = UrlFetchApp.fetch(endpoint, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
 const code = response.getResponseCode(); const result = JSON.parse(response.getContentText());
 if (code !== 200) throw new Error("Vertex API error " + code + ": " + JSON.stringify(result.error||result));
 const candidates = result.candidates;
 if (!candidates || !candidates[0] || !candidates[0].content) { const reason = (candidates && candidates[0] && candidates[0].finishReason) || "unknown"; throw new Error("Response kosong dari Vertex AI. Finish reason: " + reason); }
 return candidates[0].content.parts.map(function(p){ return p.text||""; }).join("");
}

function runDailyReport() {
 const date = todayWIB(); const td = getTodayCheckins(); const missing = getMissingToday(); const total = CONFIG.TEAM_MEMBERS.length; const pct = Math.round(td.length / total * 100);
 const teamStats = {}; CONFIG.TEAM_MEMBERS.forEach(m => { if (!teamStats[m.team]) teamStats[m.team] = { total:0, done:0 }; teamStats[m.team].total++; }); td.forEach(c => { if (teamStats[c["Tim"]]) teamStats[c["Tim"]].done++; });
 const teamSummary = Object.entries(teamStats).map(([t,d]) => `• ${t}: ${d.done}/${d.total} (${Math.round(d.done/d.total*100)}%)`).join("\n");
 const missingText = Object.keys(missing.byTeam).length ? Object.entries(missing.byTeam).map(([t,ns]) => `• ${t}: ${ns.join(", ")}`).join("\n") : "✅ Semua anggota sudah check-in";
 const promptsText = td.length ? td.map(c => `[${c["Tim"]}] ${c["Nama"]}: "${c["Prompt"]}" (${c["Kategori"]}, ${c["AI Tool"]})`).join("\n") : "(tidak ada data check-in)";
 const geminiPrompt = `Kamu adalah AI analyst untuk tim Operation & Service Squad — tim operasional yang sedang membangun budaya penggunaan AI dalam pekerjaan sehari-hari.\n\nLAPORAN HARIAN — ${date}\nTotal anggota: ${total} orang\nCheck-in: ${td.length} orang (${pct}%)\n\nPARTISIPASI PER TIM:\n${teamSummary}\n\nANGGOTA BELUM CHECK-IN (${missing.count} orang):\n${missingText}\n\nPROMPT YANG DIGUNAKAN (${td.length} entri):\n${promptsText}\n\n---\nTulis laporan analisa dalam Bahasa Indonesia. Format WAJIB:\n\n## 📊 Ringkasan\n2-3 kalimat: total check-in, persentase, highlight hari ini.\n\n## 🏆 Top 3 Prompt Terbaik\nPilih 3 prompt paling berguna. Format tiap item:\n**[Nama] — [Tim]**\nPrompt: "[tuliskan prompt-nya]"\nKenapa efektif: [1 kalimat]\n\n## 💡 Insight Tren\n2 insight singkat berdasarkan pola penggunaan AI hari ini.\n\n## 🎯 Action Items Besok\n3 langkah konkret untuk Leader. Gunakan bullet pendek, langsung actionable.\n\nNada: profesional, ringkas, tidak basa-basi. Maksimal 2500 kata.`;

 let analysisText = ""; let vertexOk = false;
 try { analysisText = callVertexGemini(geminiPrompt); vertexOk = true; } catch(err) { analysisText = `⚠️ Vertex AI error: ${err.message}\n\nData check-in tetap tersimpan di Sheets.`; }
 const now = Utilities.formatDate(new Date(), "Asia/Jakarta", "HH:mm");
 const aSh = getSheet(SHEETS.ANALYSIS, ["Tanggal","Jam","Check-in","Missing","Model","Analisa"], "#1b5e20");
 aSh.appendRow([date, now, td.length, missing.count, vertexOk ? "Gemini " + CONFIG.VERTEX_MODEL : "error", analysisText]); aSh.setColumnWidth(6, 600);
 const tgMessage = buildTelegramMessage(date, td.length, total, pct, teamStats, missing, analysisText);
 const tgResult = sendTelegram(tgMessage);
 const tSh = getSheet(SHEETS.TG_LOG, ["Tanggal","Jam","Pesan","Status","Chat ID"], "#0d47a1");
 tSh.appendRow([date, now, tgMessage, tgResult ? "✅ Terkirim" : "❌ Gagal", CONFIG.TG_CHAT_ID]); tSh.setColumnWidth(3, 400);
 Logger.log(`📊 Daily Report selesai: ${date} | ${td.length}/${total} check-in | TG: ${tgResult?"OK":"FAIL"}`);
}

function buildTelegramMessage(date, ciCount, total, pct, teamStats, missing, analysis) {
 const bar = pct >= 80 ? "🟢" : pct >= 50 ? "🟡" : "🔴";
 const teamLines = Object.entries(teamStats).map(([t,d]) => { const p = Math.round(d.done/d.total*100); const icon = p===100?"✅":p>=50?"🔶":"❌"; return `${icon} *${t}*: ${d.done}/${d.total}`; }).join("\n");
 const missLines = missing.count > 0 ? Object.entries(missing.byTeam).map(([t,ns]) => `• ${t}: ${ns.join(", ")}`).join("\n") : "✅ Semua sudah check-in\\!";
 const shortAnalysis = extractShortAnalysis(analysis);
 const msg = `🤖 *AI Daily Check\\-in Report*\n📅 ${escTG(date)} \\| ${bar} ${pct}% Partisipasi\n\n━━━━━━━━━━━━━━━━━━\n📊 *Check\\-in: ${ciCount}/${total} orang*\n\n${teamLines}\n\n━━━━━━━━━━━━━━━━━━\n⚠️ *Belum Check\\-in:*\n${escTG(missLines)}\n\n━━━━━━━━━━━━━━━━━━\n${escTG(shortAnalysis)}\n\n━━━━━━━━━━━━━━━━━━\n_Dianalisa oleh AI_\n_Operation & Service Squad_`;
 return msg;
}

function escTG(text) { return String(text).replace(/[_*[\]()~`>#+\-=|{}.!\\]/g, '\\$&'); }
function extractShortAnalysis(text) {
 if (!text || text.startsWith("⚠️")) return text; const sections = [];
 const promptMatch = text.match(/## 🏆 Top 3 Prompt Terbaik([\s\S]*?)(?=## |$)/);
 if (promptMatch) { const lines = promptMatch[1].trim().split("\n").slice(0,5); sections.push("🏆 *Prompt Terbaik Hari Ini*\n" + lines.join("\n")); }
 const actionMatch = text.match(/## 🎯 Action Items([\s\S]*?)(?=## |$)/);
 if (actionMatch) { sections.push("🎯 *Action Items Leader*\n" + actionMatch[1].trim()); }
 return sections.join("\n\n") || text.slice(0,800);
}

function sendTelegram(message) {
 try { const url = `https://api.telegram.org/bot${CONFIG.TG_BOT_TOKEN}/sendMessage`; const res = UrlFetchApp.fetch(url, { method: "post", headers: { "Content-Type":"application/json" }, payload: JSON.stringify({ chat_id:CONFIG.TG_CHAT_ID, text:message, parse_mode:"MarkdownV2" }), muteHttpExceptions: true }); const result = JSON.parse(res.getContentText()); if (!result.ok) { if (result.description && result.description.includes("parse")) return sendTelegramPlain(stripMarkdown(message)); return false; } return true; } catch(err) { return false; }
}
function sendTelegramPlain(message) { try { const url = `https://api.telegram.org/bot${CONFIG.TG_BOT_TOKEN}/sendMessage`; const res = UrlFetchApp.fetch(url, { method: "post", headers: {"Content-Type":"application/json"}, payload: JSON.stringify({ chat_id:CONFIG.TG_CHAT_ID, text:message }), muteHttpExceptions: true }); return JSON.parse(res.getContentText()).ok; } catch { return false; } }
function stripMarkdown(text) { return text.replace(/[*_`\[\]()~>#+=|{}.!\\]/g,''); }

function setupDailyTrigger() {
 ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "runDailyReport") ScriptApp.deleteTrigger(t); });
 const hourUTC = CONFIG.REPORT_HOUR_WIB - 7;
 ScriptApp.newTrigger("runDailyReport").timeBased().everyDays(1).atHour(hourUTC < 0 ? hourUTC + 24 : hourUTC).create();
 Logger.log(`✅ Trigger aktif: runDailyReport setiap hari jam ${CONFIG.REPORT_HOUR_WIB}:00 WIB`);
}