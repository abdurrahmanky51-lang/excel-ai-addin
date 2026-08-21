/* Excel AI Asistanı - taskpane.js (v2 - Puter.js sürümü)
   Office.js ile Excel'den veri okur/yazar.
   AI için Puter.js kullanır: API anahtarı gerekmez, geliştirici için tamamen
   ücretsiz ve sınırsızdır (Puter'ın "User-Pays" modeli). İlk kullanımda
   kullanıcıya ücretsiz, tek seferlik bir Puter girişi sorulabilir. */

const MODEL = "claude-sonnet-5";
const INFO_SEEN_KEY = "excelAiAssistant_infoSeen";

Office.onReady(() => {
  document.getElementById("app").classList.remove("hidden");
  initTabs();
  initInfoPanel();
  bindActions();
});

/* ---------- Bilgi paneli ---------- */

function initInfoPanel() {
  const btn = document.getElementById("infoBtn");
  const panel = document.getElementById("infoPanel");

  if (!localStorage.getItem(INFO_SEEN_KEY)) {
    panel.classList.remove("hidden");
  }

  btn.addEventListener("click", () => panel.classList.toggle("hidden"));
  document.getElementById("closeInfoBtn").addEventListener("click", () => {
    panel.classList.add("hidden");
    localStorage.setItem(INFO_SEEN_KEY, "1");
  });
}

/* ---------- Sekmeler ---------- */

function initTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
    });
  });
}

/* ---------- Yükleniyor göstergesi ---------- */

function setLoading(on) {
  document.getElementById("loading").classList.toggle("hidden", !on);
  document.querySelectorAll(".primary-btn").forEach((b) => (b.disabled = on));
}

/* ---------- Excel'den veri okuma ---------- */

async function readSelection() {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "values", "formulas", "rowCount", "columnCount"]);
    await context.sync();
    return {
      address: range.address,
      values: range.values,
      formulas: range.formulas,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
    };
  });
}

function rangeToText(sel, maxRows = 25) {
  const rows = sel.values.slice(0, maxRows);
  const formulas = sel.formulas.slice(0, maxRows);
  let out = `Aralık: ${sel.address} (${sel.rowCount} satır x ${sel.columnCount} sütun)\n`;
  out += "Değerler:\n" + rows.map((r) => r.join(" | ")).join("\n");
  const hasFormulas = formulas.some((r) => r.some((c) => typeof c === "string" && c.startsWith("=")));
  if (hasFormulas) {
    out += "\n\nFormüller:\n" + formulas.map((r) => r.join(" | ")).join("\n");
  }
  return out;
}

/* ---------- Puter.js üzerinden AI çağrısı ---------- */

async function callAI(systemPrompt, userPrompt) {
  if (typeof puter === "undefined" || !puter.ai) {
    throw new Error("AI servisi yüklenemedi. İnternet bağlantını kontrol edip tekrar dene.");
  }
  const messages = [
    { role: "system", content: systemPrompt },
    { role: "user", content: userPrompt },
  ];
  const response = await puter.ai.chat(messages, { model: MODEL });

  // Puter yanıt şekli genelde response.message.content[0].text şeklindedir,
  // bazı durumlarda düz metin/obje de dönebilir; hepsini karşılıyoruz.
  let text = "";
  if (response?.message?.content?.[0]?.text) {
    text = response.message.content[0].text;
  } else if (typeof response === "string") {
    text = response;
  } else if (response?.toString && response.toString() !== "[object Object]") {
    text = response.toString();
  }

  if (!text) throw new Error("AI yanıtı alınamadı, lütfen tekrar dene.");
  return text;
}

function showResult(elId, html) {
  const el = document.getElementById(elId);
  el.innerHTML = html;
  el.classList.remove("hidden");
}

function showError(elId, message) {
  showResult(elId, `<div class="status error">${escapeHtml(message)}</div>`);
}

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

/* Basit satır içi kod bloğu tespiti: ```...``` veya `=FORMUL(...)` yakalar */
function extractFormula(text) {
  const codeBlock = text.match(/```(?:excel)?\n?([\s\S]*?)```/);
  if (codeBlock) return codeBlock[1].trim().split("\n")[0].trim();
  const inline = text.match(/=[A-ZÇĞİÖŞÜ][A-ZÇĞİÖŞÜ0-9._]*\([^\n]*\)/i);
  if (inline) return inline[0].trim();
  return null;
}

function renderWithFormula(text) {
  const formula = extractFormula(text);
  let html = `<div>${escapeHtml(text)}</div>`;
  if (formula) {
    html += `<span class="formula-box">${escapeHtml(formula)}</span>`;
    html += `<button class="apply-btn" data-formula="${escapeHtml(formula)}">Seçili hücreye uygula</button>`;
  }
  return html;
}

function bindApplyButtons(containerId) {
  const container = document.getElementById(containerId);
  container.querySelectorAll(".apply-btn").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const formula = btn.dataset.formula;
      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.formulas = [[formula]];
          await context.sync();
        });
        btn.textContent = "Uygulandı ✓";
        btn.disabled = true;
      } catch (e) {
        alert("Uygulanamadı: " + e.message);
      }
    });
  });
}

/* ---------- Eylemler ---------- */

function bindActions() {
  document.getElementById("formulaBtn").addEventListener("click", onFormulaClick);
  document.getElementById("fixBtn").addEventListener("click", onFixClick);
  document.getElementById("tableBtn").addEventListener("click", onTableClick);
  document.getElementById("chartBtn").addEventListener("click", onChartClick);
  document.getElementById("askBtn").addEventListener("click", onAskClick);
}

async function onFormulaClick() {
  const goal = document.getElementById("formulaGoal").value.trim();
  if (!goal) return;
  setLoading(true);
  try {
    let context = "";
    try {
      const sel = await readSelection();
      context = "\n\nSeçili hücrelerdeki mevcut veri:\n" + rangeToText(sel, 10);
    } catch (e) { /* seçim yoksa devam et */ }

    const system =
      "Sen profesyonel bir Excel formül uzmanısın. Kullanıcının hedefine uygun TEK bir Excel formülü öner. " +
      "Cevabını kısa tut: önce formülü ```...``` kod bloğu içinde ver, sonra 1-2 cümlelik Türkçe açıklama yap. " +
      "Türkçe Excel fonksiyon isimlerini kullan (örn: TOPLA.ÇARPIM, DÜŞEYARA, EĞERSAY).";
    const text = await callAI(system, `Hedef: ${goal}${context}`);
    showResult("formulaResult", renderWithFormula(text));
    bindApplyButtons("formulaResult");
  } catch (e) {
    showError("formulaResult", e.message);
  } finally {
    setLoading(false);
  }
}

async function onFixClick() {
  setLoading(true);
  try {
    const sel = await readSelection();
    const system =
      "Sen profesyonel bir Excel hata ayıklama uzmanısın. Verilen aralıktaki formülleri ve değerleri incele. " +
      "Hata varsa (#REF!, #VALUE!, #DIV/0!, mantık hatası, yanlış referans vb.) belirt ve düzeltilmiş formülü " +
      "```...``` kod bloğu içinde ver. Hata yoksa kısaca 'Hata bulunamadı' de ve varsa iyileştirme önerisi sun.";
    const text = await callAI(system, rangeToText(sel));
    showResult("fixResult", renderWithFormula(text));
    bindApplyButtons("fixResult");
  } catch (e) {
    showError("fixResult", e.message);
  } finally {
    setLoading(false);
  }
}

async function onTableClick() {
  setLoading(true);
  try {
    const sel = await readSelection();
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.add(range, true);
      table.style = "TableStyleMedium2";
      await context.sync();
    });

    const system =
      "Sen profesyonel bir Excel tablo/veri tasarım uzmanısın. Verilen veriye bakarak kısaca: " +
      "1) Bu veri için hangi ek sütunlar (hesaplanan alanlar) faydalı olur, " +
      "2) Hangi koşullu biçimlendirme kuralları anlamlı olur, kısaca öner. Kısa ve maddeler halinde yaz.";
    const text = await callAI(system, rangeToText(sel));
    showResult("tableResult", `<div class="status success">Tablo oluşturuldu.</div><div>${escapeHtml(text)}</div>`);
  } catch (e) {
    showError("tableResult", e.message);
  } finally {
    setLoading(false);
  }
}

async function onChartClick() {
  setLoading(true);
  try {
    const sel = await readSelection();
    const system =
      "Sen profesyonel bir veri görselleştirme uzmanısın. Verilen veriye en uygun grafik türünü seç. " +
      "SADECE şu seçeneklerden birini birinci satırda tek kelime olarak yaz: ColumnClustered, Line, Pie, Bar, Area, ScatterMarkers. " +
      "İkinci satırdan itibaren 1-2 cümle Türkçe gerekçe yaz.";
    const text = await callAI(system, rangeToText(sel));
    const firstLine = text.split("\n")[0].trim();
    const validTypes = ["ColumnClustered", "Line", "Pie", "Bar", "Area", "ScatterMarkers"];
    const chartType = validTypes.find((t) => firstLine.includes(t)) || "ColumnClustered";

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.add(chartType, range, "Auto");
      chart.title.text = "Otomatik Oluşturulan Grafik";
      await context.sync();
    });

    showResult("chartResult", `<div class="status success">"${chartType}" türünde grafik oluşturuldu.</div><div>${escapeHtml(text)}</div>`);
  } catch (e) {
    showError("chartResult", e.message);
  } finally {
    setLoading(false);
  }
}

async function onAskClick() {
  const question = document.getElementById("askInput").value.trim();
  if (!question) return;
  setLoading(true);
  try {
    let context = "";
    try {
      const sel = await readSelection();
      context = "\n\nBağlam (seçili hücreler):\n" + rangeToText(sel, 10);
    } catch (e) { /* seçim yoksa devam et */ }

    const system =
      "Sen profesyonel bir Excel danışmanısın. Kullanıcının sorusuna net, kısa ve uygulanabilir bir cevap ver. " +
      "Formül önerirken Türkçe fonksiyon isimlerini kullan ve formülü ```...``` bloğu içinde ver.";
    const text = await callAI(system, `${question}${context}`);
    showResult("askResult", renderWithFormula(text));
    bindApplyButtons("askResult");
  } catch (e) {
    showError("askResult", e.message);
  } finally {
    setLoading(false);
  }
}
