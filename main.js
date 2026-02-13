const dropZone = document.getElementById("dropZone");
const appRoot = document.querySelector("main.app");
const fileInput = document.getElementById("fileInput");
const statusText = document.getElementById("statusText");
const resultList = document.getElementById("resultList");
const resultTemplate = document.getElementById("resultTemplate");
const downloadAllBtn = document.getElementById("downloadAllBtn");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const clearAllBtn = document.getElementById("clearAllBtn");
const reloadBtn = document.getElementById("reloadBtn");
const compactMode = document.getElementById("compactMode");
const groupByMonth = document.getElementById("groupByMonth");
const langJaBtn = document.getElementById("langJaBtn");
const langEnBtn = document.getElementById("langEnBtn");
const heroTitle = document.getElementById("heroTitle");
const heroLead = document.getElementById("heroLead");
const securityBadgeText = document.getElementById("securityBadgeText");
const dropTitle = document.getElementById("dropTitle");
const dropSub = document.getElementById("dropSub");
const flowTitle = document.getElementById("flowTitle");
const flowStep1Label = document.getElementById("flowStep1Label");
const flowStep1Sub = document.getElementById("flowStep1Sub");
const flowStep2Label = document.getElementById("flowStep2Label");
const flowStep2Sub = document.getElementById("flowStep2Sub");
const flowStep3Label = document.getElementById("flowStep3Label");
const flowStep3Sub = document.getElementById("flowStep3Sub");
const compactLabel = document.getElementById("compactLabel");
const groupLabel = document.getElementById("groupLabel");
const noticeTitle = document.getElementById("noticeTitle");
const noticeSection = document.getElementById("noticeSection");
const noticeLead = document.getElementById("noticeLead");
const noticeList = document.getElementById("noticeList");
const officialNote = document.getElementById("officialNote");
const contactLink = document.getElementById("contactLink");
const privacyLink = document.getElementById("privacyLink");
const termsLink = document.getElementById("termsLink");
const updatesSection = document.getElementById("updatesSection");
const updatesTitle = document.getElementById("updatesTitle");
const updatesList = document.getElementById("updatesList");
const flowSection = document.getElementById("flowSection");
const step2Panel = document.getElementById("step2Panel");
const guidesSection = document.getElementById("guidesSection");
const guidesTitle = document.getElementById("guidesTitle");
const guideExplainerLink = document.getElementById("guideExplainerLink");
const guideExplainerTag = document.getElementById("guideExplainerTag");
const guideExplainerTitle = document.getElementById("guideExplainerTitle");
const guideExplainerSub = document.getElementById("guideExplainerSub");
const guideComparisonLink = document.getElementById("guideComparisonLink");
const guideComparisonTag = document.getElementById("guideComparisonTag");
const guideComparisonTitle = document.getElementById("guideComparisonTitle");
const guideComparisonSub = document.getElementById("guideComparisonSub");
const guideHowtoLink = document.getElementById("guideHowtoLink");
const guideHowtoTag = document.getElementById("guideHowtoTag");
const guideHowtoTitle = document.getElementById("guideHowtoTitle");
const guideHowtoSub = document.getElementById("guideHowtoSub");

const batchItems = [];
let updatesData = [];
const LANG_STORE_KEY = "apple_receipt_lang";
const BUILD_SEEN_KEY = "sctol_last_seen_build";
let currentLang = localStorage.getItem(LANG_STORE_KEY) || "ja";

const I18N = {
  ja: {
    title: "メールで送られてきたAppleの領収書スクショに自動でタイトルをつけるツール",
    hero: "ジョークっぽいのに実用的。タイトルつけるのがめんどくせえスクショ領収書を同じ命名規則で一括整理できます。",
    securityBadge: "画像は外部送信されません",
    dropTitle: "ここにスクリーンショット/PDFをドロップ",
    dropSub: "またはファイル選択から追加",
    dropAria: "スクリーンショットまたはPDFをドロップ",
    flowTitle: "3ステップで完了",
    flowStep1Label: "画像を選択",
    flowStep1Sub: "ドラッグ&ドロップ",
    flowStep2Label: "ブラウザ内で処理",
    flowStep2Sub: "外部送信なし",
    flowStep3Label: "リネーム完了",
    flowStep3Sub: "ZIPでまとめて保存",
    compactLabel: "Apple領収書プリセット（会社名 + 日付 + 合計 + サービス名）",
    groupLabel: "ZIPで月別フォルダ分け",
    downloadAll: "まとめて保存 (ZIP)",
    exportCsv: "CSV出力",
    clearAll: "全件クリア",
    reload: "再読み込み",
    retry: "再解析",
    noticeTitle: "このサイトの取り扱いと免責",
    noticeAria: "説明と免責",
    noticeLead: "このツールはブラウザ上で画像を処理し、ファイル名候補を作る用途です。サーバーへ画像をアップロードして保存する機能はありません。",
    noticeItems: [
      "画像と抽出テキストはブラウザ内でのみ処理します。",
      "このサイト側で画像ファイルやOCR結果を保管・収集する仕組みはありません。",
      "OCRライブラリはCDNから読み込みます（通信は発生）。画像外部送信を目的とした実装はしていません。",
      "抽出結果の正確性は保証できません。金額・日付・品目は必ず原本で確認してください。",
      "Appleの請求書スクリーンショット以外も読み取れる場合がありますが、精度は高くないため必ず原本で確認してください。",
      "本ツール利用によって生じた損害・不利益について、開発者は責任を負いません。",
    ],
    official: "本サービスはApple社とは独立して運営しており、Apple社の提供・認定・保証を受けたものではありません。",
    contact: "お問い合わせ",
    privacy: "プライバシーポリシー",
    terms: "利用規約",
    updatesTitle: "お知らせ",
    updatesAria: "お知らせ",
    guidesTitle: "ゆるい活用ガイド",
    guidesAria: "活用ガイド",
    guideExplainerTag: "解説",
    guideExplainerTitle: "Appleの領収書スクショ、まさか手打ちでタイトルリネームしてないですよね？",
    guideExplainerSub: "もちろん私はしてます。",
    guideComparisonTag: "比較",
    guideComparisonTitle: "おい…他に方法はなかったのか？",
    guideComparisonSub: "みんな同じところで詰まってる。",
    guideHowtoTag: "使い方",
    guideHowtoTitle: "使い方（というほどのものでもない）",
    guideHowtoSub: "以上。",
    statusReady: "準備完了",
    statusPresetOn: "Apple領収書プリセット: ON",
    statusPresetOff: "Apple領収書プリセット: OFF",
    statusNeedImage: "画像またはPDFファイルをドロップしてください",
    statusPdfConverting: (name, page, total) => `PDF変換中: ${name} (${page}/${total})`,
    statusPdfFail: (name) => `PDF変換失敗: ${name}`,
    statusProcessing: (n) => `${n}件を処理中...`,
    statusAnalyzing: (name) => `解析中: ${name}`,
    statusDone: (name) => `完了: ${name}`,
    statusOcrSkip: (name) => `OCRスキップ: ${name}`,
    statusNeedFirst: "先に画像を追加してください",
    statusZipLibErr: "ZIPライブラリの読み込みに失敗しました",
    statusZipping: "ZIPを作成中...",
    statusZipDone: (name) => `保存完了: ${name}`,
    statusZipFail: "ZIP作成に失敗しました",
    statusCsvDone: (name) => `CSV保存完了: ${name}`,
    statusCsvFail: "CSV出力に失敗しました",
    statusCleared: "一覧をクリアしました",
    statusRetrying: (name) => `再解析中: ${name}`,
    statusRetryDone: (name) => `再解析完了: ${name}`,
    statusBuildChanged: (from, to) => `更新があります（${from} -> ${to}）。再読み込みを推奨します。`,
    statusNeedPresetOff: (name) => `非Apple領収書を検出: ${name}（AppleプリセットをOFFにしてください）`,
    statusReviewNeeded: (name) => `要確認: ${name}（抽出信頼度が低いためZIPから自動除外）`,
    sceneFixed: "判定: Apple領収書固定",
    sceneGeneral: "判定: 汎用",
    sceneOcrFail: "判定: OCR失敗",
    sceneAnalyzing: "判定: 解析中",
    sceneInvalidInput: "判定: 入力画像エラー",
    sceneNonApplePreset: "判定: 非Apple（プリセット対象外）",
    sceneReviewNeeded: "判定: 要確認（低信頼）",
    invalidTitlePlaceholder: "入力画像エラー（保存不可）",
    nonAppleTitlePlaceholder: "非Apple（プリセットOFFで保存可）",
    ocrPrefix: "OCR:",
    noText: "テキストを検出できませんでした",
    ocrFailMsg: "OCR失敗。ネット接続またはCDNブロックを確認してください。",
    ocrUiCaptureMsg: "この画像はツール画面のスクショです。原本の領収書スクショを選んでください。",
    ocrNonApplePresetMsg: "Apple領収書プリセットONでは非Apple領収書は自動命名しません。プリセットをOFFにして保存してください。",
    ocrReviewMsg: (score, reasons) => `抽出信頼度 ${score}%。${reasons ? `要確認理由: ${reasons}` : "原本と照合してください。"}`,
    reviewReasonDate: "日付抽出が不安定",
    reviewReasonAmount: "合計抽出が不安定",
    reviewReasonService: "サービス名抽出が不安定",
    reviewReasonCount: "明細件数の根拠が弱い",
    reviewReasonConflict: "件数とサービス名が矛盾",
    reviewReasonMixed: "ストア種別が混在",
    statusNeedOriginal: (name) => `非対応画像を検出: ${name}（原本の領収書スクショを指定してください）`,
    include: "ZIPに含める",
    duplicate: "重複候補（自動除外）",
    date: "日付",
    amount: "合計",
    itemCount: "明細件数",
    service: "サービス名 / アプリ名",
    filename: "最終ファイル名",
    thumbTip: "左の原本プレビューを見ながら右の項目を確認できます。",
    loupeOn: "虫眼鏡: ON",
    loupeOff: "虫眼鏡: OFF",
    downloadSingle: "この名前で単体保存",
    originalPrefix: "元ファイル:",
    savedPrefix: "保存済み:",
    analyzing: "解析中",
  },
  en: {
    title: "Auto-title your Apple receipt screenshots",
    hero: "A joke-ish tool that's actually useful. Batch-organize Apple receipt screenshots with one naming rule.",
    securityBadge: "Your images never leave this browser",
    dropTitle: "Drop screenshots or PDFs here",
    dropSub: "or add files from file picker",
    dropAria: "Drop screenshots or PDFs",
    flowTitle: "Done in 3 steps",
    flowStep1Label: "Select files",
    flowStep1Sub: "Drag & drop",
    flowStep2Label: "Process in browser",
    flowStep2Sub: "No external upload",
    flowStep3Label: "Renamed",
    flowStep3Sub: "Save as ZIP",
    compactLabel: "Apple receipt preset (Company + Date + Total + Service)",
    groupLabel: "Group by month folders in ZIP",
    downloadAll: "Download All (ZIP)",
    exportCsv: "Export CSV",
    clearAll: "Clear All",
    reload: "Reload",
    retry: "Reanalyze",
    noticeTitle: "Handling & Disclaimer",
    noticeAria: "Handling and disclaimer",
    noticeLead: "This tool processes images in your browser and generates filename candidates. It does not upload and store your images on this site.",
    noticeItems: [
      "Images and extracted text are processed inside your browser only.",
      "This site has no mechanism to store or collect uploaded images or OCR output.",
      "OCR libraries are loaded from CDN (network access occurs), but this implementation is not intended to send your images externally.",
      "Extraction accuracy is not guaranteed. Always verify amounts, dates, and service names against the original receipt.",
      "The tool may still parse non-Apple receipt screenshots, but accuracy is lower, so always verify against the original.",
      "The developer is not liable for any loss or damage caused by using this tool.",
    ],
    official: "This service is operated independently from Apple and is not provided, certified, or guaranteed by Apple.",
    contact: "Contact",
    privacy: "Privacy Policy",
    terms: "Terms",
    updatesTitle: "Updates",
    updatesAria: "Updates",
    guidesTitle: "A Chill Practical Guide",
    guidesAria: "Practical guides",
    guideExplainerTag: "Explainer",
    guideExplainerTitle: "You’re not still manually renaming Apple receipt screenshots, are you?",
    guideExplainerSub: "Yeah, I was.",
    guideComparisonTag: "Comparison",
    guideComparisonTitle: "Wait... was there really no better way?",
    guideComparisonSub: "Turns out everyone gets stuck in the same place.",
    guideHowtoTag: "How-To",
    guideHowtoTitle: "How to use it (if this even counts as a how-to)",
    guideHowtoSub: "That’s it.",
    statusReady: "Ready",
    statusPresetOn: "Apple receipt preset: ON",
    statusPresetOff: "Apple receipt preset: OFF",
    statusNeedImage: "Please drop image or PDF files",
    statusPdfConverting: (name, page, total) => `Converting PDF: ${name} (${page}/${total})`,
    statusPdfFail: (name) => `Failed to convert PDF: ${name}`,
    statusProcessing: (n) => `Processing ${n} file(s)...`,
    statusAnalyzing: (name) => `Analyzing: ${name}`,
    statusDone: (name) => `Done: ${name}`,
    statusOcrSkip: (name) => `OCR skipped: ${name}`,
    statusNeedFirst: "Add images first",
    statusZipLibErr: "Failed to load ZIP library",
    statusZipping: "Creating ZIP...",
    statusZipDone: (name) => `Saved: ${name}`,
    statusZipFail: "Failed to create ZIP",
    statusCsvDone: (name) => `CSV saved: ${name}`,
    statusCsvFail: "Failed to export CSV",
    statusCleared: "Cleared all items",
    statusRetrying: (name) => `Reanalyzing: ${name}`,
    statusRetryDone: (name) => `Reanalysis done: ${name}`,
    statusBuildChanged: (from, to) => `Update available (${from} -> ${to}). Reload recommended.`,
    statusNeedPresetOff: (name) => `Non-Apple receipt detected: ${name} (turn Apple preset OFF)`,
    statusReviewNeeded: (name) => `Needs review: ${name} (auto-excluded from ZIP due to low confidence)`,
    sceneFixed: "Mode: Apple receipt preset",
    sceneGeneral: "Mode: Generic",
    sceneOcrFail: "Mode: OCR failed",
    sceneAnalyzing: "Mode: Analyzing",
    sceneInvalidInput: "Mode: Invalid input image",
    sceneNonApplePreset: "Mode: Non-Apple (preset blocked)",
    sceneReviewNeeded: "Mode: Needs review (low confidence)",
    invalidTitlePlaceholder: "Invalid input (cannot save)",
    nonAppleTitlePlaceholder: "Non-Apple (save with preset OFF)",
    ocrPrefix: "OCR:",
    noText: "No text detected",
    ocrFailMsg: "OCR failed. Check network/CDN blocking.",
    ocrUiCaptureMsg: "This looks like a screenshot of this tool UI. Please upload the original receipt screenshot.",
    ocrNonApplePresetMsg: "Apple preset ON does not auto-name non-Apple receipts. Turn preset OFF to save.",
    ocrReviewMsg: (score, reasons) => `Confidence ${score}%. ${reasons ? `Review reasons: ${reasons}` : "Please verify against the original."}`,
    reviewReasonDate: "Date extraction is unstable",
    reviewReasonAmount: "Total amount extraction is unstable",
    reviewReasonService: "Service name extraction is unstable",
    reviewReasonCount: "Line-item count evidence is weak",
    reviewReasonConflict: "Line count and service name conflict",
    reviewReasonMixed: "Mixed store sections detected",
    statusNeedOriginal: (name) => `Invalid input image: ${name} (select original receipt screenshot)`,
    include: "Include in ZIP",
    duplicate: "Duplicate candidate (auto-excluded)",
    date: "Date",
    amount: "Total",
    itemCount: "Line Items",
    service: "Service / App",
    filename: "Final Filename",
    thumbTip: "Compare the original preview on the left with the extracted fields on the right.",
    loupeOn: "Loupe: ON",
    loupeOff: "Loupe: OFF",
    downloadSingle: "Download this file",
    originalPrefix: "Original:",
    savedPrefix: "Saved:",
    analyzing: "Analyzing",
  },
};

const DEFAULT_SERVICE_CATALOG = {
  aliases: [
    { canonical: "ABEMA", re: /abema|アベマ/i },
    { canonical: "MINNA_BANK", re: /みんなの銀行|minna\s*no\s*ginko/i },
    { canonical: "ICLOUD", re: /icloud\+?|i cloud/i },
    { canonical: "APP_STORE", re: /app\s?store|itunes/i },
    { canonical: "APPLE_MUSIC", re: /app[l1]e\s?music/i },
    { canonical: "APPLE_ONE", re: /app[l1]e\s?one/i },
    { canonical: "YOUTUBE", re: /youtube/i },
    { canonical: "NETFLIX", re: /netflix/i },
    { canonical: "CHATGPT", re: /chatgpt|openai/i },
    { canonical: "APPLE_DEVELOPER", re: /app[l1]e\s*developer|developer\s*program|デベロッパ/i },
    { canonical: "LIFE_CYCLE", re: /life\s*cycle/i },
  ],
  display: {
    ABEMA: "ABEMA",
    MINNA_BANK: "みんなの銀行",
    ICLOUD: "iCloud",
    APP_STORE: "AppStore",
    APPLE_MUSIC: "Apple Music",
    APPLE_ONE: "Apple One",
    YOUTUBE: "YouTube",
    NETFLIX: "Netflix",
    CHATGPT: "ChatGPT",
    APPLE_DEVELOPER: "Apple Developer Program",
    LIFE_CYCLE: "Life Cycle",
  },
};
const SERVICE_CATALOG = getServiceCatalog();
const LEARN_STORE_KEY = "apple_receipt_renamer_learning_v3";
const ENABLE_LEARN_OVERRIDES = false;
const BUILD_ID = "20260213af";
const APPLE_SINGLE_DEBUG_TARGET = "";
const PDFJS_WORKER_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
const MULTIPLICITY_ONLY_MODE = true;

init();

function getServiceCatalog() {
  const source = window.SCTOL_SERVICE_CATALOG || {};
  const aliasesSrc = Array.isArray(source.aliases) ? source.aliases : DEFAULT_SERVICE_CATALOG.aliases;
  const aliases = aliasesSrc
    .map((row) => {
      if (!row || typeof row.canonical !== "string" || !(row.re instanceof RegExp)) return null;
      return { canonical: row.canonical.trim(), re: row.re };
    })
    .filter(Boolean);

  return {
    aliases: aliases.length ? aliases : DEFAULT_SERVICE_CATALOG.aliases,
    display: {
      ...DEFAULT_SERVICE_CATALOG.display,
      ...(source.display && typeof source.display === "object" ? source.display : {}),
    },
  };
}

function init() {
  if (
    !dropZone ||
    !fileInput ||
    !statusText ||
    !resultList ||
    !resultTemplate ||
    !downloadAllBtn ||
    !exportCsvBtn ||
    !clearAllBtn ||
    !reloadBtn ||
    !compactMode ||
    !groupByMonth ||
    !langJaBtn ||
    !langEnBtn ||
    !heroTitle ||
    !heroLead ||
    !securityBadgeText ||
    !dropTitle ||
    !dropSub ||
    !flowTitle ||
    !flowStep1Label ||
    !flowStep1Sub ||
    !flowStep2Label ||
    !flowStep2Sub ||
    !flowStep3Label ||
    !flowStep3Sub ||
    !compactLabel ||
    !groupLabel ||
    !noticeTitle ||
    !noticeLead ||
    !noticeList ||
    !officialNote ||
    !contactLink ||
    !privacyLink ||
    !termsLink ||
    !updatesList
  ) {
    console.error("UI init failed");
    return;
  }
  setupPdfRuntime();

  fileInput.addEventListener("change", (e) => handleFiles(e.target.files));
  downloadAllBtn.addEventListener("click", downloadAllAsZip);
  exportCsvBtn.addEventListener("click", exportAsCsv);
  clearAllBtn.addEventListener("click", clearAllCards);
  reloadBtn.addEventListener("click", () => window.location.reload());
  compactMode.addEventListener("change", refreshAllCards);
  langJaBtn.addEventListener("click", () => setLanguage("ja"));
  langEnBtn.addEventListener("click", () => setLanguage("en"));
  downloadAllBtn.disabled = true;
  exportCsvBtn.disabled = true;
  clearAllBtn.disabled = true;
  applyLanguage();
  updateToolbarState();
  setStatus(withBuildTag(t("statusPresetOn")));
  notifyBuildUpdateIfNeeded();
  loadUpdates();

  window.addEventListener("dragover", preventDefaults);
  window.addEventListener("drop", preventDefaults);

  dropZone.addEventListener("dragenter", (e) => {
    e.preventDefault();
    dropZone.classList.add("active");
  });
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("active");
  });
  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("active"));
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("active");
    handleFiles(e.dataTransfer?.files);
  });
}

function setupPdfRuntime() {
  if (!window.pdfjsLib?.GlobalWorkerOptions) return;
  try {
    if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
    }
  } catch (error) {
    console.warn("Failed to initialize pdf.js worker:", error);
  }
}

function t(key) {
  const table = I18N[currentLang] || I18N.ja;
  return table[key];
}

const REVIEW_REASON_I18N_KEY = {
  date: "reviewReasonDate",
  amount: "reviewReasonAmount",
  service: "reviewReasonService",
  count: "reviewReasonCount",
  conflict: "reviewReasonConflict",
  mixed: "reviewReasonMixed",
};

function formatReviewReasons(reasonCodes = []) {
  const keys = uniqueNonEmpty(
    (reasonCodes || [])
      .map((code) => REVIEW_REASON_I18N_KEY[String(code || "").trim().toLowerCase()] || "")
      .filter(Boolean)
  );
  return keys.map((key) => t(key)).filter(Boolean).join(" / ");
}

function buildReviewExcerpt(item) {
  const score = Number.isFinite(item?.confidenceScore) ? item.confidenceScore : 0;
  const reasons = formatReviewReasons(item?.confidenceReasons || []);
  return t("ocrReviewMsg")(score, reasons);
}

function notifyBuildUpdateIfNeeded() {
  try {
    const prev = localStorage.getItem(BUILD_SEEN_KEY) || "";
    if (prev && prev !== BUILD_ID) {
      setStatus(withBuildTag(t("statusBuildChanged")(prev, BUILD_ID)), true);
    }
    localStorage.setItem(BUILD_SEEN_KEY, BUILD_ID);
  } catch {
    // ignore storage failures
  }
}

function setLanguage(lang) {
  currentLang = lang === "en" ? "en" : "ja";
  localStorage.setItem(LANG_STORE_KEY, currentLang);
  applyLanguage();
  refreshAllCards();
}

function applyLanguage() {
  const tr = I18N[currentLang] || I18N.ja;
  document.documentElement.lang = currentLang;
  document.title = tr.title;
  heroTitle.textContent = tr.title;
  heroLead.textContent = tr.hero;
  securityBadgeText.textContent = tr.securityBadge;
  dropTitle.textContent = tr.dropTitle;
  dropSub.textContent = tr.dropSub;
  if (dropZone) dropZone.setAttribute("aria-label", tr.dropAria);
  flowTitle.textContent = tr.flowTitle;
  flowStep1Label.textContent = tr.flowStep1Label;
  flowStep1Sub.textContent = tr.flowStep1Sub;
  flowStep2Label.textContent = tr.flowStep2Label;
  flowStep2Sub.textContent = tr.flowStep2Sub;
  flowStep3Label.textContent = tr.flowStep3Label;
  flowStep3Sub.textContent = tr.flowStep3Sub;
  compactLabel.textContent = tr.compactLabel;
  groupLabel.textContent = tr.groupLabel;
  downloadAllBtn.textContent = tr.downloadAll;
  exportCsvBtn.textContent = tr.exportCsv;
  clearAllBtn.textContent = tr.clearAll;
  reloadBtn.textContent = tr.reload;
  noticeTitle.textContent = tr.noticeTitle;
  if (noticeSection) noticeSection.setAttribute("aria-label", tr.noticeAria);
  noticeLead.textContent = tr.noticeLead;
  if (updatesTitle) updatesTitle.textContent = tr.updatesTitle;
  if (updatesSection) updatesSection.setAttribute("aria-label", tr.updatesAria);
  if (guidesTitle) guidesTitle.textContent = tr.guidesTitle;
  if (guidesSection) guidesSection.setAttribute("aria-label", tr.guidesAria);
  officialNote.textContent = tr.official;
  contactLink.textContent = tr.contact;
  privacyLink.textContent = tr.privacy;
  termsLink.textContent = tr.terms;
  if (guideExplainerTag) guideExplainerTag.textContent = tr.guideExplainerTag;
  if (guideExplainerTitle) guideExplainerTitle.textContent = tr.guideExplainerTitle;
  if (guideExplainerSub) guideExplainerSub.textContent = tr.guideExplainerSub;
  if (guideComparisonTag) guideComparisonTag.textContent = tr.guideComparisonTag;
  if (guideComparisonTitle) guideComparisonTitle.textContent = tr.guideComparisonTitle;
  if (guideComparisonSub) guideComparisonSub.textContent = tr.guideComparisonSub;
  if (guideHowtoTag) guideHowtoTag.textContent = tr.guideHowtoTag;
  if (guideHowtoTitle) guideHowtoTitle.textContent = tr.guideHowtoTitle;
  if (guideHowtoSub) guideHowtoSub.textContent = tr.guideHowtoSub;
  if (currentLang === "en") {
    contactLink.href = "./contact-en.html";
    privacyLink.href = "./privacy-en.html";
    termsLink.href = "./terms-en.html";
    if (guideExplainerLink) guideExplainerLink.href = "./guide-explainer-en.html";
    if (guideComparisonLink) guideComparisonLink.href = "./guide-comparison-en.html";
    if (guideHowtoLink) guideHowtoLink.href = "./guide-howto-en.html";
  } else {
    contactLink.href = "./contact.html";
    privacyLink.href = "./privacy.html";
    termsLink.href = "./terms.html";
    if (guideExplainerLink) guideExplainerLink.href = "./guide-explainer.html";
    if (guideComparisonLink) guideComparisonLink.href = "./guide-comparison.html";
    if (guideHowtoLink) guideHowtoLink.href = "./guide-howto.html";
  }
  noticeList.innerHTML = "";
  for (const item of tr.noticeItems) {
    const li = document.createElement("li");
    li.textContent = item;
    noticeList.appendChild(li);
  }
  langJaBtn.classList.toggle("active", currentLang === "ja");
  langEnBtn.classList.toggle("active", currentLang === "en");
  if (
    !statusText.textContent ||
    /\[build:/.test(statusText.textContent) ||
    statusText.textContent === I18N.ja.statusReady ||
    statusText.textContent === I18N.en.statusReady
  ) {
    setStatus(withBuildTag(tr.statusReady));
  }
  renderUpdates();
  for (const item of batchItems) {
    updateCardLanguage(item);
  }
}

async function loadUpdates() {
  if (Array.isArray(window.SCTOL_UPDATES)) {
    updatesData = window.SCTOL_UPDATES.filter((row) => row && row.date && (row.ja || row.en));
    renderUpdates();
    return;
  }

  try {
    const res = await fetch("./updates.json", { cache: "no-store" });
    if (!res.ok) throw new Error(`updates.json HTTP ${res.status}`);
    const data = await res.json();
    if (!Array.isArray(data)) throw new Error("updates.json must be array");
    updatesData = data.filter((row) => row && row.date && (row.ja || row.en));
  } catch (error) {
    console.error("Failed to load updates.json:", error);
    updatesData = [];
  }
  renderUpdates();
}

function renderUpdates() {
  if (!updatesList) return;
  updatesList.innerHTML = "";

  const rows = updatesData.length
    ? updatesData
    : [
        {
          date: "2026-02-11",
          ja: "updates.json を読み込めなかったため、デフォルト表示です。",
          en: "Fallback notice because updates.json could not be loaded.",
        },
      ];

  for (const row of rows) {
    const li = document.createElement("li");
    const date = document.createElement("span");
    date.textContent = row.date;
    const text = document.createTextNode(` ${currentLang === "en" ? row.en || row.ja : row.ja || row.en}`);
    li.appendChild(date);
    li.appendChild(text);
    updatesList.appendChild(li);
  }
}

async function handleFiles(fileList) {
  const files = fileList ? [...fileList] : [];
  const supportedFiles = files.filter(isSupportedInputFile);
  if (!supportedFiles.length) {
    setStatus(t("statusNeedImage"), true);
    return;
  }

  const normalizedFiles = await expandInputFiles(supportedFiles);
  if (!normalizedFiles.length) {
    fileInput.value = "";
    return;
  }

  setStatus(t("statusProcessing")(normalizedFiles.length));

  for (const file of normalizedFiles) {
    const item = buildCard(file);
    resultList.prepend(item.element);
    batchItems.unshift(item);
    await analyzeCard(item, false);
  }

  recomputeDuplicates();
  updateToolbarState();
  fileInput.value = "";
}

function isPdfFile(file) {
  if (!file) return false;
  const type = String(file.type || "").toLowerCase();
  const name = String(file.name || "");
  return type === "application/pdf" || /\.pdf$/i.test(name);
}

function isSupportedInputFile(file) {
  if (!file) return false;
  if (String(file.type || "").startsWith("image/")) return true;
  return isPdfFile(file);
}

async function expandInputFiles(files = []) {
  const expanded = [];
  for (const file of files) {
    if (!isSupportedInputFile(file)) continue;
    if (isPdfFile(file)) {
      try {
        const pages = await convertPdfToImageFiles(file);
        expanded.push(...pages);
      } catch (error) {
        console.error("PDF conversion failed:", error);
        setStatus(t("statusPdfFail")(file.name), true);
      }
      continue;
    }
    expanded.push(file);
  }
  return expanded;
}

function isCanvasToBlobSupported(canvas) {
  return canvas && typeof canvas.toBlob === "function";
}

function canvasToPngBlob(canvas) {
  return new Promise((resolve, reject) => {
    if (!isCanvasToBlobSupported(canvas)) {
      reject(new Error("canvas.toBlob is not supported"));
      return;
    }
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("Failed to convert canvas to blob"));
    }, "image/png");
  });
}

function formatPdfPageSuffix(page, total) {
  const digits = Math.max(2, String(Math.max(1, total)).length);
  return `p${String(page).padStart(digits, "0")}`;
}

function decoratePdfPageFile(pageFile, sourcePdfName, page, total) {
  pageFile.__fromPdf = true;
  pageFile.__sourcePdfName = sourcePdfName;
  pageFile.__pdfPage = page;
  pageFile.__pdfPageTotal = total;
}

async function convertPdfToImageFiles(pdfFile) {
  if (!window.pdfjsLib?.getDocument) {
    throw new Error("pdf.js not loaded");
  }
  setupPdfRuntime();

  const data = await pdfFile.arrayBuffer();
  const loadingTask = window.pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;
  const total = Math.max(1, pdf.numPages || 1);
  const pages = [];
  const base = removeExt(pdfFile.name) || "document";

  for (let pageNum = 1; pageNum <= total; pageNum += 1) {
    setStatus(t("statusPdfConverting")(pdfFile.name, pageNum, total));
    const page = await pdf.getPage(pageNum);
    const rawViewport = page.getViewport({ scale: 1 });
    const basePixels = Math.max(1, rawViewport.width * rawViewport.height);
    const maxPixels = 8_500_000;
    const targetScale = Math.sqrt(maxPixels / basePixels);
    const scale = Math.max(1.4, Math.min(2.4, targetScale));
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.floor(viewport.width));
    canvas.height = Math.max(1, Math.floor(viewport.height));
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) {
      throw new Error("Failed to create canvas context for PDF page");
    }

    await page.render({ canvasContext: ctx, viewport, intent: "print" }).promise;
    const blob = await canvasToPngBlob(canvas);
    const suffix = formatPdfPageSuffix(pageNum, total);
    const pageFile = new File([blob], `${base}-${suffix}.png`, {
      type: "image/png",
      lastModified: pdfFile.lastModified || Date.now(),
    });
    decoratePdfPageFile(pageFile, pdfFile.name, pageNum, total);
    pages.push(pageFile);

    page.cleanup?.();
    canvas.width = 1;
    canvas.height = 1;
  }

  pdf.cleanup?.();
  loadingTask.destroy?.();
  return pages;
}

function originalLabelForFile(file) {
  if (!file) return "";
  if (file.__fromPdf) {
    const page = Number(file.__pdfPage || 1);
    const total = Number(file.__pdfPageTotal || 1);
    const source = file.__sourcePdfName || file.name || "PDF";
    return `${source} (PDF ${page}/${total})`;
  }
  return file.name || "";
}

async function analyzeCard(item, isRetry = false) {
  try {
    setCardAnalyzingState(item, true);
    applyCardState(item);
    const sourceName = originalLabelForFile(item.file);
    if (isRetry) setStatus(t("statusRetrying")(sourceName));
    else setStatus(t("statusAnalyzing")(sourceName));
    const ocrData = await extractOcrData(item.file);
    item.ocrData = ocrData;
    item.ocrText = ocrData.text;
    item.learnKey = makeLearningKey(ocrData.text, item.file.name);
    item.issuer = detectIssuerCode(ocrData.text || "");

    if (isLikelyToolUiCapture(ocrData.text || "")) {
      item.invalidInput = true;
      item.nonApplePresetBlocked = false;
      item.reviewRequired = false;
      item.confidenceScore = 0;
      item.confidenceReasons = [];
      item.fields = {
        company: "APPLE",
        date: "",
        amount: "",
        itemCount: 1,
        service: "",
      };
      bindFieldsToCard(item);
      applyFieldsToCard(item);
      applyCardState(item);
      item.include = false;
      item.manualIncludeTouched = true;
      item.includeCheckbox.checked = false;
      item.sceneEl.textContent = withBuildTag(t("sceneInvalidInput"));
      item.excerpt.textContent = t("ocrUiCaptureMsg");
      setStatus(t("statusNeedOriginal")(sourceName), true);
      return;
    }

    try {
      if (item.issuer === "APPLE") {
        item.fields = extractAppleFields(ocrData, item.file.name);
      } else {
        item.fields = extractGenericFields(ocrData, item.file.name, item.issuer);
      }
    } catch (fieldError) {
      console.error("Field extraction failed:", fieldError);
      if (item.issuer === "APPLE") {
        item.fields = buildFallbackFieldsFromOcr(ocrData, item.file.name);
      } else {
        item.fields = buildGenericFieldsFromOcr(ocrData, item.file.name, item.issuer);
      }
    }

    if (item.issuer === "APPLE") {
      const currentDate = normalizeDateValue(item.fields?.date || "");
      const needsDateRecovery = (
        !isPlausibleDate(currentDate) ||
        !hasAnchoredAppleDateEvidence(ocrData, currentDate)
      );
      if (needsDateRecovery) {
        const recoveredDate = await recoverAppleBillingDateFromHeader(item.file, ocrData);
        if (isPlausibleDate(recoveredDate)) item.fields.date = recoveredDate;
      }
    }

    const wasInvalid = item.invalidInput;
    item.invalidInput = false;
    if (wasInvalid) {
      item.manualIncludeTouched = false;
      item.include = true;
      item.includeCheckbox.checked = true;
    }
    updatePresetBlockState(item);
    if (item.nonApplePresetBlocked) {
      item.reviewRequired = false;
      item.confidenceScore = 100;
      item.confidenceReasons = [];
    }

    if (ENABLE_LEARN_OVERRIDES) {
      const learned = loadLearnedOverrides(item.learnKey);
      if (learned && shouldApplyLearnedOverrides(item.fields, learned)) {
        item.fields = { ...item.fields, ...learned };
      }
    }
    if (item.issuer === "APPLE") {
      item.fields = enforceConservativeNaming(item.fields, ocrData.text || "", ocrData);
    }
    reassessConfidence(item, false);
    bindFieldsToCard(item);
    applyFieldsToCard(item);
    applyCardState(item);
    if (item.nonApplePresetBlocked) {
      item.sceneEl.textContent = withBuildTag(t("sceneNonApplePreset"));
      item.excerpt.textContent = t("ocrNonApplePresetMsg");
      setStatus(t("statusNeedPresetOff")(sourceName), true);
    } else if (item.reviewRequired) {
      item.sceneEl.textContent = withBuildTag(t("sceneReviewNeeded"));
      item.excerpt.textContent = buildReviewExcerpt(item);
      setStatus(t("statusReviewNeeded")(sourceName), true);
    } else {
      item.sceneEl.textContent = withBuildTag(compactMode.checked ? t("sceneFixed") : t("sceneGeneral"));
      item.excerpt.textContent = `${t("ocrPrefix")} ${truncate(ocrData.text.replace(/\s+/g, " "), 140) || t("noText")}`;
      setStatus(isRetry ? t("statusRetryDone")(sourceName) : t("statusDone")(sourceName));
    }
  } catch (error) {
    console.error(error);
    const errMsg = truncate(normalizeDateText(error?.message || String(error || "")), 80);
    item.invalidInput = false;
    item.nonApplePresetBlocked = false;
    item.reviewRequired = false;
    item.confidenceScore = 0;
    item.confidenceReasons = [];
    item.issuer = "OTHER";
    item.fields = {
      company: issuerLabel(item.issuer),
      date: "",
      amount: "",
      itemCount: 1,
      service: issuerServiceFallback(item.issuer),
    };
    bindFieldsToCard(item);
    applyFieldsToCard(item);
    applyCardState(item);
    item.sceneEl.textContent = withBuildTag(t("sceneOcrFail"));
    item.excerpt.textContent = errMsg ? `${t("ocrFailMsg")} (${errMsg})` : t("ocrFailMsg");
    setStatus(t("statusOcrSkip")(originalLabelForFile(item.file)), true);
  } finally {
    setCardAnalyzingState(item, false);
    applyCardState(item);
    recomputeDuplicates();
    updateToolbarState();
  }
}

function clearAllCards() {
  for (const item of batchItems) {
    try {
      if (item.previewUrl) URL.revokeObjectURL(item.previewUrl);
    } catch {
      // ignore URL cleanup errors
    }
  }
  batchItems.length = 0;
  resultList.innerHTML = "";
  fileInput.value = "";
  updateToolbarState();
  setStatus(withBuildTag(t("statusCleared")));
}

function updateToolbarState() {
  const hasAny = batchItems.length > 0;
  downloadAllBtn.disabled = !hasAny;
  exportCsvBtn.disabled = !hasAny;
  clearAllBtn.disabled = !hasAny;
  if (step2Panel) step2Panel.hidden = !hasAny;
  if (flowSection) flowSection.hidden = hasAny;
  if (appRoot) {
    appRoot.classList.toggle("has-items", hasAny);
    appRoot.classList.toggle("is-empty", !hasAny);
  }
}

function updatePresetBlockState(item) {
  item.nonApplePresetBlocked = Boolean(!item.invalidInput && compactMode.checked && item.issuer && item.issuer !== "APPLE");
}

function buildFallbackFieldsFromOcr(ocrData, fallbackName = "") {
  const text = ocrData?.text || "";
  let amount = "";
  try {
    amount = extractAppleTotalAmount(text, ocrData);
  } catch {
    amount = "";
  }
  let date = "";
  try {
    date = extractAppleBillingDate(text, ocrData);
  } catch {
    date = "";
  }
  let service = "";
  try {
    const names = extractAppNamesFromStoreBlock(ocrData, text);
    const known = canonicalServiceName(text);
    service = names[0] || (known ? displayServiceName(known) : "") || normalizeServiceName(removeExt(fallbackName));
  } catch {
    service = normalizeServiceName(removeExt(fallbackName));
  }
  const itemCount = normalizeItemCount(estimateAppleItemCount(text, ocrData, amount));
  const fields = {
    company: "APPLE",
    date: normalizeDateValue(date),
    amount: normalizeAmountText(amount),
    itemCount,
    service: sanitizeServiceFinal(service || "AppleService", "general"),
  };
  return enforceConservativeNaming(fields, text, ocrData);
}

const ISSUER_LABELS = {
  APPLE: "APPLE",
  GOOGLE: "GOOGLE",
  MICROSOFT: "MICROSOFT",
  AMAZON: "AMAZON",
  OPENAI: "OPENAI",
  OTHER: "OTHER",
};

function issuerLabel(code = "OTHER") {
  return ISSUER_LABELS[code] || "OTHER";
}

function issuerServiceFallback(code = "OTHER") {
  if (code === "APPLE") return "AppleService";
  if (code === "GOOGLE") return "GoogleService";
  if (code === "MICROSOFT") return "MicrosoftService";
  if (code === "AMAZON") return "AmazonService";
  if (code === "OPENAI") return "OpenAIService";
  return "Service";
}

function detectIssuerCode(text = "") {
  const n = normalizeDateText(text || "").toLowerCase();
  if (!n) return "OTHER";

  const googleScore =
    ((n.match(/g[o0]{2}g[l1]e\s*p[l1]ay/g) || []).length * 3) +
    ((n.match(/g[o0]{2}g[l1]e\s*d[i1]g[i1]ta[l1]/g) || []).length * 3) +
    ((n.match(/g[o0]{2}g[l1]e\s*[o0]ne/g) || []).length * 2);
  const appleScore =
    ((n.match(/app[l1]e\s*id/g) || []).length * 3) +
    ((n.match(/app[l1]e\s*acc[o0]unt/g) || []).length * 3) +
    ((n.match(/app\s*st[o0]re|app[l1]e\s*b[o0]{2}ks|i[t1]unes|icloud/g) || []).length * 2);
  const microsoftScore = (n.match(/microsoft|xbox|office\s*365|azure/g) || []).length * 2;
  const amazonScore = (n.match(/amazon|aws|kindle|prime\s*video/g) || []).length * 2;
  const openaiScore = (n.match(/openai|chatgpt/g) || []).length * 2;

  const scored = [
    ["GOOGLE", googleScore],
    ["APPLE", appleScore],
    ["MICROSOFT", microsoftScore],
    ["AMAZON", amazonScore],
    ["OPENAI", openaiScore],
  ].sort((a, b) => b[1] - a[1]);

  if (scored[0][1] >= 2) return scored[0][0];
  return "OTHER";
}

function buildGenericFieldsFromOcr(ocrData, fallbackName = "", issuer = "OTHER") {
  const text = ocrData?.text || "";
  const amount = normalizeAmountText(extractAppleTotalAmount(text, ocrData) || parseMoneyFromText(text));
  const date = normalizeDateValue(extractAppleBillingDate(text, ocrData) || extractDate(text, collectYearHints(text, getLines(ocrData))));
  const service = extractGenericServiceName(ocrData, text, fallbackName, issuer);
  const itemCount = normalizeItemCount(estimateGenericItemCount(ocrData, text, amount));
  return {
    company: issuerLabel(issuer),
    date,
    amount,
    itemCount,
    service,
  };
}

function extractGenericFields(ocrData, fallbackName = "", issuer = "OTHER") {
  const fields = buildGenericFieldsFromOcr(ocrData, fallbackName, issuer);
  return fields;
}

function estimateGenericItemCount(ocrData, fullText = "", totalAmount = "") {
  const lines = getLines(ocrData);
  const total = parseMoneyValue(totalAmount);
  let rows = 0;
  for (const line of lines) {
    const raw = normalizeDateText(line.text || "");
    if (!raw) continue;
    if (!/[¥￥$]\s?\d|\d+\s?円/.test(raw)) continue;
    if (/合計|total|subtotal|小計|tax|税|お支払い方法|payment|billing|visa|master|amex|jcb/i.test(raw)) continue;
    for (const m of parseMoneyListFromTextAll(raw)) {
      const v = parseMoneyValue(m);
      if (!v || (total > 0 && Math.abs(v - total) <= 1)) continue;
      rows += 1;
      break;
    }
  }
  if (rows >= 1) return rows;
  const inferred = inferLineItemCountFromTextAmounts(fullText, totalAmount);
  return inferred >= 1 ? inferred : 1;
}

function extractGenericServiceName(ocrData, text = "", fallbackName = "", issuer = "OTHER") {
  const n = normalizeDateText(text || "");
  if (issuer === "GOOGLE") {
    const g = extractGoogleServiceName(ocrData, n);
    if (g) return g;
  }

  const lines = getLines(ocrData);
  for (const line of lines) {
    const raw = normalizeDateText(line.text || "");
    if (!raw) continue;
    if (!/[¥￥$]\s?\d|\d+\s?円/.test(raw)) continue;
    if (/合計|total|subtotal|小計|tax|税|お支払い方法|payment|billing|visa|master|amex|jcb/i.test(raw)) continue;
    const c = extractNameFromMoneyLine(raw);
    if (c && !isLikelyGarbageServiceName(c)) return sanitizeGenericServiceName(c, issuer);
  }

  const known = canonicalServiceName(text);
  if (known && !["APP_STORE", "ICLOUD"].includes(known)) {
    return sanitizeGenericServiceName(displayServiceName(known), issuer);
  }
  return sanitizeGenericServiceName(normalizeServiceName(removeExt(fallbackName)), issuer);
}

function extractGoogleServiceName(ocrData, text = "") {
  const lines = getLines(ocrData);
  const start = lines.findIndex((line) => /アイテム|item/i.test(line.text || ""));
  const end = lines.findIndex((line, idx) => idx > start && /合計|total|お支払い方法|payment|billing/i.test(line.text || ""));
  const begin = start >= 0 ? start + 1 : 0;
  const stop = end > begin ? end : lines.length;
  for (let i = begin; i < stop; i += 1) {
    const raw = normalizeDateText(lines[i].text || "");
    if (!raw) continue;
    if (!/[¥￥$]\s?\d|\d+\s?円/.test(raw)) continue;
    let c = extractNameFromMoneyLine(raw);
    c = normalizeServiceName(c)
      .replace(/\(\s*開発元[^)]*\)/gi, " ")
      .replace(/開発元[^)\n]*/gi, " ")
      .replace(/\(\s*g[o0]{2}g[l1]e\s*[o0]ne[^)]*\)/gi, " ")
      .replace(/\b(g[o0]{2}g[l1]e|llc|inc)\b/gi, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (c && !isLikelyGarbageServiceName(c)) return sanitizeGenericServiceName(c, "GOOGLE");
  }

  const direct = normalizeDateText(text || "").match(/(g[o0]{2}g[l1]e\s*a[i1]\s*u[l1]tra[^¥\n]*)/i);
  if (direct?.[1]) return sanitizeGenericServiceName(direct[1], "GOOGLE");
  return "";
}

function sanitizeGenericServiceName(value, issuer = "OTHER") {
  let v = String(value || "")
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/g[0o]{2}g[l1i]e/gi, "Google")
    .replace(/\ba[l1i]\b/gi, "AI")
    .replace(/\bu[l1i]tra\b/gi, "Ultra")
    .replace(/\b0ne\b/gi, "One")
    .replace(/\(\s*\d+\s*(gb|tb)\s*\)/gi, " ")
    .replace(/\(\s*google\s*one[^)]*\)/gi, " ")
    .replace(/\(\s*開発元[^)]*\)/gi, " ")
    .replace(/開発元[^)\n]*/gi, " ");
  v = trimServiceSuffixes(normalizeServiceName(v));
  if (!v || isLikelyGarbageServiceName(v)) return issuerServiceFallback(issuer);
  return v;
}

function buildCard(file) {
  const node = resultTemplate.content.firstElementChild.cloneNode(true);
  const thumbWrap = node.querySelector(".thumb-wrap");
  const thumb = node.querySelector(".thumb");
  const loupe = node.querySelector(".thumb-loupe");
  const loupeToggle = node.querySelector(".loupe-toggle");
  const includeCheckbox = node.querySelector(".include-checkbox");
  const dupBadge = node.querySelector(".dup-badge");
  const dateInput = node.querySelector(".date-input");
  const amountInput = node.querySelector(".amount-input");
  const itemCountInput = node.querySelector(".itemcount-input");
  const serviceInput = node.querySelector(".service-input");
  const titleInput = node.querySelector(".title-input");
  const original = node.querySelector(".original");
  const sceneEl = node.querySelector(".scene");
  const excerpt = node.querySelector(".excerpt");
  const retryBtn = node.querySelector(".retry-btn");
  const downloadBtn = node.querySelector(".download-btn");

  const previewUrl = URL.createObjectURL(file);
  thumb.src = previewUrl;
  original.textContent = `${t("originalPrefix")} ${originalLabelForFile(file)}`;
  sceneEl.textContent = withBuildTag(compactMode.checked ? t("sceneFixed") : t("sceneGeneral"));
  excerpt.textContent = `${t("ocrPrefix")} ...`;

  const item = {
    file,
    previewUrl,
    element: node,
    thumbWrap,
    thumb,
    loupe,
    loupeToggle,
    loupeEnabled: false,
    ocrData: { text: "", lines: [] },
    ocrText: "",
    include: true,
    duplicate: false,
    manualIncludeTouched: false,
    invalidInput: false,
    nonApplePresetBlocked: false,
    analyzing: false,
    reviewRequired: false,
    confidenceScore: 100,
    confidenceReasons: [],
    issuer: "OTHER",
    learnKey: "",
    fields: { company: "APPLE", date: "", amount: "", itemCount: 1, service: "" },
    includeCheckbox,
    dupBadge,
    dateInput,
    amountInput,
    itemCountInput,
    serviceInput,
    titleInput,
    sceneEl,
    excerpt,
    retryBtn,
    downloadBtn,
  };

  if (MULTIPLICITY_ONLY_MODE && itemCountInput) {
    itemCountInput.type = "text";
    itemCountInput.readOnly = true;
  }

  includeCheckbox.addEventListener("change", () => {
    item.manualIncludeTouched = true;
    item.include = includeCheckbox.checked;
  });

  dateInput.addEventListener("input", () => {
    item.fields.date = normalizeDateValue(dateInput.value);
    reassessConfidence(item, true);
    updateCardTitle(item);
    applyCardState(item);
    saveLearnedOverrides(item.learnKey, item.fields);
    recomputeDuplicates();
  });

  amountInput.addEventListener("input", () => {
    item.fields.amount = normalizeAmountText(amountInput.value);
    reassessConfidence(item, true);
    updateCardTitle(item);
    applyCardState(item);
    saveLearnedOverrides(item.learnKey, item.fields);
    recomputeDuplicates();
  });

  if (itemCountInput) {
    itemCountInput.addEventListener("input", () => {
      if (MULTIPLICITY_ONLY_MODE) return;
      item.fields.itemCount = normalizeItemCount(itemCountInput.value);
      reassessConfidence(item, true);
      updateCardTitle(item);
      applyCardState(item);
      saveLearnedOverrides(item.learnKey, item.fields);
      recomputeDuplicates();
    });
  }

  serviceInput.addEventListener("input", () => {
    item.fields.service = normalizeServiceName(serviceInput.value);
    reassessConfidence(item, true);
    updateCardTitle(item);
    applyCardState(item);
    saveLearnedOverrides(item.learnKey, item.fields);
    recomputeDuplicates();
  });

  titleInput.addEventListener("input", () => {
    item.manualIncludeTouched = true;
  });

  if (retryBtn) {
    retryBtn.addEventListener("click", async () => {
      retryBtn.disabled = true;
      try {
        await analyzeCard(item, true);
      } finally {
        retryBtn.disabled = Boolean(item.invalidInput);
      }
    });
  }

  downloadBtn.addEventListener("click", () => {
    if (item.invalidInput) {
      setStatus(t("statusNeedOriginal")(originalLabelForFile(file)), true);
      return;
    }
    if (item.nonApplePresetBlocked) {
      setStatus(t("statusNeedPresetOff")(originalLabelForFile(file)), true);
      return;
    }
    const ext = extFromName(file.name) || extFromMime(file.type) || "png";
    const finalName = `${toSafeName(titleInput.value) || toSafeName(removeExt(file.name)) || "renamed"}.${ext}`;
    downloadBlob(file, finalName);
    excerpt.textContent = `${t("savedPrefix")} ${finalName}`;
    saveLearnedOverrides(item.learnKey, item.fields);
  });

  setupThumbLoupe(item);
  updateCardLanguage(item);
  return item;
}

function setupThumbLoupe(item) {
  const wrap = item.thumbWrap;
  const img = item.thumb;
  const loupe = item.loupe;
  const toggle = item.loupeToggle;
  if (!wrap || !img || !loupe || !toggle) return;

  let rafId = 0;
  let pendingPointer = null;

  const render = () => {
    rafId = 0;
    if (!item.loupeEnabled || !pendingPointer) return;

    const imgRect = img.getBoundingClientRect();
    const wrapRect = wrap.getBoundingClientRect();
    const naturalW = img.naturalWidth || 0;
    const naturalH = img.naturalHeight || 0;
    if (!naturalW || !naturalH || imgRect.width < 2 || imgRect.height < 2) {
      loupe.style.opacity = "0";
      return;
    }

    const scale = Math.min(imgRect.width / naturalW, imgRect.height / naturalH);
    const displayW = naturalW * scale;
    const displayH = naturalH * scale;
    const offsetX = (imgRect.width - displayW) / 2;
    const offsetY = (imgRect.height - displayH) / 2;
    const localX = pendingPointer.clientX - imgRect.left;
    const localY = pendingPointer.clientY - imgRect.top;

    const isInside =
      localX >= offsetX &&
      localX <= offsetX + displayW &&
      localY >= offsetY &&
      localY <= offsetY + displayH;
    if (!isInside) {
      loupe.style.opacity = "0";
      return;
    }

    const xInDisplay = localX - offsetX;
    const yInDisplay = localY - offsetY;
    const lensSize = 140;
    const half = lensSize / 2;
    const pointerX = pendingPointer.clientX - wrapRect.left;
    const pointerY = pendingPointer.clientY - wrapRect.top;
    const pad = 4;
    const left = Math.max(pad, Math.min(wrapRect.width - lensSize - pad, pointerX - half));
    const top = Math.max(pad, Math.min(wrapRect.height - lensSize - pad, pointerY - half));
    const zoom = 2.6;

    loupe.style.left = `${left}px`;
    loupe.style.top = `${top}px`;
    loupe.style.backgroundImage = `url("${img.src}")`;
    loupe.style.backgroundSize = `${displayW * zoom}px ${displayH * zoom}px`;
    loupe.style.backgroundPosition = `${-(xInDisplay * zoom - half)}px ${-(yInDisplay * zoom - half)}px`;
    loupe.style.opacity = "1";
  };

  const queueRender = (ev) => {
    pendingPointer = { clientX: ev.clientX, clientY: ev.clientY };
    if (!rafId) rafId = requestAnimationFrame(render);
  };

  wrap.addEventListener("pointerenter", (ev) => {
    if (!item.loupeEnabled) return;
    queueRender(ev);
  });
  wrap.addEventListener("pointermove", (ev) => {
    if (!item.loupeEnabled) return;
    queueRender(ev);
  });
  wrap.addEventListener("pointerdown", (ev) => {
    if (!item.loupeEnabled) return;
    queueRender(ev);
  });
  wrap.addEventListener("pointerleave", () => {
    pendingPointer = null;
    loupe.style.opacity = "0";
  });
  toggle.addEventListener("click", () => {
    setLoupeEnabled(item, !item.loupeEnabled);
  });

  setLoupeEnabled(item, false);
}

function setLoupeEnabled(item, enabled) {
  item.loupeEnabled = Boolean(enabled);
  if (item.thumbWrap) item.thumbWrap.classList.toggle("loupe-enabled", item.loupeEnabled);
  if (item.loupe) item.loupe.style.opacity = "0";
  if (item.loupeToggle) item.loupeToggle.setAttribute("aria-pressed", item.loupeEnabled ? "true" : "false");
  updateLoupeToggleLabel(item);
}

function updateLoupeToggleLabel(item) {
  const tr = I18N[currentLang] || I18N.ja;
  if (!item.loupeToggle) return;
  item.loupeToggle.textContent = item.loupeEnabled ? tr.loupeOn : tr.loupeOff;
}

function isMultipleClassification(fields = {}) {
  const count = normalizeItemCount(fields?.itemCount || 1);
  const service = normalizeServiceName(fields?.service || "");
  return count >= 2 || isStrictMultipleServiceLabel(service);
}

function multiplicityLabel(fields = {}, lang = currentLang) {
  const isMulti = isMultipleClassification(fields);
  if (lang === "en") return isMulti ? "Multiple" : "Single";
  return isMulti ? "複数" : "単数";
}

function bindFieldsToCard(item) {
  item.dateInput.value = item.fields.date;
  item.amountInput.value = item.fields.amount;
  if (item.itemCountInput) {
    if (MULTIPLICITY_ONLY_MODE) {
      item.itemCountInput.type = "text";
      item.itemCountInput.readOnly = true;
      item.itemCountInput.value = multiplicityLabel(item.fields, currentLang);
    } else {
      item.itemCountInput.readOnly = false;
      item.itemCountInput.value = String(normalizeItemCount(item.fields.itemCount));
    }
  }
  item.serviceInput.value = item.fields.service;
}

function applyFieldsToCard(item) {
  updateCardTitle(item);
}

function updateCardTitle(item) {
  if (item.analyzing) {
    item.titleInput.value = "";
    item.titleInput.placeholder = t("analyzing");
    return;
  }
  if (item.invalidInput) {
    item.titleInput.value = "";
    item.titleInput.placeholder = t("invalidTitlePlaceholder");
    return;
  }
  if (item.nonApplePresetBlocked) {
    item.titleInput.value = "";
    item.titleInput.placeholder = t("nonAppleTitlePlaceholder");
    return;
  }
  item.titleInput.placeholder = "";
  item.titleInput.value = buildPresetTitle(item.fields);
}

function setCardAnalyzingState(item, analyzing) {
  item.analyzing = Boolean(analyzing);
  if (!item.analyzing) {
    const prevType = item.itemCountInput?.dataset?.prevType || "";
    if (prevType && item.itemCountInput && item.itemCountInput.type !== prevType) {
      item.itemCountInput.type = prevType;
    }
    if (item.itemCountInput?.dataset) delete item.itemCountInput.dataset.prevType;
    return;
  }
  const p = t("analyzing");
  item.dateInput.value = p;
  item.amountInput.value = p;
  if (item.itemCountInput) {
    if (item.itemCountInput.type !== "text") {
      item.itemCountInput.dataset.prevType = item.itemCountInput.type || "number";
      item.itemCountInput.type = "text";
    }
    item.itemCountInput.value = p;
  }
  item.serviceInput.value = p;
  item.titleInput.value = p;
  item.dateInput.placeholder = p;
  item.amountInput.placeholder = p;
  if (item.itemCountInput) item.itemCountInput.placeholder = p;
  item.serviceInput.placeholder = p;
  item.titleInput.placeholder = p;
  item.sceneEl.textContent = withBuildTag(t("sceneAnalyzing"));
}

function applyCardState(item) {
  const disabled = Boolean(item.analyzing || item.invalidInput || item.nonApplePresetBlocked);
  const controls = [item.dateInput, item.amountInput, item.itemCountInput, item.serviceInput, item.titleInput];
  for (const control of controls) {
    if (control) control.disabled = disabled;
  }
  if (item.analyzing) {
    const p = t("analyzing");
    item.dateInput.placeholder = p;
    item.amountInput.placeholder = p;
    if (item.itemCountInput) item.itemCountInput.placeholder = p;
    item.serviceInput.placeholder = p;
    item.titleInput.placeholder = p;
  } else if (disabled) {
    item.dateInput.placeholder = "";
    item.amountInput.placeholder = "";
    if (item.itemCountInput) item.itemCountInput.placeholder = "";
    item.serviceInput.placeholder = "";
  }
  if (item.includeCheckbox) {
    item.includeCheckbox.disabled = disabled;
  }
  if (item.downloadBtn) {
    item.downloadBtn.disabled = disabled;
  }
  if (item.retryBtn) {
    item.retryBtn.disabled = Boolean(item.analyzing || item.invalidInput);
  }
  updateCardTitle(item);
}

function reassessConfidence(item, refreshScene = false) {
  if (!item || item.analyzing || item.invalidInput || item.nonApplePresetBlocked) {
    item.reviewRequired = false;
    item.confidenceScore = item.invalidInput ? 0 : 100;
    item.confidenceReasons = [];
    return;
  }

  const confidence = assessExtractionConfidence(item, item.ocrData || {});
  item.reviewRequired = confidence.reviewRequired;
  item.confidenceScore = confidence.score;
  item.confidenceReasons = confidence.reasons;

  if (!item.manualIncludeTouched && item.includeCheckbox) {
    item.include = !item.reviewRequired;
    item.includeCheckbox.checked = item.include;
  }

  if (!refreshScene) return;
  if (item.reviewRequired) {
    item.sceneEl.textContent = withBuildTag(t("sceneReviewNeeded"));
    item.excerpt.textContent = buildReviewExcerpt(item);
  } else {
    item.sceneEl.textContent = withBuildTag(compactMode.checked ? t("sceneFixed") : t("sceneGeneral"));
    item.excerpt.textContent = `${t("ocrPrefix")} ${truncate((item.ocrText || "").replace(/\s+/g, " "), 140) || t("noText")}`;
  }
}

function assessExtractionConfidence(item, ocrData = {}) {
  const reasons = [];
  let score = 100;
  const fields = item?.fields || {};

  const date = normalizeDateValue(fields.date || "");
  if (!isPlausibleDate(date)) {
    score -= 30;
    reasons.push("date");
  }

  const total = parseMoneyValue(fields.amount || "");
  if (!total) {
    score -= 28;
    reasons.push("amount");
  }

  const service = normalizeServiceName(fields.service || "");
  const genericFallback = /^(AppleService|GoogleService|MicrosoftService|AmazonService|OpenAIService|Service|AppStore)$/i.test(service);
  if (!service || genericFallback || isLikelyGarbageServiceName(service)) {
    score -= 24;
    reasons.push("service");
  }

  const count = normalizeItemCount(fields.itemCount || 1);
  if (item?.issuer === "APPLE") {
    const text = ocrData?.text || "";
    const lines = getLines(ocrData);
    const words = getWords(ocrData);
    const scope = findAppleCommerceScope(lines);
    const isMultiLabel = /^(複数|他複数|multi(?:\s*items?)?)$/i.test(service);

    if (scope.found) {
      const directItems = compactLineItems(extractAppleLineItems(lines, words, scope, fields.amount || ""));
      const scoped = lines.slice(scope.begin, scope.end);
      const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
      const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
      const distinctNames = uniqueNonEmpty(
        directItems
          .map((x) => normalizeProductNameCandidate(x.name || ""))
          .filter((x) => x && !isLikelyGarbageServiceName(x) && !isGenericServiceLabel(x) && !isSummaryLikeName(x))
      ).length;
      const uniquePrices = uniqueNonEmpty(
        collectPriceTokensFromLines(lines, scope, total).map((v) => String(v))
      ).length;
      const hardMax = Math.max(1, directItems.length, reportRows, reviewRows, distinctNames, uniquePrices);

      if (count > hardMax + 1 || (count >= 4 && hardMax <= 2)) {
        score -= 34;
        reasons.push("count");
      }

      if (isMultiLabel && hardMax <= 1) {
        score -= 22;
        reasons.push("conflict");
      }
      if (!isMultiLabel && count >= 2 && hardMax >= 2) {
        score -= 18;
        reasons.push("conflict");
      }
    }

    if (hasMixedCommerceSections(text) && !isMultiLabel) {
      score -= 16;
      reasons.push("mixed");
    }
  } else if (count >= 4 && total > 0) {
    score -= 12;
    reasons.push("count");
  }

  score = Math.max(0, Math.min(100, Math.round(score)));
  const uniqueReasons = uniqueNonEmpty(reasons.map(String)).slice(0, 3);
  const reviewRequired = score < 72 || uniqueReasons.includes("count") || uniqueReasons.includes("conflict");
  return { score, reasons: uniqueReasons, reviewRequired };
}

function updateCardLanguage(item) {
  const tr = I18N[currentLang] || I18N.ja;
  const includeText = item.element.querySelector(".include-text");
  const dateLabel = item.element.querySelector(".date-label");
  const amountLabel = item.element.querySelector(".amount-label");
  const itemCountLabel = item.element.querySelector(".itemcount-label");
  const serviceLabel = item.element.querySelector(".service-label");
  const filenameLabel = item.element.querySelector(".filename-label");
  const thumbTip = item.element.querySelector(".thumb-tip");
  if (includeText) includeText.textContent = tr.include;
  if (dateLabel) dateLabel.textContent = tr.date;
  if (amountLabel) amountLabel.textContent = tr.amount;
  if (itemCountLabel) itemCountLabel.textContent = MULTIPLICITY_ONLY_MODE
    ? (currentLang === "en" ? "Line Type" : "明細")
    : tr.itemCount;
  if (serviceLabel) serviceLabel.textContent = tr.service;
  if (filenameLabel) filenameLabel.textContent = tr.filename;
  if (thumbTip) thumbTip.textContent = tr.thumbTip;
  updateLoupeToggleLabel(item);
  item.dateInput.placeholder = "YYYYMMDD";
  item.amountInput.placeholder = currentLang === "en" ? "¥450 / $9.99" : "¥450";
  if (item.itemCountInput) {
    item.itemCountInput.placeholder = MULTIPLICITY_ONLY_MODE
      ? (currentLang === "en" ? "Single / Multiple" : "単数 / 複数")
      : "";
  }
  item.serviceInput.placeholder = currentLang === "en" ? "iCloud-200GB" : "iCloud-200GB";
  item.dupBadge.textContent = tr.duplicate;
  item.downloadBtnText = tr.downloadSingle;
  item.retryBtnText = tr.retry;
  const retry = item.element.querySelector(".retry-btn");
  if (retry) retry.textContent = tr.retry;
  const btn = item.element.querySelector(".download-btn");
  if (btn) btn.textContent = tr.downloadSingle;
  const original = item.element.querySelector(".original");
  if (original) original.textContent = `${tr.originalPrefix} ${originalLabelForFile(item.file)}`;
  updateCardTitle(item);
}

function refreshAllCards() {
  for (const item of batchItems) {
    if (item.analyzing) {
      item.sceneEl.textContent = withBuildTag(t("sceneAnalyzing"));
      updateCardLanguage(item);
      applyCardState(item);
      continue;
    }
    updatePresetBlockState(item);
    if (item.invalidInput) {
      item.sceneEl.textContent = withBuildTag(t("sceneInvalidInput"));
      updateCardLanguage(item);
      applyCardState(item);
      continue;
    }
    if (item.nonApplePresetBlocked) {
      item.sceneEl.textContent = withBuildTag(t("sceneNonApplePreset"));
      item.excerpt.textContent = t("ocrNonApplePresetMsg");
      updateCardLanguage(item);
      applyCardState(item);
      continue;
    }
    if (item.reviewRequired) {
      item.sceneEl.textContent = withBuildTag(t("sceneReviewNeeded"));
      item.excerpt.textContent = buildReviewExcerpt(item);
      updateCardLanguage(item);
      applyCardState(item);
      continue;
    }
    if (compactMode.checked) {
      updateCardTitle(item);
      item.sceneEl.textContent = withBuildTag(t("sceneFixed"));
    } else {
      item.sceneEl.textContent = withBuildTag(t("sceneGeneral"));
    }
    updateCardLanguage(item);
    applyCardState(item);
  }
  recomputeDuplicates();
  updateToolbarState();
  setStatus(withBuildTag(compactMode.checked ? t("statusPresetOn") : t("statusPresetOff")));
}

function recomputeDuplicates() {
  const firstSeen = new Map();

  for (const item of batchItems) {
    const sig = makeSignature(item.fields);
    const isDup = sig && firstSeen.has(sig);

    if (!sig) {
      item.duplicate = false;
      item.dupBadge.hidden = true;
      continue;
    }

    if (!isDup) {
      firstSeen.set(sig, true);
      item.duplicate = false;
      item.dupBadge.hidden = true;
    } else {
      item.duplicate = true;
      item.dupBadge.hidden = false;
      if (!item.manualIncludeTouched) {
        item.include = false;
        item.includeCheckbox.checked = false;
      }
    }
  }
}

function makeSignature(fields) {
  const date = normalizeDateValue(fields.date);
  const amount = normalizeAmountText(fields.amount);
  const service = normalizeServiceName(fields.service).toLowerCase();
  if (!date || !amount || !service) return "";
  return `${date}|${amount}|${service}`;
}

function buildPresetTitle(fields) {
  const company = normalizeCompanyName(fields.company || "APPLE");
  const date = normalizeDateValue(fields.date);
  const amount = normalizeAmountText(fields.amount);
  // Final filename stability is prioritized; do not append derived item-count suffixes.
  const service = normalizeServiceName(fields.service);
  const parts = [company, date, amount, service].filter(Boolean);
  return toSafeName(parts.join("-"));
}

async function downloadAllAsZip() {
  if (!batchItems.length) {
    setStatus(t("statusNeedFirst"), true);
    return;
  }
  if (!window.JSZip) {
    setStatus(t("statusZipLibErr"), true);
    return;
  }

  downloadAllBtn.disabled = true;
  setStatus(t("statusZipping"));

  try {
    const zip = new JSZip();
    const nameCounter = new Map();

    for (const item of batchItems) {
      if (!item.include || item.invalidInput || item.nonApplePresetBlocked) continue;
      const ext = extFromName(item.file.name) || extFromMime(item.file.type) || "png";
      const baseName = toSafeName(item.titleInput.value) || toSafeName(removeExt(item.file.name)) || "renamed";
      const finalName = getUniqueFilename(`${baseName}.${ext}`, nameCounter);
      if (groupByMonth.checked) {
        const folder = monthFolderFromDate(item.fields.date);
        zip.file(`${folder}/${finalName}`, item.file);
      } else {
        zip.file(finalName, item.file);
      }
    }

    const blob = await zip.generateAsync({ type: "blob" });
    const zipName = `apple-receipts-${dateStamp()}.zip`;
    downloadBlob(blob, zipName);
    setStatus(t("statusZipDone")(zipName));
  } catch (error) {
    console.error(error);
    setStatus(t("statusZipFail"), true);
  } finally {
    downloadAllBtn.disabled = false;
  }
}

function exportAsCsv() {
  if (!batchItems.length) {
    setStatus(t("statusNeedFirst"), true);
    return;
  }

  try {
    const headers = [
      "index",
      "include_in_zip",
      "duplicate",
      "date",
      "amount",
      "item_count",
      "service",
      "final_filename",
      "original_filename",
      "scene",
      "ocr_excerpt",
    ];

    const rows = batchItems.map((item, idx) => {
      const ext = extFromName(item.file.name) || extFromMime(item.file.type) || "png";
      const base = toSafeName(item.titleInput.value) || toSafeName(removeExt(item.file.name)) || "renamed";
      const finalName = `${base}.${ext}`;
      const scene = normalizeDateText((item.sceneEl?.textContent || "").replace(/\s*\[build:[^\]]+\]\s*$/i, ""));
      const excerpt = normalizeDateText(item.ocrText || "").slice(0, 400);
      return [
        idx + 1,
        item.include && !item.invalidInput && !item.nonApplePresetBlocked ? "1" : "0",
        item.duplicate ? "1" : "0",
        normalizeDateValue(item.fields?.date || ""),
        normalizeAmountText(item.fields?.amount || ""),
        MULTIPLICITY_ONLY_MODE
          ? multiplicityLabel(item.fields, "en")
          : normalizeItemCount(item.fields?.itemCount || 1),
        normalizeServiceName(item.fields?.service || ""),
        finalName,
        originalLabelForFile(item.file) || "",
        scene,
        excerpt,
      ];
    });

    const csvText = [headers, ...rows]
      .map((line) => line.map(csvEscape).join(","))
      .join("\r\n");
    const csvBlob = new Blob([`\uFEFF${csvText}`], { type: "text/csv;charset=utf-8;" });
    const csvName = `apple-receipts-${dateStamp()}.csv`;
    downloadBlob(csvBlob, csvName);
    setStatus(t("statusCsvDone")(csvName));
  } catch (error) {
    console.error(error);
    setStatus(t("statusCsvFail"), true);
  }
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\r\n]/.test(text)) return `"${text.replace(/"/g, "\"\"")}"`;
  return text;
}

async function extractOcrData(file) {
  if (!window.Tesseract?.recognize) throw new Error("Tesseract not loaded");

  const first = await runOcrPassWithFallback(file, "OCR中");
  let best = first;
  let bestScore = scoreOcrRecord(first);

  // Retry with a high-contrast preprocessed image when OCR quality is low.
  if (bestScore < 75) {
    const preprocessed = await buildPreprocessedImage(file);
    if (preprocessed) {
      try {
        const second = await runOcrPassWithFallback(preprocessed, "OCR再解析中");
        const secondScore = scoreOcrRecord(second);
        if (secondScore > bestScore) {
          best = second;
          bestScore = secondScore;
        }
      } catch {
        // Keep first pass.
      }
    }
  }

  // Coverage recovery for long or suspicious receipts.
  if (needsSegmentedOcr(best) || isLikelyMultiItemReceipt(best)) {
    const segmented = await runSegmentedOcrPass(file, "OCR分割解析中");
    const segmentedScore = scoreOcrRecord(segmented);
    const merged = mergeOcrRecords(best, segmented);
    const mergedScore = scoreOcrRecord(merged);
    if (mergedScore > bestScore) {
      best = merged;
      bestScore = mergedScore;
    } else if (segmentedScore > bestScore) {
      best = segmented;
      bestScore = segmentedScore;
    }
  }
  return best;
}

async function runOcrPass(input, label = "OCR中", lang = "jpn+eng") {
  const result = await Tesseract.recognize(input, lang, {
    logger: (m) => {
      if (m.status === "recognizing text") {
        setStatus(`${label}... ${Math.round(m.progress * 100)}%`);
      }
    },
  });
  return normalizeOcrResult(result.data);
}

async function runOcrPassWithFallback(input, label = "OCR中") {
  const langs = ["jpn+eng", "jpn", "eng"];
  let lastError = null;
  for (const lang of langs) {
    try {
      return await runOcrPass(input, `${label} [${lang}]`, lang);
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("OCR failed");
}

function normalizeOcrResult(data = {}) {
  const text = (data.text || "").trim();
  const lines = (data.lines || [])
    .map((line, idx) => ({
      text: (line.text || "").trim(),
      y0: Number.isFinite(line?.bbox?.y0) ? line.bbox.y0 : idx * 10,
      x0: Number.isFinite(line?.bbox?.x0) ? line.bbox.x0 : 0,
    }))
    .filter((line) => line.text);
  const words = (data.words || [])
    .map((word, idx) => ({
      text: (word.text || "").trim(),
      y0: Number.isFinite(word?.bbox?.y0) ? word.bbox.y0 : idx * 6,
      x0: Number.isFinite(word?.bbox?.x0) ? word.bbox.x0 : 0,
    }))
    .filter((word) => word.text);
  lines.sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
  words.sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
  return {
    text,
    lines: dedupeOcrLines(lines, 10),
    words: dedupeOcrWords(words, 8, 30),
  };
}

function scoreOcrRecord(record) {
  const text = normalizeDateText(record?.text || "");
  let score = 0;
  score += Math.min(parseMoneyListFromTextAll(text).length || 0, 24) * 14;
  score += ((text.match(/問題を報告|report a problem/gi) || []).length) * 16;
  score += ((text.match(/app\s*st[o0]re|app[l1]e\s*b[o0]{2}ks|icloud|合計|ご注文合計/gi) || []).length) * 12;
  score += ((text.match(/20\d{2}\s*[\/.\-年]\s*\d{1,2}\s*[\/.\-月]\s*\d{1,2}\s*日?/g) || []).length) * 10;
  score += Math.min((record?.lines || []).length, 120) * 0.2;
  return score;
}

function needsSegmentedOcr(record) {
  const text = normalizeDateText(record?.text || "");
  const lineCount = (record?.lines || []).length;
  const hasCommerce = hasCommerceSectionHint(text);
  const hasTotal = /ご注文合計|合計|total/i.test(text);
  const moneyCount = parseMoneyListFromTextAll(text).length;
  const reportCount = (text.match(/問題を報告|report a problem/gi) || []).length;
  const mixedSections = hasMixedCommerceSections(text);
  if (hasCommerce && hasTotal && !mixedSections && moneyCount >= 2 && reportCount <= 1 && lineCount >= 16) return false;
  if (!hasCommerce) return true;
  if (!hasTotal) return true;
  if (moneyCount < 3) return true;
  if (lineCount < 26) return true;
  if (/app[l1]e\s*b[o0]{2}ks|i[t1]unes\s*store/i.test(text) && reportCount <= 2) return true;
  return false;
}

function isLikelyMultiItemReceipt(record) {
  const text = normalizeDateText(record?.text || "");
  if (!text) return false;
  const hasMixedStore = hasMixedCommerceSections(text);
  const reportCount = (text.match(/問題を報告|report a problem/gi) || []).length;
  const moneyCount = parseMoneyListFromTextAll(text).length;
  if (hasMixedStore) return true;
  if (reportCount >= 2 && moneyCount >= 4) return true;
  return false;
}

async function runSegmentedOcrPass(file, label = "OCR分割解析中") {
  try {
    const bmp = await createImageBitmap(file);
    const width = bmp.width;
    const height = bmp.height;
    const segments = [
      { top: 0.0, bottom: 0.48 },
      { top: 0.32, bottom: 0.78 },
      { top: 0.62, bottom: 1.0 },
    ];
    const all = [];

    for (let i = 0; i < segments.length; i += 1) {
      const s = segments[i];
      const sy = Math.max(0, Math.floor(height * s.top));
      const sh = Math.max(1, Math.floor(height * (s.bottom - s.top)));
      const canvas = document.createElement("canvas");
      const scale = 2;
      canvas.width = Math.max(1, Math.floor(width * scale));
      canvas.height = Math.max(1, Math.floor(sh * scale));
      const ctx = canvas.getContext("2d");
      if (!ctx) continue;
      ctx.drawImage(bmp, 0, sy, width, sh, 0, 0, canvas.width, canvas.height);

      const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!blob) continue;
      const part = await runOcrPassWithFallback(blob, `${label} ${i + 1}/${segments.length}`);

      // Map segmented OCR coordinates back to original image coordinates.
      const shiftedLines = part.lines.map((line) => ({
        ...line,
        y0: sy + (line.y0 || 0) / scale,
        x0: (line.x0 || 0) / scale,
      }));
      const shiftedWords = part.words.map((word) => ({
        ...word,
        y0: sy + (word.y0 || 0) / scale,
        x0: (word.x0 || 0) / scale,
      }));
      all.push({ text: part.text, lines: shiftedLines, words: shiftedWords });
    }
    bmp.close?.();

    if (!all.length) return normalizeOcrResult({});
    const merged = {
      text: all.map((x) => x.text || "").join("\n"),
      lines: all.flatMap((x) => x.lines || []),
      words: all.flatMap((x) => x.words || []),
    };
    merged.lines.sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
    merged.words.sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));

    merged.lines = dedupeOcrLines(merged.lines, 22);
    merged.words = dedupeOcrWords(merged.words, 20, 88);
    return merged;
  } catch {
    return normalizeOcrResult({});
  }
}

function mergeOcrRecords(a, b) {
  const left = normalizeOcrResult(a || {});
  const right = normalizeOcrResult(b || {});
  const merged = {
    text: `${left.text || ""}\n${right.text || ""}`.trim(),
    lines: [...(left.lines || []), ...(right.lines || [])],
    words: [...(left.words || []), ...(right.words || [])],
  };
  merged.lines.sort((x, y) => (x.y0 - y.y0) || (x.x0 - y.x0));
  merged.words.sort((x, y) => (x.y0 - y.y0) || (x.x0 - y.x0));

  merged.lines = dedupeOcrLines(merged.lines, 20);
  merged.words = dedupeOcrWords(merged.words, 18, 72);
  return normalizeOcrResult(merged);
}

function normalizeLineDedupKey(value = "") {
  return normalizeDateText(value || "")
    .toLowerCase()
    .replace(/[\s\-‐‑‒–—―_.,，。:：;；/\\|()[\]{}<>「」『』【】'"`´~]/g, "");
}

function isHintLineForDedup(value = "") {
  const t = normalizeDateText(value || "");
  return /問題を報告|report a problem|レビューを書く|write a review|更新\s*[：:]|updated\s*[:]/i.test(t);
}

function computeCoordinateScale(rows = [], axis = "y0", baseSpan = 1200, minScale = 0.7, maxScale = 2.8) {
  const vals = (rows || [])
    .map((row) => Number(row?.[axis]))
    .filter((v) => Number.isFinite(v));
  if (vals.length < 2) return 1;
  const min = Math.min(...vals);
  const max = Math.max(...vals);
  const span = Math.max(1, max - min);
  const raw = span / Math.max(1, Number(baseSpan) || 1);
  return Math.max(minScale, Math.min(maxScale, raw));
}

function dedupeOcrLines(lines = [], yTolerance = 12) {
  const sorted = [...(lines || [])].sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
  const out = [];
  const yScale = computeCoordinateScale(sorted, "y0", 1200, 0.7, 2.8);
  const yTol = Math.max(4, Number(yTolerance || 0) * yScale);
  const hintYExtra = 12 * yScale;

  for (const line of sorted) {
    const text = String(line?.text || "").trim();
    if (!text) continue;
    const key = normalizeLineDedupKey(text);
    if (!key) continue;
    const y = Number(line?.y0 || 0);
    const x = Number(line?.x0 || 0);
    const hint = isHintLineForDedup(text);

    const dup = out.find((row) => {
      const nearY = Math.abs((row.y0 || 0) - y) <= (hint || row.__hint ? yTol + hintYExtra : yTol);
      if (!nearY) return false;
      if (row.__key === key) return true;
      if (hint && row.__hint) return true;
      return false;
    });

    if (dup) {
      if (text.length > String(dup.text || "").length) {
        dup.text = text;
        dup.__key = key;
      }
      dup.y0 = Math.min(dup.y0 || y, y);
      dup.x0 = Math.min(dup.x0 || x, x);
      continue;
    }

    out.push({
      text,
      y0: y,
      x0: x,
      __key: key,
      __hint: hint,
    });
  }

  return out.map(({ __key, __hint, ...line }) => line);
}

function dedupeOcrWords(words = [], yTolerance = 10, xTolerance = 40) {
  const sorted = [...(words || [])].sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
  const out = [];
  const yScale = computeCoordinateScale(sorted, "y0", 1200, 0.7, 2.8);
  const xScale = computeCoordinateScale(sorted, "x0", 900, 0.7, 2.8);
  const yTol = Math.max(3, Number(yTolerance || 0) * yScale);
  const xTol = Math.max(10, Number(xTolerance || 0) * xScale);
  const hintYExtra = 12 * yScale;
  const hintXExtra = 80 * xScale;

  for (const word of sorted) {
    const text = String(word?.text || "").trim();
    if (!text) continue;
    const key = normalizeLineDedupKey(text);
    if (!key) continue;
    const y = Number(word?.y0 || 0);
    const x = Number(word?.x0 || 0);
    const hint = isHintLineForDedup(text);

    const dup = out.find((row) => {
      const nearY = Math.abs((row.y0 || 0) - y) <= (hint || row.__hint ? yTol + hintYExtra : yTol);
      if (!nearY) return false;
      const nearX = Math.abs((row.x0 || 0) - x) <= (hint || row.__hint ? xTol + hintXExtra : xTol);
      if (!nearX) return false;
      return row.__key === key;
    });

    if (dup) {
      if (text.length > String(dup.text || "").length) {
        dup.text = text;
        dup.__key = key;
      }
      dup.y0 = Math.min(dup.y0 || y, y);
      dup.x0 = Math.min(dup.x0 || x, x);
      continue;
    }

    out.push({
      text,
      y0: y,
      x0: x,
      __key: key,
      __hint: hint,
    });
  }

  return out.map(({ __key, __hint, ...word }) => word);
}

async function buildPreprocessedImage(file) {
  try {
    const bmp = await createImageBitmap(file);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.floor(bmp.width * 2));
    canvas.height = Math.max(1, Math.floor(bmp.height * 2));
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
    bmp.close?.();

    const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const d = img.data;
    // Light denoise + contrast + binary threshold to stabilize OCR on tiny text.
    for (let i = 0; i < d.length; i += 4) {
      const lum = d[i] * 0.2126 + d[i + 1] * 0.7152 + d[i + 2] * 0.0722;
      const v = lum > 178 ? 255 : 0;
      d[i] = v;
      d[i + 1] = v;
      d[i + 2] = v;
    }
    ctx.putImageData(img, 0, 0);

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
    return blob || null;
  } catch {
    return null;
  }
}

async function recoverAppleBillingDateFromHeader(file, baseOcrData) {
  try {
    const baseDate = extractAppleBillingDate(baseOcrData?.text || "", baseOcrData || {});
    if (isPlausibleDate(baseDate)) return baseDate;

    const yearHints = collectYearHints(baseOcrData?.text || "", getLines(baseOcrData || {}));
    const bmp = await createImageBitmap(file);
    const regions = [
      { left: 0.0, top: 0.0, right: 1.0, bottom: 0.46, scale: 3 },
      { left: 0.03, top: 0.05, right: 0.58, bottom: 0.52, scale: 4 },
    ];

    for (let i = 0; i < regions.length; i += 1) {
      const region = regions[i];
      const blob = await buildHeaderCropBlob(bmp, region);
      if (!blob) continue;
      const pass = await runOcrPassWithFallback(blob, `日付再解析 ${i + 1}/${regions.length}`);
      const date = extractAppleBillingDate(pass.text || "", pass) || extractDate(pass.text || "", yearHints);
      if (isPlausibleDate(date)) {
        bmp.close?.();
        return date;
      }
    }
    bmp.close?.();
  } catch {
    // Ignore recovery errors and keep original fields.
  }
  return "";
}

async function buildHeaderCropBlob(imageBitmap, region) {
  if (!imageBitmap) return null;
  const width = imageBitmap.width;
  const height = imageBitmap.height;
  const left = Math.max(0, Math.floor(width * (region.left || 0)));
  const top = Math.max(0, Math.floor(height * (region.top || 0)));
  const right = Math.min(width, Math.ceil(width * (region.right || 1)));
  const bottom = Math.min(height, Math.ceil(height * (region.bottom || 1)));
  const cropW = Math.max(1, right - left);
  const cropH = Math.max(1, bottom - top);
  const scale = Math.max(2, Number(region.scale || 3));

  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.floor(cropW * scale));
  canvas.height = Math.max(1, Math.floor(cropH * scale));
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  ctx.drawImage(imageBitmap, left, top, cropW, cropH, 0, 0, canvas.width, canvas.height);
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const d = img.data;
  for (let i = 0; i < d.length; i += 4) {
    const lum = d[i] * 0.2126 + d[i + 1] * 0.7152 + d[i + 2] * 0.0722;
    const v = lum > 176 ? 255 : 0;
    d[i] = v;
    d[i + 1] = v;
    d[i + 2] = v;
  }
  ctx.putImageData(img, 0, 0);

  return new Promise((resolve) => canvas.toBlob((blob) => resolve(blob || null), "image/png"));
}

function isStrictMultipleServiceLabel(value = "") {
  return /^(複数|他複数|multi(?:\s*items?)?)$/i.test(normalizeServiceName(value || ""));
}

function collectAppleSingleMetrics(ocrData, totalAmount = "", summary = {}) {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const scope = findAppleCommerceScope(lines);
  const summaryCount = normalizeItemCount(summary?.itemCount || 1);
  if (!scope.found) {
    return {
      summaryCount,
      reportRows: 0,
      reviewRows: 0,
      uniquePrices: 0,
      amountRows: 0,
      distinctNames: 0,
      iconRows: 0,
      scopeFound: false,
    };
  }

  const total = parseMoneyValue(totalAmount);
  const scoped = lines.slice(scope.begin, scope.end);
  const items = compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount));
  return {
    summaryCount,
    reportRows: countHintRowsDedup(scoped, /問題を報告|report a problem/i),
    reviewRows: countHintRowsDedup(scoped, /レビューを書く|write a review/i),
    uniquePrices: extractPriceCandidatesFromLines(lines, scope, total).length,
    amountRows: countAmountRowsFromWords(lines, words, scope, totalAmount),
    distinctNames: extractDistinctItemNames(items || []).length,
    iconRows: countIconLikeItemRows(lines, scope, totalAmount),
    scopeFound: true,
  };
}

function shouldLogAppleSingleDebug(fallbackName = "") {
  const target = normalizeDateText(APPLE_SINGLE_DEBUG_TARGET).toLowerCase();
  if (!target) return false;
  const base = normalizeDateText(removeExt(fallbackName || "")).toLowerCase();
  return base.includes(target);
}

function logAppleSingleDebug(fallbackName = "", phase = "", payload = {}) {
  if (!shouldLogAppleSingleDebug(fallbackName)) return;
  console.info(`[apple-single-debug:${phase}]`, payload);
}

function chooseSingleReceiptService(
  ocrData,
  fullText = "",
  totalAmount = "",
  summary = {},
  fallbackName = "",
  firstSummary = {},
  secondSummary = {}
) {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const scope = findAppleCommerceScope(lines);
  const items = scope.found ? compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount)) : [];
  const primary = scope.found ? choosePrimaryServiceName(items, lines, scope, fullText) : "";
  const scopeText = scope.found
    ? lines.slice(scope.begin, scope.end).map((line) => normalizeDateText(line.text || "")).join(" ")
    : "";
  const explicitHits = findCanonicalServicesInText(`${normalizeDateText(fullText || "")} ${scopeText}`)
    .filter((x) => !["APP_STORE", "ICLOUD"].includes(x))
    .map((x) => displayServiceName(x));
  const appNames = extractAppNamesFromStoreBlock(ocrData, fullText);
  const candidates = [
    ...explicitHits,
    primary,
    appNames[0] || "",
    summary?.service || "",
    firstSummary?.service || "",
    secondSummary?.service || "",
    normalizeServiceName(removeExt(fallbackName || "")),
    "AppleService",
  ];

  for (const raw of candidates) {
    const candidate = normalizeServiceName(raw || "");
    if (!candidate) continue;
    if (isStrictMultipleServiceLabel(candidate)) continue;
    return sanitizeServiceFinal(candidate, "appstore");
  }
  return "AppleService";
}

function extractAppleFields(ocrData, fallbackName) {
  const text = ocrData.text || "";
  const firstSummary = extractAppleServiceSummary(text, ocrData, "", fallbackName);
  const rawAmount = extractAppleTotalAmount(text, ocrData);
  const firstAmount = chooseBestTotalAmount(rawAmount, firstSummary.itemPrices || []);
  const secondSummary = extractAppleServiceSummary(text, ocrData, firstAmount, fallbackName);
  const summary = {
    ...secondSummary,
    service:
      normalizeServiceName(secondSummary.service || "") ||
      normalizeServiceName(firstSummary.service || "") ||
      "AppleService",
    itemPrices:
      (secondSummary.itemPrices && secondSummary.itemPrices.length)
        ? secondSummary.itemPrices
        : (firstSummary.itemPrices || []),
    itemCount: normalizeItemCount(secondSummary.itemCount || firstSummary.itemCount || 1),
  };
  const amount = chooseBestTotalAmount(firstAmount, summary.itemPrices || []);
  const initialMetrics = collectAppleSingleMetrics(ocrData, amount, summary);
  const hardSingle = hasHardSingleReceiptEvidence(ocrData, text, amount, summary);
  const confirmedMulti = hasConfirmedMultipleEvidence(ocrData, text, amount, summary);
  const hardMulti = hasHardMultipleReceiptEvidence(ocrData, text, amount, summary);
  logAppleSingleDebug(fallbackName, "initial", {
    ...initialMetrics,
    hardSingle,
    confirmedMulti,
    hardMulti,
    itemCount: summary.itemCount,
    service: summary.service,
  });

  let singleOverride = null;
  if (hardSingle || !confirmedMulti) {
    singleOverride = resolveSingleItemOverride(ocrData, text, amount, summary, fallbackName);
    summary.itemCount = 1;
    if (singleOverride?.service) {
      summary.service = singleOverride.service;
    } else if (isStrictMultipleServiceLabel(summary.service || "")) {
      summary.service = sanitizeServiceFinal(firstSummary.service || secondSummary.service || "AppleService", "appstore");
    }
  }
  if (!singleOverride) {
    singleOverride = resolveSingleItemOverride(ocrData, text, amount, summary, fallbackName);
  }
  if (singleOverride) {
    summary.service = singleOverride.service;
    summary.itemCount = 1;
  } else if (!hardSingle && confirmedMulti && hardMulti && shouldForceMultipleFallback(summary, amount, text, ocrData)) {
    summary.service = "複数";
    summary.itemCount = normalizeMultipleItemCount(summary.itemCount);
  }

  const finalHardSingle = hasHardSingleReceiptEvidence(ocrData, text, amount, summary);
  const finalSingleHardRule = hasFinalSingleReceiptEvidence(ocrData, text, amount, summary);
  const finalSingleOverride = resolveSingleItemOverride(ocrData, text, amount, summary, fallbackName);
  const finalServiceIsMultiple = isStrictMultipleServiceLabel(summary.service || "");
  const finalConfirmedMulti = hasConfirmedMultipleEvidence(ocrData, text, amount, summary);
  const finalHardMulti = hasHardMultipleReceiptEvidence(ocrData, text, amount, summary);
  const finalShouldForceMulti = shouldForceMultipleFallback(summary, amount, text, ocrData);
  const finalLines = getLines(ocrData);
  const finalScope = findAppleCommerceScope(finalLines);
  const finalIconRows = finalScope.found ? countIconLikeItemRows(finalLines, finalScope, amount) : 0;
  const finalSectionKinds = finalScope.found ? detectCommerceSectionKinds(finalLines, finalScope) : new Set();
  const finalTotal = parseMoneyValue(amount);
  const finalUniquePrices = finalScope.found ? extractPriceCandidatesFromLines(finalLines, finalScope, finalTotal).length : 0;
  const finalYearHints = collectYearHints(text, finalLines);
  const finalDistinctUpdateDates = finalScope.found ? countDistinctUpdateDates(finalLines, finalScope, finalYearHints) : 0;
  const lowAmountSingleGuard = (
    finalTotal > 0 &&
    finalTotal <= 1200 &&
    finalSectionKinds.size <= 1 &&
    finalIconRows <= 1 &&
    finalUniquePrices <= 1
  );
  const finalHardMultiAnchors = (
    finalSectionKinds.size >= 2 ||
    finalIconRows >= 2
  );
  const keepMultipleLabel = (
    finalServiceIsMultiple &&
    !finalHardSingle &&
    !finalSingleHardRule &&
    !finalSingleOverride &&
    finalConfirmedMulti &&
    finalHardMulti &&
    finalHardMultiAnchors &&
    !lowAmountSingleGuard &&
    finalShouldForceMulti &&
    finalIconRows !== 1
  );

  // Final hard guard: when single evidence is present, force one-line output.
  if (finalHardSingle || finalSingleHardRule || finalSingleOverride || lowAmountSingleGuard) {
    summary.itemCount = 1;
    summary.service = sanitizeServiceFinal(
      finalSingleOverride?.service || chooseSingleReceiptService(
        ocrData,
        text,
        amount,
        summary,
        fallbackName,
        firstSummary,
        secondSummary
      ),
      "appstore"
    );
  } else if (finalServiceIsMultiple && !keepMultipleLabel) {
    summary.itemCount = 1;
    summary.service = chooseSingleReceiptService(
      ocrData,
      text,
      amount,
      summary,
      fallbackName,
      firstSummary,
      secondSummary
    );
  } else if (keepMultipleLabel) {
    summary.service = "複数";
    summary.itemCount = normalizeMultipleItemCount(summary.itemCount);
  }

  const finalMetrics = collectAppleSingleMetrics(ocrData, amount, summary);
  logAppleSingleDebug(fallbackName, "final", {
    ...finalMetrics,
    finalHardSingle,
    finalSingleHardRule,
    finalConfirmedMulti,
    finalHardMulti,
    finalHardMultiAnchors,
    lowAmountSingleGuard,
    finalShouldForceMulti,
    finalIconRows,
    finalUniquePrices,
    finalDistinctUpdateDates,
    finalSingleOverride: Boolean(finalSingleOverride),
    itemCount: summary.itemCount,
    service: summary.service,
  });

  return {
    company: "APPLE",
    date: extractAppleBillingDate(text, ocrData),
    amount,
    itemCount: summary.itemCount,
    service: summary.service,
  };
}

function hasHardSingleReceiptEvidence(ocrData, fullText = "", totalAmount = "", summary = {}) {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const scope = findAppleCommerceScope(lines);
  if (!scope.found) return false;

  const sectionKinds = detectCommerceSectionKinds(lines, scope);
  if (sectionKinds.size >= 2) return false;

  const total = parseMoneyValue(totalAmount);
  if (!total) return false;

  const items = compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount));
  const distinctNames = extractDistinctItemNames(items || []);
  const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
  const amountRows = countAmountRowsFromWords(lines, words, scope, totalAmount);
  const scoped = lines.slice(scope.begin, scope.end);
  const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
  const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
  const yearHints = collectYearHints(fullText, lines);
  const distinctUpdateDates = countDistinctUpdateDates(lines, scope, yearHints);
  const iconRows = countIconLikeItemRows(lines, scope, totalAmount);

  const itemValues = (items || [])
    .map((x) => Number(x.amountValue || 0))
    .filter((v) => Number.isFinite(v) && v > 0);
  const nonTotalValues = itemValues.filter((v) => Math.abs(v - total) > 1);
  const nonTotalLikelyNoise = (
    nonTotalValues.length === 0 ||
    (nonTotalValues.length === 1 && nonTotalValues[0] >= Math.max(80, Math.round(total * 0.82)))
  );
  const summaryCount = normalizeItemCount(summary?.itemCount || 1);

  return (
    summaryCount <= 2 &&
    distinctNames.length <= 1 &&
    uniquePrices.length <= 1 &&
    amountRows <= 2 &&
    reportRows <= 2 &&
    reviewRows <= 1 &&
    distinctUpdateDates <= 1 &&
    iconRows <= 1 &&
    nonTotalLikelyNoise
  );
}

function hasFinalSingleReceiptEvidence(ocrData, fullText = "", totalAmount = "", summary = {}) {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const scope = findAppleCommerceScope(lines);
  if (!scope.found) return false;

  const sectionKinds = detectCommerceSectionKinds(lines, scope);
  if (sectionKinds.size >= 2) return false;

  const total = parseMoneyValue(totalAmount);
  if (!total) return false;

  const items = compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount));
  const distinctNames = extractDistinctItemNames(items || []);
  const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
  const amountRows = countAmountRowsFromWords(lines, words, scope, totalAmount);
  const scoped = lines.slice(scope.begin, scope.end);
  const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
  const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
  const yearHints = collectYearHints(fullText, lines);
  const distinctUpdateDates = countDistinctUpdateDates(lines, scope, yearHints);
  const iconRows = countIconLikeItemRows(lines, scope, totalAmount);
  const summaryCount = normalizeItemCount(summary?.itemCount || 1);

  const itemValues = (items || [])
    .map((x) => Number(x.amountValue || 0))
    .filter((v) => Number.isFinite(v) && v > 0);
  const nonTotalValues = itemValues.filter((v) => Math.abs(v - total) > 1);
  const nonTotalLikelyNoise = (
    nonTotalValues.length === 0 ||
    (nonTotalValues.length === 1 && nonTotalValues[0] >= Math.max(80, Math.round(total * 0.82)))
  );

  const iconSingle = (
    iconRows === 1 &&
    sectionKinds.size <= 1 &&
    distinctUpdateDates <= 1 &&
    uniquePrices.length <= 1
  );
  const noisySingleHint = (
    summaryCount <= 2 &&
    total <= 1200 &&
    reportRows === 0 &&
    reviewRows === 0 &&
    distinctNames.length <= 1 &&
    distinctUpdateDates === 0 &&
    amountRows <= 2 &&
    iconRows <= 2 &&
    uniquePrices.length <= 2 &&
    nonTotalLikelyNoise
  );

  return (
    iconSingle || noisySingleHint || (
      summaryCount <= 2 &&
      distinctNames.length <= 1 &&
      uniquePrices.length <= 1 &&
      amountRows <= 2 &&
      reportRows <= 4 &&
      reviewRows <= 1 &&
      distinctUpdateDates <= 1 &&
      iconRows <= 1 &&
      nonTotalLikelyNoise
    )
  );
}

function hasConfirmedMultipleEvidence(ocrData, fullText = "", totalAmount = "", summary = {}) {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const scope = findAppleCommerceScope(lines);
  const summaryCount = normalizeItemCount(summary?.itemCount || 1);
  if (!scope.found) return summaryCount >= 2 && hasMixedCommerceSections(fullText);
  if (hasFinalSingleReceiptEvidence(ocrData, fullText, totalAmount, summary)) return false;

  const sectionKinds = detectCommerceSectionKinds(lines, scope);
  if (sectionKinds.size >= 2) return true;

  const scoped = lines.slice(scope.begin, scope.end);
  const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
  const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
  const yearHints = collectYearHints(fullText, lines);
  const distinctUpdateDates = countDistinctUpdateDates(lines, scope, yearHints);
  const total = parseMoneyValue(totalAmount);
  const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
  const amountRows = countAmountRowsFromWords(lines, words, scope, totalAmount);
  const items = compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount));
  const distinctNames = extractDistinctItemNames(items || []).length;
  const iconRows = countIconLikeItemRows(lines, scope, totalAmount);
  const likelyNoisySingle = (
    summaryCount <= 2 &&
    total <= 1200 &&
    reportRows === 0 &&
    reviewRows === 0 &&
    distinctNames <= 1 &&
    distinctUpdateDates === 0 &&
    amountRows <= 2 &&
    iconRows <= 2 &&
    uniquePrices.length <= 2
  );
  if (likelyNoisySingle) return false;
  if (iconRows === 1 && sectionKinds.size <= 1 && uniquePrices.length <= 1 && distinctUpdateDates <= 1) return false;
  if (distinctUpdateDates >= 2 && (uniquePrices.length >= 2 || iconRows >= 2)) return true;
  if (uniquePrices.length <= 1 && distinctNames <= 1 && summaryCount <= 2 && sectionKinds.size <= 1) return false;
  if (iconRows >= 2 && (reportRows >= 2 || reviewRows >= 1 || distinctNames >= 2 || summaryCount >= 2 || uniquePrices.length >= 1 || distinctUpdateDates >= 1)) return true;

  if (uniquePrices.length >= 2) return true;
  if ((reportRows >= 2 || reviewRows >= 2) && (uniquePrices.length >= 2 || iconRows >= 2)) return true;
  if (distinctNames >= 2 && (uniquePrices.length >= 2 || iconRows >= 2)) return true;
  if (summaryCount >= 3 && (uniquePrices.length >= 2 || iconRows >= 2)) return true;
  return false;
}

function hasHardMultipleReceiptEvidence(ocrData, fullText = "", totalAmount = "", summary = {}) {
  const lines = getLines(ocrData);
  const scope = findAppleCommerceScope(lines);
  if (!scope.found) return hasMixedCommerceSections(fullText);

  const sectionKinds = detectCommerceSectionKinds(lines, scope);
  if (sectionKinds.size >= 2) return true;

  const yearHints = collectYearHints(fullText, lines);
  const distinctUpdateDates = countDistinctUpdateDates(lines, scope, yearHints);
  const iconRows = countIconLikeItemRows(lines, scope, totalAmount);
  return iconRows >= 2 && distinctUpdateDates >= 2;
}

function resolveSingleItemOverride(ocrData, fullText = "", totalAmount = "", summary = {}, fallbackName = "") {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const scope = findAppleCommerceScope(lines);
  if (!scope.found) return null;

  const sectionKinds = detectCommerceSectionKinds(lines, scope);
  if (sectionKinds.size >= 2) return null;

  const total = parseMoneyValue(totalAmount);
  if (!total) return null;

  const items = compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount));
  const itemValues = items.map((x) => Number(x.amountValue || 0)).filter((v) => Number.isFinite(v) && v > 0);
  const allItemEqualTotal = itemValues.length > 0 && itemValues.every((v) => Math.abs(v - total) <= 1);
  const scoped = lines.slice(scope.begin, scope.end);
  const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
  const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
  const yearHints = collectYearHints(fullText, lines);
  const distinctUpdateDates = countDistinctUpdateDates(lines, scope, yearHints);
  const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
  const amountRows = countAmountRowsFromWords(lines, words, scope, totalAmount);
  const distinctNames = extractDistinctItemNames(items || []);
  const iconRows = countIconLikeItemRows(lines, scope, totalAmount);

  const summaryCount = normalizeItemCount(summary?.itemCount || 1);
  const summaryService = normalizeServiceName(summary?.service || "");
  const summaryIsMultipleLabel = /^(複数|他複数|multi(?:\s*items?)?)$/i.test(summaryService);
  const nearTotalOnly = itemValues.length > 0 && itemValues.every((v) => v >= Math.max(80, Math.round(total * 0.82)));
  const recoverFromMultiLabel =
    summaryIsMultipleLabel &&
    summaryCount <= 2 &&
    uniquePrices.length <= 1 &&
    distinctNames.length <= 1 &&
    amountRows <= 2 &&
    reportRows <= 2 &&
    reviewRows <= 1 &&
    distinctUpdateDates <= 1 &&
    iconRows <= 1 &&
    total <= 2500 &&
    nearTotalOnly;

  // Guard against false multi from duplicated total-row extraction and duplicated "report problem" lines.
  const likelySingleDuplicate =
    summaryCount <= 2 &&
    amountRows <= 2 &&
    uniquePrices.length <= 1 &&
    distinctNames.length <= 1 &&
    reviewRows <= 1 &&
    reportRows <= 2 &&
    distinctUpdateDates <= 1 &&
    iconRows <= 1 &&
    (allItemEqualTotal || nearTotalOnly);

  if (!likelySingleDuplicate && !recoverFromMultiLabel && !summaryIsMultipleLabel) return null;
  if (!likelySingleDuplicate && !recoverFromMultiLabel) return null;

  const primary = choosePrimaryServiceName(items, lines, scope, fullText);
  const appNames = extractAppNamesFromStoreBlock(ocrData, fullText);
  const service = sanitizeServiceFinal(
    primary || appNames[0] || summaryService || normalizeServiceName(removeExt(fallbackName)) || "AppleService",
    "appstore"
  );
  return {
    service,
    reason: "single-duplicate-guard",
  };
}

function shouldForceMultipleFallback(summary, amount, fullText = "", ocrData = null) {
  const count = normalizeItemCount(summary?.itemCount || 1);

  const total = parseMoneyValue(amount || "");
  const service = normalizeServiceName(summary?.service || "");
  const text = normalizeDateText(fullText || "");
  const hasCommerceHint = hasCommerceSectionHint(text);
  const hasMixedSections = hasMixedCommerceSections(text);
  let scopeChecked = false;
  let scopeMixedSections = false;
  const textPriceRows = parseMoneyListFromTextAll(text)
    .map((m) => parseMoneyValue(m))
    .filter((v) => Number.isFinite(v) && v >= 80 && (total <= 0 || Math.abs(v - total) > 1)).length;

  if (ocrData) {
    const lines = getLines(ocrData);
    const words = getWords(ocrData);
    const scope = findAppleCommerceScope(lines);
    if (scope.found) {
      scopeChecked = true;
      const items = compactLineItems(extractAppleLineItems(lines, words, scope, amount));
      if (count <= 2 && hasSingleItemEvidenceInScope(lines, words, scope, amount, items, fullText)) return false;
      const sectionKinds = detectCommerceSectionKinds(lines, scope);
      scopeMixedSections = sectionKinds.size >= 2;
      const scoped = lines.slice(scope.begin, scope.end);
      const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
      const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
      const priceTokens = countPriceTokensFromLines(lines, scope, total);
      const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
      const totalTokenCount = countTotalTokenOccurrencesInScope(lines, scope, amount);
      const distinctNames = extractDistinctItemNames(items || []);
      const yearHints = collectYearHints(fullText, lines);
      const distinctUpdateDates = countDistinctUpdateDates(lines, scope, yearHints);
      const iconRows = countIconLikeItemRows(lines, scope, amount);
      const strongMultiple = hasStrongMultipleEvidenceInScope(lines, scope, amount, items, fullText);
      const feedbackDuplicated = reportRows >= 2 || reviewRows >= 2;
      const feedbackCorroborated = (
        uniquePrices.length >= 2 ||
        sectionKinds.size >= 2 ||
        distinctUpdateDates >= 2 ||
        iconRows >= 2 ||
        (distinctNames.length >= 2 && uniquePrices.length >= 1 && total >= 1200)
      );
      if (iconRows === 1 && sectionKinds.size <= 1 && uniquePrices.length <= 1 && distinctUpdateDates <= 1) return false;
      if (
        sectionKinds.size === 1 &&
        uniquePrices.length <= 1 &&
        distinctUpdateDates <= 1 &&
        total <= 2000 &&
        distinctNames.length <= 2
      ) {
        return false;
      }
      if (sectionKinds.size === 1 && totalTokenCount >= 2 && uniquePrices.length <= 1 && distinctNames.length <= 1) return false;
      if (sectionKinds.size >= 2) return true;
      if (feedbackDuplicated && feedbackCorroborated) return true;
      if (strongMultiple && count >= 2) return true;
      if (count <= 1 && total >= 1800 && priceTokens >= 3) return true;
    }
  }
  // Mixed Apple commerce sections are frequently under-counted by OCR.
  if ((scopeChecked ? scopeMixedSections : hasMixedSections)) return true;
  if (count >= 3 && hasCommerceHint && (textPriceRows >= 3 || total >= 2500)) return true;
  if (count <= 1 && hasCommerceHint && total >= 1200 && textPriceRows >= 3) return true;
  if (count >= 2 && total >= 2500 && hasCommerceHint && textPriceRows >= 2) {
    if (!/icloud|app[l1]e\s*developer|developer\s*program/i.test(service)) return true;
  }
  return false;
}

function enforceConservativeNaming(fields, fullText = "", ocrData = null) {
  const f = {
    company: fields?.company || "APPLE",
    date: normalizeDateValue(fields?.date || ""),
    amount: normalizeAmountText(fields?.amount || ""),
    itemCount: normalizeItemCount(fields?.itemCount || 1),
    service: normalizeServiceName(fields?.service || ""),
  };
  if (ocrData && f.company === "APPLE") {
    const lines = getLines(ocrData);
    const scope = findAppleCommerceScope(lines);
    const sectionKinds = scope.found ? detectCommerceSectionKinds(lines, scope) : new Set();
    const total = parseMoneyValue(f.amount || "");
    const iconRows = scope.found ? countIconLikeItemRows(lines, scope, f.amount) : 0;
    const uniquePrices = scope.found ? extractPriceCandidatesFromLines(lines, scope, total).length : 0;
    const yearHints = collectYearHints(fullText, lines);
    const distinctUpdateDates = scope.found ? countDistinctUpdateDates(lines, scope, yearHints) : 0;
    const hardMultiAnchors = (
      sectionKinds.size >= 2 ||
      iconRows >= 2
    );
    const lowAmountSingleGuard = (
      total > 0 &&
      total <= 1200 &&
      sectionKinds.size <= 1 &&
      iconRows <= 1 &&
      uniquePrices <= 1
    );
    const serviceIsMultiple = isStrictMultipleServiceLabel(f.service || "");
    const finalHardSingle = hasHardSingleReceiptEvidence(ocrData, fullText, f.amount, f);
    const finalSingleHardRule = hasFinalSingleReceiptEvidence(ocrData, fullText, f.amount, f);
    const finalSingleOverride = resolveSingleItemOverride(ocrData, fullText, f.amount, f);
    if (finalHardSingle || finalSingleHardRule || finalSingleOverride) {
      f.itemCount = 1;
      f.service = sanitizeServiceFinal(
        finalSingleOverride?.service || chooseSingleReceiptService(ocrData, fullText, f.amount, f),
        "appstore"
      );
      return f;
    }
    if (
      lowAmountSingleGuard ||
      (
      serviceIsMultiple &&
      !hardMultiAnchors &&
      iconRows <= 1 &&
      uniquePrices <= 1 &&
      distinctUpdateDates <= 1
      )
    ) {
      f.itemCount = 1;
      f.service = sanitizeServiceFinal(chooseSingleReceiptService(ocrData, fullText, f.amount, f), "appstore");
      return f;
    }
    const hardMulti = hasHardMultipleReceiptEvidence(ocrData, fullText, f.amount, f);
    if (shouldForceMultipleFallback(f, f.amount, fullText, ocrData)) {
      if (hardMulti && hardMultiAnchors) {
        f.service = "複数";
        f.itemCount = normalizeMultipleItemCount(f.itemCount);
      } else {
        f.itemCount = 1;
        f.service = sanitizeServiceFinal(chooseSingleReceiptService(ocrData, fullText, f.amount, f), "appstore");
      }
      return f;
    }
    return f;
  }
  return f;
}

function normalizeMultipleItemCount(value) {
  return Math.max(2, normalizeItemCount(value || 2));
}

function hasCommerceSectionHint(value = "") {
  const t = normalizeDateText(value || "");
  return /app\s*st[o0]re|app[l1]e\s*b[o0]{2}ks|i[t1]unes(\s*store)?|ダウンロード製品|download\s*products?/i.test(t);
}

function hasMixedCommerceSections(value = "") {
  const t = normalizeDateText(value || "");
  const hasAppStore = /app\s*st[o0]re/i.test(t);
  const hasBooks = /app[l1]e\s*b[o0]{2}ks/i.test(t);
  const hasItunes = /i[t1]unes(\s*store)?/i.test(t);
  if (hasAppStore && hasBooks) return true;
  if (hasItunes && (hasAppStore || hasBooks)) return true;
  return false;
}

function countHintRowsDedup(lines, pattern) {
  const ys = [];
  const scale = computeCoordinateScale(lines || [], "y0", 1200, 0.7, 2.8);
  const yTol = Math.max(24, 34 * scale);
  for (const line of lines || []) {
    const raw = normalizeDateText(line?.text || "");
    if (!raw) continue;
    if (!pattern.test(raw)) continue;
    const y = Number(line?.y0 || 0);
    if (ys.some((v) => Math.abs(v - y) <= yTol)) continue;
    ys.push(y);
  }
  return ys.length;
}

function countDistinctUpdateDates(lines, scope, yearHints = []) {
  if (!scope?.found) return 0;
  const out = new Set();
  for (let i = scope.begin; i < scope.end; i += 1) {
    const raw = normalizeDateText(lines[i]?.text || "");
    if (!raw) continue;
    if (!/(更新\s*[：:]|updated\s*:)/i.test(raw)) continue;

    const next = normalizeDateText(lines[i + 1]?.text || "");
    const merged = `${raw} ${next}`.trim();
    const d1 = normalizeDateValue(extractDate(raw, yearHints, { allowTokenFallback: false }));
    const d2 = normalizeDateValue(extractDate(merged, yearHints, { allowTokenFallback: false }));
    const date = isPlausibleDate(d1) ? d1 : (isPlausibleDate(d2) ? d2 : "");
    if (date) out.add(date);
  }
  return out.size;
}

function countIconLikeItemRows(lines, scope, totalAmount = "") {
  if (!scope?.found) return 0;
  const total = parseMoneyValue(totalAmount);
  const scoped = lines.slice(scope.begin, scope.end);
  const yScale = computeCoordinateScale(scoped, "y0", 1200, 0.7, 2.8);
  const scopeBeginY = Number(lines[scope.begin]?.y0 || 0);
  const scopeEndY = Number(lines[scope.end - 1]?.y0 || scopeBeginY + 1);
  const scopeHeight = Math.max(1, scopeEndY - scopeBeginY);
  const bottomStartY = scopeBeginY + scopeHeight * 0.72;
  const priceDupTol = Math.max(12, 18 * yScale);
  const priceItemNearTol = Math.max(24, 44 * yScale);
  const itemDupTol = Math.max(20, 34 * yScale);
  const priceRows = [];

  for (const line of scoped) {
    const raw = normalizeDateText(line?.text || "");
    if (!raw) continue;
    if (/(合計|total|小計|subtotal|jct|tax|請求とお支払い|お支払い情報|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex|jpn)/i.test(raw)) {
      continue;
    }
    const vals = parseMoneyListFromTextAll(raw)
      .map((m) => parseMoneyValue(m))
      .filter((v) => Number.isFinite(v) && v >= 60 && v <= 500000);
    if (!vals.length) continue;
    const onlyTotal = total > 0 && vals.every((v) => Math.abs(v - total) <= 1);
    if (onlyTotal && /(合計|total)/i.test(raw)) continue;
    const y = Number(line?.y0 || 0);
    if (onlyTotal && y >= bottomStartY) continue;
    const dup = priceRows.find((row) => Math.abs(row.y - y) <= priceDupTol);
    if (dup) {
      dup.onlyTotal = dup.onlyTotal && onlyTotal;
      continue;
    }
    priceRows.push({ y, onlyTotal });
  }

  const priceYs = priceRows.map((row) => row.y);
  const itemYs = [];
  for (const line of scoped) {
    const raw = normalizeDateText(line?.text || "");
    if (!raw) continue;
    if (/(app\s*st[o0]re|app[l1]e\s*b[o0]{2}ks|i[t1]unes\s*store|更新|updated|問題を報告|report a problem|レビューを書く|write a review|月額|年額|premium|subscription|合計|total|小計|subtotal|jct|tax|請求とお支払い|お支払い情報|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex|jpn)/i.test(raw)) {
      continue;
    }

    let candidate = "";
    if (/[¥￥$]\s?\d|\d+\s?円/.test(raw)) candidate = extractNameFromMoneyLine(raw);
    if (!candidate) candidate = normalizeProductNameCandidate(raw);
    candidate = normalizeServiceName(candidate);
    if (!candidate) continue;
    if (!isLikelyValidAppCandidate(candidate)) continue;
    if (isLikelyGarbageServiceName(candidate) || isGenericServiceLabel(candidate) || isSummaryLikeName(candidate)) continue;

    const x = Number(line?.x0 || 0);
    if (x > 320) continue;
    const y = Number(line?.y0 || 0);
    if (!priceYs.some((py) => Math.abs(py - y) <= priceItemNearTol)) continue;
    if (itemYs.some((iy) => Math.abs(iy - y) <= itemDupTol)) continue;
    itemYs.push(y);
  }

  if (itemYs.length) return itemYs.length;
  const nonTotalPriceRows = priceRows.filter((row) => !row.onlyTotal).length;
  if (nonTotalPriceRows >= 2) return nonTotalPriceRows;
  if (priceRows.length >= 2 && nonTotalPriceRows === 0) return 1;
  if (priceRows.length === 1) return 1;
  if (nonTotalPriceRows === 1) return 1;
  return 0;
}

function isCommerceSectionStartLine(value = "") {
  const t = normalizeDateText(value || "");
  return /app\s*st[o0]re|app[l1]e\s*b[o0]{2}ks|i[t1]unes(\s*store)?|ダウンロード製品|download\s*products?/i.test(t);
}

function isCommerceFooterStartLine(value = "") {
  const t = normalizeDateText(value || "");
  return /請求とお支払い|お支払い情報|payment|billing|american express|visa|master|amex|jcb|ご注文合計/i.test(t);
}

function extractAppleServiceSummary(text, ocrData, totalAmount = "", fallbackName = "") {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const t = (text || "").toLowerCase();

  // iCloud single invoices are usually one line item.
  if (/icloud\+|icloud/.test(t) && !hasCommerceSectionHint(t)) {
    const storage = extractStorageCapacity(text);
    return {
      service: sanitizeServiceFinal(storage ? `iCloud-${storage}` : "iCloud", "icloud"),
      itemCount: 1,
    };
  }

  const scope = findAppleCommerceScope(lines);
  const items = scope.found ? compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount)) : [];
  if (items.length) {
    const primary = choosePrimaryServiceName(items, lines, scope, text);
    const itemCount = estimateItemCountFromCommerceData(items, lines, words, scope, totalAmount, text);
    const sectionKinds = detectCommerceSectionKinds(lines, scope);
    const total = parseMoneyValue(totalAmount);
    const iconRows = countIconLikeItemRows(lines, scope, totalAmount);
    const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
    const itemPrices = items.map((x) => Number(x.amountValue || 0)).filter((v) => Number.isFinite(v) && v > 0);
    const strongMultiple = hasStrongMultipleEvidenceInScope(lines, scope, totalAmount, items, text);
    const strongSingle = hasSingleItemEvidenceInScope(lines, words, scope, totalAmount, items, text);
    const mixedCount = sectionKinds.size >= 2
      ? estimateMixedSectionItemCount(lines, words, scope, totalAmount, items, text)
      : normalizeItemCount(itemCount);
    if (iconRows === 1 && sectionKinds.size <= 1 && uniquePrices.length <= 1) {
      return {
        service: sanitizeServiceFinal(primary, "appstore"),
        itemCount: 1,
        itemPrices,
      };
    }
    if (strongSingle) {
      return {
        service: sanitizeServiceFinal(primary, "appstore"),
        itemCount: 1,
        itemPrices,
      };
    }
    // Mixed store receipts (e.g. App Store + Apple Books) are OCR-unstable for exact representative name.
    // Use a safe generic label and suppress +N in title to avoid confidently wrong filenames.
    if (sectionKinds.size >= 2 || (looksIncompleteCommerceSummary(lines, scope, totalAmount, itemCount, itemPrices, text) && iconRows !== 1)) {
      return {
        service: "複数",
        itemCount: normalizeMultipleItemCount(sectionKinds.size >= 2 ? mixedCount : itemCount),
        itemPrices,
      };
    }
    if (normalizeItemCount(itemCount) >= 2 && strongMultiple && iconRows !== 1) {
      return {
        service: "複数",
        itemCount: normalizeMultipleItemCount(itemCount),
        itemPrices,
      };
    }
    return {
      service: sanitizeServiceFinal(primary, "appstore"),
      itemCount: strongMultiple ? normalizeItemCount(itemCount) : 1,
      itemPrices,
    };
  }

  const names = extractAppNamesFromStoreBlock(ocrData, text);
  if (names.length) {
    const guessedCount = normalizeItemCount(countStoreLineItems(ocrData, text, totalAmount));
    const iconRows = scope.found ? countIconLikeItemRows(lines, scope, totalAmount) : 0;
    const guessedStrongMultiple = guessedCount >= 2 && hasStrongMultipleEvidenceInScope(
      lines,
      scope,
      totalAmount,
      scope.found ? compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount)) : [],
      text
    );
    const total = parseMoneyValue(totalAmount);
    const normText = normalizeDateText(text || "");
    if (iconRows === 1) {
      return {
        service: sanitizeServiceFinal(names[0], "appstore"),
        itemCount: 1,
        itemPrices: [],
      };
    }
    if (guessedStrongMultiple) {
      return {
        service: "複数",
        itemCount: normalizeMultipleItemCount(guessedCount),
        itemPrices: [],
      };
    }
    if (
      hasCommerceSectionHint(normText) &&
      ((total >= 2500 && guessedCount <= 3) || (/app[l1]e\s*b[o0]{2}ks|i[t1]unes/i.test(normText) && guessedCount <= 3 && total >= 1200))
    ) {
      return {
        service: "複数",
        itemCount: normalizeMultipleItemCount(guessedCount),
        itemPrices: [],
      };
    }
    return {
      service: sanitizeServiceFinal(names[0], "appstore"),
      itemCount: guessedStrongMultiple ? guessedCount : 1,
      itemPrices: [],
    };
  }

  const known = canonicalServiceName(text);
  if (known) {
    return {
      service: sanitizeServiceFinal(displayServiceName(known), "general"),
      itemCount: 1,
      itemPrices: [],
    };
  }

  const fallback = normalizeServiceName(removeExt(fallbackName));
  return {
    service: sanitizeServiceFinal(fallback || "AppleService", "general"),
    itemCount: 1,
    itemPrices: [],
  };
}

function findAppleCommerceScope(lines) {
  const start = lines.findIndex((line) => isCommerceSectionStartLine(line.text || ""));
  if (start < 0) return { found: false, begin: 0, end: 0 };
  const begin = start;
  let end = lines.length;
  for (let i = begin + 1; i < lines.length; i += 1) {
    if (isCommerceFooterStartLine(lines[i].text || "")) {
      end = i;
      break;
    }
  }
  return { found: true, begin, end };
}

function detectCommerceSectionKinds(lines, scope) {
  const kinds = new Set();
  if (!scope?.found) return kinds;
  for (let i = scope.begin; i < scope.end; i += 1) {
    const raw = normalizeDateText(lines[i]?.text || "");
    if (!raw) continue;
    if (/app\s*st[o0]re/i.test(raw)) kinds.add("appstore");
    if (/app[l1]e\s*b[o0]{2}ks/i.test(raw)) kinds.add("books");
    if (/i[t1]unes\s*store/i.test(raw)) kinds.add("itunes");
  }
  return kinds;
}

function looksIncompleteCommerceSummary(lines, scope, totalAmount = "", itemCount = 0, itemPrices = [], fullText = "") {
  if (!scope?.found) return false;
  const total = parseMoneyValue(totalAmount);
  const count = normalizeItemCount(itemCount);
  const vals = (itemPrices || []).map((v) => Number(v)).filter((v) => Number.isFinite(v) && v > 0);
  const totalMatchCount = total > 0 ? vals.filter((v) => Math.abs(v - total) <= 1).length : 0;
  const valsNoTotal = total > 0 ? vals.filter((v) => Math.abs(v - total) > 1) : vals;
  const sumNoTotal = valsNoTotal.reduce((a, b) => a + b, 0);
  // A single item can naturally equal the grand total. Treat as suspicious only when total-like
  // amount appears multiple times in extracted rows.
  const hasLikelyTotalDuplicate = totalMatchCount >= 2;
  const scopeText = lines.slice(scope.begin, scope.end).map((line) => normalizeDateText(line.text || "")).join(" ");
  const hasBooksLike = /app[l1]e\s*b[o0]{2}ks|i[t1]unes\s*store/i.test(scopeText);
  const hasAppStore = /app\s*st[o0]re/i.test(scopeText);
  const priceTokens = countPriceTokensFromLines(lines, scope, total);

  // Guard: duplicated extraction of the same single line item (often equal to total)
  // should not be promoted to multi-item.
  if (hasLikelyTotalDuplicate && valsNoTotal.length === 0) return false;
  if (hasLikelyTotalDuplicate && count >= 2 && valsNoTotal.length >= 1) return true;
  if (hasLikelyTotalDuplicate && valsNoTotal.length <= 1 && total >= 1000 && (priceTokens >= 2 || hasBooksLike)) return true;
  if (total > 0 && valsNoTotal.length >= 2) {
    if (count >= 2 && sumNoTotal <= total * 0.8 && count <= valsNoTotal.length + 1) return true;
    if (count === 1 && sumNoTotal <= total * 0.55) return true;
    const residual = total - sumNoTotal;
    const minPart = valsNoTotal.length ? Math.min(...valsNoTotal) : 0;
    if (count <= 3 && residual >= Math.max(200, minPart || 200)) return true;
  }
  if (total >= 2500 && count <= 3 && hasAppStore) return true;
  if (hasBooksLike && hasAppStore && count <= 4) return true;
  if (hasBooksLike && count <= 3 && total >= 1200) return true;
  if (priceTokens >= count + 3 && count <= 4 && (valsNoTotal.length >= 2 || (hasBooksLike && total >= 1200))) return true;
  return false;
}

function hasStrongMultipleEvidenceInScope(lines, scope, totalAmount = "", items = [], fullText = "") {
  if (!scope?.found) return hasMixedCommerceSections(fullText);

  const sectionKinds = detectCommerceSectionKinds(lines, scope);
  if (sectionKinds.size >= 2) return true;

  const scoped = lines.slice(scope.begin, scope.end);
  const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
  const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
  const feedbackDuplicated = reportRows >= 2 || reviewRows >= 2;

  const distinctNames = extractDistinctItemNames(items || []);
  const total = parseMoneyValue(totalAmount);
  const totalTokenCount = countTotalTokenOccurrencesInScope(lines, scope, totalAmount);
  const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
  const yearHints = collectYearHints(fullText, lines);
  const distinctUpdateDates = countDistinctUpdateDates(lines, scope, yearHints);
  const iconRows = countIconLikeItemRows(lines, scope, totalAmount);
  const priceTokens = countPriceTokensFromLines(lines, scope, total);
  if (sectionKinds.size === 1 && totalTokenCount >= 2 && uniquePrices.length <= 1 && distinctNames.length <= 1) return false;
  if (iconRows === 1 && uniquePrices.length <= 1 && distinctUpdateDates <= 1) return false;
  if (iconRows >= 2 && (feedbackDuplicated || uniquePrices.length >= 2 || distinctUpdateDates >= 2 || distinctNames.length >= 2)) return true;
  if (feedbackDuplicated && (uniquePrices.length >= 2 || distinctUpdateDates >= 2)) return true;
  if (distinctNames.length >= 2 && (uniquePrices.length >= 2 || distinctUpdateDates >= 2)) return true;
  if (uniquePrices.length >= 2) return true;

  return false;
}

function hasSingleItemEvidenceInScope(lines, words, scope, totalAmount = "", items = [], fullText = "") {
  if (!scope?.found) return false;

  const sectionKinds = detectCommerceSectionKinds(lines, scope);
  if (sectionKinds.size >= 2) return false;

  const scoped = lines.slice(scope.begin, scope.end);
  const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
  const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
  if (reportRows >= 2 || reviewRows >= 2) return false;

  const total = parseMoneyValue(totalAmount);
  const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total);
  if (uniquePrices.length >= 2) return false;

  const distinctNames = extractDistinctItemNames(items || []);
  if (distinctNames.length >= 2) return false;

  const itemRows = normalizeItemCount((items || []).length || 1);
  if (itemRows >= 2) return false;

  const amountWordRows = countAmountRowsFromWords(lines, words, scope, totalAmount);
  if (amountWordRows >= 3) return false;

  return true;
}

function extractDistinctItemNames(items = []) {
  const out = [];
  for (const item of items || []) {
    const name = normalizeProductNameCandidate(item?.name || "");
    if (!name) continue;
    if (isLikelyGarbageServiceName(name) || isGenericServiceLabel(name) || isSummaryLikeName(name)) continue;
    const key = normalizeLineDedupKey(name);
    if (!key) continue;
    const dup = out.find((row) => row.key === key || row.key.includes(key) || key.includes(row.key));
    if (!dup) {
      out.push({ name, key });
      continue;
    }
    if (name.length > dup.name.length) {
      dup.name = name;
      dup.key = key;
    }
  }
  return out.map((row) => row.name);
}

function estimateMixedSectionItemCount(lines, words, scope, totalAmount = "", items = [], fullText = "") {
  if (!scope?.found) return normalizeItemCount((items || []).length || 2);
  const total = parseMoneyValue(totalAmount);
  const scoped = lines.slice(scope.begin, scope.end);
  const reportRows = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
  const reviewRows = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
  const amountRows = countAmountRowsFromWords(lines, words, scope, totalAmount);
  const linePriceRows = countPriceTokensFromLines(lines, scope, total);
  const uniquePrices = extractPriceCandidatesFromLines(lines, scope, total).length;
  const distinctNames = extractDistinctItemNames(items || []).length;
  const mentionCount = Math.min(
    findCanonicalServicesInText(fullText).filter((x) => !["APP_STORE", "ICLOUD"].includes(x)).length,
    9
  );

  const anchors = [
    normalizeItemCount((items || []).length || 1),
    normalizeItemCount(reportRows),
    normalizeItemCount(reviewRows),
    normalizeItemCount(amountRows),
    normalizeItemCount(distinctNames),
    normalizeItemCount(mentionCount),
  ].filter((n) => Number.isFinite(n) && n >= 1);

  let base = anchors.length ? Math.max(...anchors) : 1;
  if (base < 2) base = 2;

  // Mixed sections often have repeated prices; use token rows as an upper correction only.
  if (linePriceRows >= base + 1 && uniquePrices >= 2) {
    base = Math.max(base, Math.min(9, linePriceRows));
  }

  return normalizeItemCount(base);
}

function extractAppleLineItems(lines, words, scope, totalAmount = "") {
  const totalValue = parseMoneyValue(totalAmount);
  const beginY = lines[scope.begin]?.y0 ?? 0;
  const endY = lines[scope.end - 1]?.y0 ?? Infinity;
  const scopeHeight = Math.max(1, endY - beginY);
  const bottomStartY = beginY + scopeHeight * 0.72;
  const maxX = Math.max(1, ...words.map((w) => w.x0 || 0), ...lines.map((l) => l.x0 || 0));

  const rows = [];
  for (const w of words) {
    if ((w.y0 || 0) < beginY || (w.y0 || 0) > endY + 18) continue;
    const amountValue = parseAmountValueFromWord(w.text, w.x0 || 0, maxX);
    if (amountValue <= 0) continue;
    if (amountValue > 500000) continue;
    if (amountValue < 60) continue;

    const nearLine = findNearestLineByY(lines, w.y0 || 0);
    const nearText = normalizeDateText(nearLine?.text || "");
    if (/(jct|tax|小計|subtotal|合計|total|含む|課税|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex)/i.test(nearText)) {
      continue;
    }
    if (
      totalValue > 0 &&
      Math.abs(amountValue - totalValue) <= 1 &&
      ((w.y0 || 0) >= bottomStartY || /ご注文合計|合計|total|小計|subtotal/i.test(nearText))
    ) {
      continue;
    }

    const existing = rows.find((r) => Math.abs(r.y - (w.y0 || 0)) <= 16);
    if (existing) {
      if (amountValue > existing.amountValue) {
        existing.amountValue = amountValue;
        existing.priceX = w.x0 || existing.priceX;
      }
      continue;
    }
    rows.push({ y: w.y0 || 0, priceX: w.x0 || 0, amountValue, name: "" });
  }

  rows.sort((a, b) => a.y - b.y);
  const out = [];
  for (const row of rows) {
    const name = extractPrimaryNameForRow(lines, words, row.y, row.priceX, scope);
    out.push({ ...row, name });
  }
  return out;
}

function compactLineItems(items) {
  const sorted = [...(items || [])].sort((a, b) => (a.y || 0) - (b.y || 0));
  const out = [];
  for (const item of sorted) {
    const name = normalizeServiceName(item.name || "").toLowerCase();
    const nameKey = normalizeLineDedupKey(name);
    const isNameless = !name;
    const dup = out.find((x) => {
      const xn = normalizeServiceName(x.name || "").toLowerCase();
      if (Math.abs((x.amountValue || 0) - (item.amountValue || 0)) > 1) return false;
      const dy = Math.abs((x.y || 0) - (item.y || 0));
      const xnKey = normalizeLineDedupKey(xn);
      const sameName = Boolean(nameKey && xnKey && (nameKey === xnKey || nameKey.includes(xnKey) || xnKey.includes(nameKey)));
      if (sameName && dy <= 38) return true;
      if ((!nameKey || !xnKey) && dy <= 20) return true;
      return false;
    });
    if (dup && !isNameless) continue;
    out.push(item);
  }

  // Remove nameless footer/summary rows that share the same amount as a named item.
  const namedRows = out.filter((x) => normalizeServiceName(x.name || ""));
  const cleaned = out.filter((x) => {
    const hasName = normalizeServiceName(x.name || "");
    if (hasName) return true;
    const sameAmountNamed = namedRows.find((n) => Math.abs((n.amountValue || 0) - (x.amountValue || 0)) <= 1);
    if (sameAmountNamed) return false;
    return true;
  });

  // Remove summary-like rows (e.g. 合計/Total) when the same amount has a richer item row.
  const filtered = cleaned.filter((x) => {
    const name = normalizeServiceName(x.name || "");
    if (!name || !isSummaryLikeName(name)) return true;
    const sameAmountReal = cleaned.find((n) => (
      n !== x &&
      Math.abs((n.amountValue || 0) - (x.amountValue || 0)) <= 1 &&
      normalizeServiceName(n.name || "") &&
      !isSummaryLikeName(n.name || "")
    ));
    return !sameAmountReal;
  });
  return filtered;
}

function parseAmountValueFromWord(wordText, x0 = 0, maxX = 1) {
  const money = parseMoneyFromText(wordText || "");
  if (money) return parseMoneyValue(money);
  const t = normalizeDateText(wordText || "");
  if (/^\d{2,5}$/.test(t) && x0 > maxX * 0.62) return Number(t);
  return 0;
}

function findNearestLineByY(lines, y) {
  let best = null;
  let bestDist = Infinity;
  for (const line of lines) {
    const d = Math.abs((line.y0 || 0) - y);
    if (d < bestDist) {
      best = line;
      bestDist = d;
    }
  }
  return best;
}

function extractPrimaryNameForRow(lines, words, rowY, priceX, scope) {
  const nearby = lines
    .filter((line, idx) => idx >= scope.begin && idx < scope.end && Math.abs((line.y0 || 0) - rowY) <= 24)
    .map((line) => normalizeDateText(line.text || ""))
    .filter(Boolean);

  const candidates = [];
  for (const raw of nearby) {
    let c = raw
      .replace(/([¥￥]\s?\d[\d,]*(?:\.\d+)?|\$\s?\d[\d,]*(?:\.\d+)?|\d[\d,]*(?:\.\d+)?\s?円)/gi, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!c) continue;
    if (/(更新|updated|問題を報告|report a problem|レビューを書く|review|author|著者|iphone|ipad|app[l1]e\s*id|app[l1]e\s*acc[o0]unt|注文番号|書類番号|請求先|ブック)$/i.test(c)) {
      continue;
    }
    c = normalizeProductNameCandidate(c);
    if (!c || isLikelyGarbageServiceName(c) || isGenericServiceLabel(c) || isSummaryLikeName(c)) continue;
    candidates.push(c);
  }
  if (candidates.length) return candidates[0];

  const leftWords = words
    .filter((w) => Math.abs((w.y0 || 0) - rowY) <= 18 && (w.x0 || 0) < priceX - 18)
    .sort((a, b) => (a.x0 || 0) - (b.x0 || 0))
    .map((w) => w.text || "")
    .filter(Boolean);
  if (!leftWords.length) return "";
  const joined = normalizeProductNameCandidate(leftWords.join(" "));
  if (!joined || isLikelyGarbageServiceName(joined) || isGenericServiceLabel(joined) || isSummaryLikeName(joined)) return "";
  return joined;
}

function choosePrimaryServiceName(items, lines, scope, fullText = "") {
  const named = items
    .map((x) => normalizeProductNameCandidate(x.name || ""))
    .filter((x) => x && !isLikelyGarbageServiceName(x) && !isGenericServiceLabel(x) && !isSummaryLikeName(x));
  if (named.length) return named[0];

  const scopeText = lines.slice(scope.begin, scope.end).map((line) => line.text || "").join(" ");
  const hits = findCanonicalServicesInText(`${scopeText} ${fullText}`)
    .filter((x) => !["APP_STORE", "ICLOUD"].includes(x))
    .map((x) => displayServiceName(x));
  if (hits.length) return hits[0];
  return "AppStore";
}

function estimateItemCountFromCommerceData(items, lines, words, scope, totalAmount = "", fullText = "") {
  const total = parseMoneyValue(totalAmount);
  const pricesFromRows = items.map((x) => Number(x.amountValue || 0)).filter((v) => Number.isFinite(v) && v > 0);
  const byRows = items.length;
  const scoped = lines.slice(scope.begin, scope.end);
  const byReport = countHintRowsDedup(scoped, /問題を報告|report a problem/i);

  const byLinePrices = extractPriceCandidatesFromLines(lines, scope, total).length;
  const byPriceTokens = countPriceTokensFromLines(lines, scope, total);
  const byWordAmountRows = countAmountRowsFromWords(lines, words, scope, total);
  const allPrices = collectAllItemPriceCandidates(lines, scope, total);
  const byTextMention = Math.min(
    findCanonicalServicesInText(fullText).filter((x) => !["APP_STORE", "ICLOUD"].includes(x)).length,
    9
  );

  const byFit = inferCountFromTotalFit(total, allPrices.length ? allPrices : pricesFromRows);
  const missingByResidual = inferMissingCountFromResidual(total, pricesFromRows);
  const byResidual = missingByResidual >= 1 ? byRows + missingByResidual : 0;

  const signals = [byRows, byReport, byLinePrices, byPriceTokens, byWordAmountRows, byFit, byResidual, byTextMention]
    .filter((n) => Number.isFinite(n) && n >= 1 && n <= 24);
  if (!signals.length) return 1;

  const candidates = uniqueNonEmpty(signals.map(String)).map(Number).sort((a, b) => b - a);
  for (const candidate of candidates) {
    const tolerance = candidate >= 4 ? 1 : 0;
    let support = 0;
    let structuralSupport = 0;
    if (byRows >= 1 && Math.abs(byRows - candidate) <= tolerance) support += 1;
    if (byRows >= 2 && Math.abs(byRows - candidate) <= tolerance) structuralSupport += 1;
    if (byReport >= 1 && Math.abs(byReport - candidate) <= tolerance) support += 1;
    if (byReport >= 2 && Math.abs(byReport - candidate) <= tolerance) structuralSupport += 1;
    if (byLinePrices >= 1 && Math.abs(byLinePrices - candidate) <= tolerance) support += 1;
    if (byPriceTokens >= 1 && Math.abs(byPriceTokens - candidate) <= tolerance) support += 1;
    if (byWordAmountRows >= 1 && Math.abs(byWordAmountRows - candidate) <= tolerance) support += 1;
    if (byFit >= 2 && Math.abs(byFit - candidate) <= tolerance) support += 1;
    if (byResidual >= 2 && Math.abs(byResidual - candidate) <= tolerance) support += 1;
    if (byTextMention >= 2 && Math.abs(byTextMention - candidate) <= tolerance) support += 1;
    if (byTextMention >= 2 && Math.abs(byTextMention - candidate) <= tolerance) structuralSupport += 1;
    if (support >= 2 && structuralSupport >= 1) return normalizeItemCount(candidate);
  }

  // If row extraction under-detected but total residual strongly explains missing rows, allow expansion.
  if (byRows >= 2 && byResidual >= byRows + 2 && byResidual <= 24) {
    return normalizeItemCount(byResidual);
  }

  if (byRows >= 2) return normalizeItemCount(byRows);
  if (byReport >= 2 && (byLinePrices >= 2 || byWordAmountRows >= 2 || byTextMention >= 2)) {
    return normalizeItemCount(byReport);
  }
  if (byReport >= 3 && byRows >= 1) return normalizeItemCount(byReport);
  if (byTextMention >= 2) return normalizeItemCount(byTextMention);
  if (byLinePrices >= 2 && byWordAmountRows >= 2) return normalizeItemCount(Math.max(byLinePrices, byWordAmountRows));
  if (byRows >= 1) return 1;
  if (byFit >= 2) return normalizeItemCount(byFit);
  return 1;
}

function extractPriceCandidatesFromLines(lines, scope, totalValue = 0) {
  const vals = [];
  for (let i = scope.begin; i < scope.end; i += 1) {
    const raw = normalizeDateText(lines[i]?.text || "");
    if (!raw) continue;
    if (/(jct|tax|小計|subtotal|合計|total|含む|課税|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex)/i.test(raw)) {
      continue;
    }
    for (const money of parseMoneyListFromText(raw)) {
      const v = parseMoneyValue(money);
      if (!v) continue;
      if (totalValue > 0 && Math.abs(v - totalValue) <= 1) continue;
      vals.push(v);
    }
  }
  return uniqueNonEmpty(vals.map(String)).map(Number);
}

function collectPriceTokensFromLines(lines, scope, totalValue = 0) {
  const values = [];
  for (let i = scope.begin; i < scope.end; i += 1) {
    const raw = normalizeDateText(lines[i]?.text || "");
    if (!raw) continue;
    if (/(jct|tax|小計|subtotal|合計|total|含む|課税|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex)/i.test(raw)) {
      continue;
    }
    const list = parseMoneyListFromTextAll(raw);
    for (const money of list) {
      const v = parseMoneyValue(money);
      if (!v) continue;
      if (totalValue > 0 && Math.abs(v - totalValue) <= 1) continue;
      if (v < 60 || v > 500000) continue;
      values.push(v);
    }
  }
  return values;
}

function countPriceTokensFromLines(lines, scope, totalValue = 0) {
  return collectPriceTokensFromLines(lines, scope, totalValue).length;
}

function countTotalTokenOccurrencesInScope(lines, scope, totalAmount = "") {
  if (!scope?.found) return 0;
  const total = parseMoneyValue(totalAmount);
  if (!total) return 0;
  let count = 0;
  for (let i = scope.begin; i < scope.end; i += 1) {
    const raw = normalizeDateText(lines[i]?.text || "");
    if (!raw) continue;
    const list = parseMoneyListFromTextAll(raw);
    for (const money of list) {
      const v = parseMoneyValue(money);
      if (!v) continue;
      if (Math.abs(v - total) <= 1) count += 1;
    }
  }
  return count;
}

function countAmountRowsFromWords(lines, words, scope, totalAmount = "") {
  if (!scope?.found) return 0;
  const totalValue = parseMoneyValue(totalAmount);
  const beginY = lines[scope.begin]?.y0 ?? 0;
  const endY = lines[scope.end - 1]?.y0 ?? Infinity;
  const scopeHeight = Math.max(1, endY - beginY);
  const bottomStartY = beginY + scopeHeight * 0.72;
  const maxX = Math.max(1, ...words.map((w) => w.x0 || 0), ...lines.map((l) => l.x0 || 0));

  const ys = [];
  for (const w of words) {
    const y = w.y0 || 0;
    const x = w.x0 || 0;
    if (y < beginY || y > endY + 18) continue;
    if (x < maxX * 0.58) continue;

    const amountValue = parseAmountValueFromWord(w.text || "", x, maxX);
    if (!amountValue || amountValue < 60 || amountValue > 500000) continue;

    const nearLine = findNearestLineByY(lines, y);
    const nearText = normalizeDateText(nearLine?.text || "");
    if (/(jct|tax|小計|subtotal|合計|total|含む|課税|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex)/i.test(nearText)) {
      continue;
    }
    if (
      totalValue > 0 &&
      Math.abs(amountValue - totalValue) <= 1 &&
      (y >= bottomStartY || /ご注文合計|合計|total|小計|subtotal/i.test(nearText))
    ) {
      continue;
    }
    ys.push(y);
  }

  if (!ys.length) return 0;
  ys.sort((a, b) => a - b);
  let rows = 0;
  let lastY = -Infinity;
  for (const y of ys) {
    if (Math.abs(y - lastY) > 20) {
      rows += 1;
      lastY = y;
    }
  }
  return rows;
}

function chooseBestTotalAmount(rawAmount, itemPrices = []) {
  const rawValue = parseMoneyValue(rawAmount);
  const symbol = String(rawAmount || "").trim().startsWith("$") ? "$" : "¥";
  const prices = (itemPrices || []).filter((v) => Number.isFinite(v) && v > 0 && v < 500000);
  if (!prices.length) return rawAmount;

  const sum = prices.reduce((a, b) => a + b, 0);
  const max = Math.max(...prices);
  if (!rawValue) return normalizeAmountText(`${symbol}${sum}`);

  // If OCR selected one line item as total, prefer reconstructed sum.
  if (prices.length >= 2 && (rawValue === max || rawValue < Math.round(sum * 0.75))) {
    return normalizeAmountText(`${symbol}${sum}`);
  }

  // Keep raw when it is close to reconstructed sum.
  if (Math.abs(rawValue - sum) <= Math.max(5, Math.round(sum * 0.03))) {
    return normalizeAmountText(`${symbol}${rawValue}`);
  }

  return normalizeAmountText(`${symbol}${rawValue}`);
}

function collectAllItemPriceCandidates(lines, scope, totalValue = 0) {
  const vals = [];
  for (let i = scope.begin; i < scope.end; i += 1) {
    const raw = normalizeDateText(lines[i]?.text || "");
    if (!raw) continue;
    if (/(jct|tax|小計|subtotal|合計|total|含む|課税|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex)/i.test(raw)) {
      continue;
    }
    const moneyList = parseMoneyListFromTextAll(raw);
    for (const money of moneyList) {
      const v = parseMoneyValue(money);
      if (!v) continue;
      if (totalValue > 0 && Math.abs(v - totalValue) <= 1) continue;
      if (v < 80) continue;
      vals.push(v);
    }
  }
  return vals;
}

function inferCountFromTotalFit(total, prices) {
  if (!total || !prices.length) return 0;
  const list = prices
    .map((v) => Number(v))
    .filter((v) => Number.isFinite(v) && v > 0 && v <= total && v >= 80)
    .sort((a, b) => b - a)
    .slice(0, 18);
  if (!list.length) return 0;

  let best = 0;
  function dfs(start, sum, count) {
    if (sum === total) {
      best = Math.max(best, count);
      return;
    }
    if (sum > total || count >= 9) return;
    for (let i = start; i < list.length; i += 1) {
      dfs(i + 1, sum + list[i], count + 1);
    }
  }
  dfs(0, 0, 0);
  return best;
}

function inferMissingCountFromResidual(total, knownPrices = []) {
  if (!total) return 0;
  const values = (knownPrices || []).map((v) => Number(v)).filter((v) => Number.isFinite(v) && v > 0 && v < total);
  if (!values.length) return 0;
  const knownSum = values.reduce((a, b) => a + b, 0);
  const residual = total - knownSum;
  if (residual <= 0) return 0;

  const freq = new Map();
  for (const v of values) {
    if (v < 80 || v > 500000) continue;
    freq.set(v, (freq.get(v) || 0) + 1);
  }
  const dominant = [...freq.entries()].sort((a, b) => (b[1] - a[1]) || (b[0] - a[0]));
  for (const [price, f] of dominant) {
    if (f < 2) continue;
    if (residual % price !== 0) continue;
    const k = residual / price;
    if (Number.isInteger(k) && k >= 1 && k <= 12) return k;
  }

  const uniq = uniqueNonEmpty(values.map(String))
    .map(Number)
    .filter((v) => Number.isFinite(v) && v >= 80 && v <= 2000)
    .sort((a, b) => b - a);
  if (!uniq.length) return 0;

  let best = 0;
  const tolerance = Math.max(2, Math.round(total * 0.002));
  function dfs(sum, count, depth) {
    if (Math.abs(sum - residual) <= tolerance && count >= 1) {
      best = Math.max(best, count);
    }
    if (depth >= 8 || sum > residual + tolerance) return;
    for (const p of uniq) {
      dfs(sum + p, count + 1, depth + 1);
    }
  }
  dfs(0, 0, 0);
  return best;
}

function normalizeProductNameCandidate(value) {
  let v = normalizeAppName(value || "");
  v = v
    .replace(/^\d+\s*/g, "")
    .replace(/\b(?:app\s*内課金|in[\s-]*app\s*purchase)\b/gi, " ")
    .replace(/\b(?:iphone|ipad|ipod)\b/gi, " ")
    .replace(/\b(?:レビューを書く|review|問題を報告する?)\b/gi, " ")
    .replace(/\b(?:book|books|ブック)\b$/i, "")
    .replace(/\s+/g, " ")
    .trim();
  v = normalizeLatinTokenCase(v);
  return shortenServiceName(trimServiceSuffixes(v));
}

function isSummaryLikeName(value) {
  const v = normalizeDateText(value || "").replace(/\s+/g, "").toLowerCase();
  if (!v) return false;
  return /^(合計|小計|ご注文合計|total|subtotal|jct|tax|課税|含む|included)$/.test(v);
}

function isGenericServiceLabel(value) {
  const v = normalizeDateText(value || "").replace(/\s+/g, "");
  if (!v) return true;
  if (/^a*app$|^a*app内課金$|^内課金$/.test(v.toLowerCase())) return true;
  if (/^app[l1]ebooks?$|^books?$/.test(v.toLowerCase())) return true;
  return false;
}

function isDateLabelLike(value = "") {
  const t = normalizeDateText(value || "");
  if (!t) return false;
  return /(日[付什]|[曰目白][付什]|注文[日曰目白]|ご注文[日曰目白]|order\s*date|date)/i.test(t);
}

function buildPlausibleYmd(year, month, day) {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return "";
  const out = `${y}${pad2(m)}${pad2(d)}`;
  return isPlausibleDate(out) ? out : "";
}

function extractAppleBillingDate(text, ocrData) {
  const lines = getLines(ocrData);
  const pageBottom = lines.reduce((max, line) => Math.max(max, line.y0 || 0), 1);
  const topLimit = pageBottom * 0.42;
  const labelBottomLimit = pageBottom * 0.62;
  const candidates = [];
  const yearHints = collectYearHints(text, lines);

  // Priority 1: explicit date labels in top area.
  for (let i = 0; i < lines.length; i += 1) {
    const row = lines[i];
    const line = normalizeDateText(row.text);
    if (!isDateLabelLike(line)) continue;
    if ((row.y0 || 0) > labelBottomLimit) continue;

    const d1 = extractDate(line, yearHints, { allowTokenFallback: true });
    if (isPlausibleDate(d1)) candidates.push({ date: d1, score: 100 - (row.y0 || 0) * 0.01 });

    const around = [lines[i - 1], lines[i], lines[i + 1], lines[i + 2], lines[i + 3]]
      .map((x) => normalizeDateText(x?.text || ""))
      .filter(Boolean)
      .join(" ");
    const dAround = extractDate(around, yearHints, { allowTokenFallback: false });
    if (isPlausibleDate(dAround)) candidates.push({ date: dAround, score: 96 - (row.y0 || 0) * 0.01 });

    const next1 = extractDate(normalizeDateText(lines[i + 1]?.text || ""), yearHints, { allowTokenFallback: false });
    if (isPlausibleDate(next1)) candidates.push({ date: next1, score: 92 - (row.y0 || 0) * 0.01 });

    const next2 = extractDate(normalizeDateText(lines[i + 2]?.text || ""), yearHints, { allowTokenFallback: false });
    if (isPlausibleDate(next2)) candidates.push({ date: next2, score: 84 - (row.y0 || 0) * 0.01 });
  }

  // Priority 2: first valid top dates.
  for (const row of lines) {
    if ((row.y0 || 0) > topLimit) continue;
    const d = extractDate(row.text, yearHints, { allowTokenFallback: false });
    if (isPlausibleDate(d)) candidates.push({ date: d, score: 70 - (row.y0 || 0) * 0.01 });
  }

  const topJoined = lines
    .filter((row) => (row.y0 || 0) <= topLimit)
    .slice(0, 14)
    .map((row) => normalizeDateText(row.text || ""))
    .join(" ");
  const topDate = extractDate(topJoined, yearHints, { allowTokenFallback: false });
  if (isPlausibleDate(topDate)) candidates.push({ date: topDate, score: 62 });

  const fallback = extractDate(text, yearHints, { allowTokenFallback: false });
  if (isPlausibleDate(fallback)) candidates.push({ date: fallback, score: 40 });
  if (!candidates.length) return "";

  candidates.sort((a, b) => b.score - a.score);
  return candidates[0].date;
}

function extractAppleTotalAmount(text, ocrData) {
  const lines = getLines(ocrData);
  const pageBottom = lines.reduce((max, line) => Math.max(max, line.y0 || 0), 1);
  const candidates = [];
  const subtotalVals = [];
  const taxVals = [];

  for (let i = 0; i < lines.length; i += 1) {
    const raw = lines[i].text || "";
    const nextRaw = lines[i + 1]?.text || "";
    const moneyList = parseMoneyListFromText(raw);
    const hasTotal = /ご注文合計|合計|total/i.test(raw);
    const hasCard = /american express|visa|master|amex|jcb/i.test(raw);
    const hasTax = /jct|消費税|税|tax/i.test(raw);
    const hasSubtotal = /小計|subtotal/i.test(raw);
    const hasNoise = /更新|問題を報告|月額|年額|subscription/i.test(raw);
    const yRatio = Math.min((lines[i].y0 || 0) / pageBottom, 1);

    if (hasTotal && !moneyList.length) {
      for (const m of parseMoneyListFromText(nextRaw)) {
        candidates.push({ amount: m, score: 250, value: parseMoneyValue(m), row: i });
      }
    }

    for (const money of moneyList) {
      const value = parseMoneyValue(money);
      let score = 0;
      if (hasTotal) score += 170;
      if (hasCard) score += 140;
      if (hasTax && !hasTotal) score -= 120;
      if (hasSubtotal && !hasTotal) score -= 95;
      if (hasNoise) score -= 55;
      if (/含む|included/i.test(raw) && !hasTotal) score -= 45;
      if (/(jct|tax)\s*\(?\s*10%/i.test(raw) && !hasTotal) score -= 35;
      if (value < 80 && (hasTax || hasSubtotal)) score -= 50;
      if (yRatio > 0.58) score += 22;
      if (yRatio < 0.2) score -= 10;
      score += Math.min(value, 99999) / 1300;
      candidates.push({ amount: money, score, value, row: i });

      if (hasSubtotal) subtotalVals.push(value);
      if (hasTax) taxVals.push(value);
    }
  }

  if (!candidates.length) {
    return parseMoneyFromText(text);
  }

  // Consistency bonus: if subtotal + tax is close to candidate total, boost it.
  for (const c of candidates) {
    if (!subtotalVals.length || !taxVals.length) continue;
    const near = isNearSubtotalTaxSum(c.value, subtotalVals, taxVals);
    if (near) c.score += 35;
  }

  const strong = candidates.filter((c) => c.score >= 175);
  const pool = strong.length ? strong : candidates;
  pool.sort((a, b) => b.score - a.score || b.value - a.value);
  return pool[0].amount;
}

function extractAppleItem(text, fallbackName, ocrData, totalAmount = "", forcedItemCount = 0) {
  const t = (text || "").toLowerCase();
  if (/ご注文ありがとうございます|ダウンロード製品|app[l1]e\s+store/i.test(t)) {
    const orderName = extractOrderProductName(ocrData, text);
    if (orderName) return sanitizeServiceFinal(orderName, "appstore");
  }

  if (/icloud\+|icloud/.test(t)) {
    const storage = extractStorageCapacity(text);
    return sanitizeServiceFinal(storage ? `iCloud-${storage}` : "iCloud", "icloud");
  }

  if (hasCommerceSectionHint(t)) {
    const names = extractAppNamesFromStoreBlock(ocrData, text);
    const itemCount = forcedItemCount || countStoreLineItems(ocrData, text, totalAmount);
    if (names.length) return sanitizeServiceFinal(summarizeServiceNames(names, itemCount), "appstore");
    const serviceHint = findCanonicalServicesInText(text).find((key) => !["APP_STORE", "ICLOUD"].includes(key));
    if (serviceHint) return sanitizeServiceFinal(displayServiceName(serviceHint), "appstore");
    return "AppStore";
  }

  const known = canonicalServiceName(text);
  if (known) return sanitizeServiceFinal(displayServiceName(known), "general");
  const fallback = normalizeServiceName(removeExt(fallbackName));
  if (isScreenshotLike(fallback)) return "AppleService";
  return sanitizeServiceFinal(fallback || "AppleService", "general");
}

function extractAppNamesFromStoreBlock(ocrData, text) {
  const lines = getLines(ocrData);
  const out = [];
  const seen = new Set();

  function pushCandidate(name, score = 0) {
    const clean = normalizeAppName(name);
    if (!clean) return;
    if (isLikelyGarbageServiceName(clean)) return;
    const key = clean.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ name: clean, score });
  }

  const dictionaryHits = findCanonicalServicesInText(text)
    .map((canonical) => displayServiceName(canonical))
    .filter((name) => !["iCloud", "AppStore"].includes(name));
  for (const hit of dictionaryHits) pushCandidate(hit, 100);

  const start = lines.findIndex((line) => isCommerceSectionStartLine(line.text || ""));
  if (start >= 0) {
    for (let i = start + 1; i < lines.length; i += 1) {
      const raw = lines[i].text || "";
      if (!raw) continue;
      if (/合計|total|小計|jct|税|請求とお支払い|お支払い情報|ご注文合計/i.test(raw)) continue;
      if (/[¥￥$]\s?\d|\d+\s?円/.test(raw)) {
        const fromMoneyLine = extractNameFromMoneyLine(raw);
        if (fromMoneyLine) pushCandidate(fromMoneyLine, 72);
        continue;
      }
      if (/更新|問題を報告|月額|app[l1]e\s*acc[o0]unt|gmail|請求先|日付|領収書|請求書|書類番号|注文番号|amex|jpn/i.test(raw)) continue;
      if (raw.length > 60) continue;
      if (!isLikelyValidAppCandidate(raw)) continue;

      let score = rankServiceCandidate(raw, lines[i]);
      if (/[ぁ-んァ-ヶ一-龠]/.test(raw)) score += 15;
      if (/[A-Za-z]{3,}/.test(raw)) score += 10;
      if (/[\\[\\]{}]|_{2,}|-{2,}/.test(raw)) score -= 25;
      if (/\\bos\\b|etmn|snete|sneteeee/i.test(raw)) score -= 25;
      if (hasLongNoisyTail(raw)) score -= 20;
      pushCandidate(raw, score);
    }
  }

  const picked = out
    .sort((a, b) => b.score - a.score || a.name.length - b.name.length)
    .filter((x) => x.score >= 52 || out.length <= 1)
    .map((x) => shortenServiceName(trimServiceSuffixes(x.name)))
    .filter((name) => !isLikelyGarbageServiceName(name))
    .filter((name) => isHighConfidenceServiceName(name));

  if (!picked.length) {
    const canonicals = findCanonicalServicesInText(text).filter((name) => !["APP_STORE", "ICLOUD"].includes(name));
    if (canonicals.length) return canonicals.map((c) => displayServiceName(c)).slice(0, 6);
    const bestUnknown = out
      .sort((a, b) => b.score - a.score)
      .map((x) => shortenServiceName(trimServiceSuffixes(x.name)))
      .find((name) => !isLikelyGarbageServiceName(name) && name.length >= 3);
    return bestUnknown ? [bestUnknown] : [];
  }
  return uniqueNonEmpty(picked).slice(0, 6);
}

function normalizeAppName(value) {
  const cleaned = normalizeServiceName(normalizeOcrNameText(value))
    .replace(/[\\[\\]{}<>]/g, " ")
    .replace(/\(.*?\)/g, " ")
    .replace(/（.*?）/g, " ")
    .replace(/\b(?:premium|month|monthly|features|backup|online)\b/gi, " ")
    .replace(/\b(?:updated?|report|problem|os|ios|jct|tax)\b/gi, " ")
    .replace(/\bAApp\b/gi, "App")
    .replace(/[_|]/g, " ")
    .replace(/[^\p{L}\p{N}\s+\-]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned) return "";

  const mapped = canonicalServiceName(cleaned);
  const normalized = mapped ? displayServiceName(mapped) : cleaned.slice(0, 40);
  return trimServiceSuffixes(normalized);
}

function extractNameFromMoneyLine(value) {
  const raw = normalizeDateText(value || "");
  if (!raw) return "";

  const left = raw
    .replace(/([¥￥$]\s?\d[\d,]*(?:\.\d+)?|\d[\d,]*(?:\.\d+)?\s?円)/gi, " ")
    .replace(/\b(?:jct|tax|含む|included|小計|subtotal|合計|total)\b/gi, " ")
    .replace(/[()（）【】[\]{}]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!left || left.length < 2) return "";
  if (!isLikelyValidAppCandidate(left)) return "";
  const normalized = normalizeAppName(left);
  if (isSummaryLikeName(normalized)) return "";
  return normalized;
}

function rankServiceCandidate(raw, row = { x0: 0, y0: 0 }) {
  let score = 30;
  const line = normalizeDateText(raw || "");
  if (/app\s?store|itunes|download|ダウンロード/i.test(line)) score += 12;
  if (/premium|月額|年額/i.test(line)) score += 8;
  if (/更新|問題を報告|app[l1]e\s*acc[o0]unt|gmail|請求先|注文番号|書類番号/i.test(line)) score -= 30;
  if ((row.x0 || 0) < 180) score += 6;
  return score;
}

function isHighConfidenceServiceName(value) {
  const v = normalizeServiceName(value || "");
  if (!v) return false;
  if (canonicalServiceName(v)) return true;
  if (/[{}\[\]_]/.test(v)) return false;
  const hasWord = /[A-Za-z]{3,}|[ぁ-んァ-ヶ一-龠々ー]{2,}/.test(v);
  if (!hasWord) return false;
  if (v.length >= 3 && v.length <= 36 && !hasLongNoisyTail(v)) return true;
  return false;
}

function summarizeServiceNames(names, itemCount = 0) {
  const list = uniqueNonEmpty((names || []).map((x) => shortenServiceName(trimServiceSuffixes(x))));
  if (!list.length) return "AppStore";
  return list[0];
}

function withItemCountSuffix(serviceName, itemCount) {
  const base = normalizeServiceName(serviceName || "").replace(/\+\d+$/g, "").trim();
  const count = normalizeItemCount(itemCount);
  if (!base) return "";
  if (/(複数|他複数|複数明細|multi[-\s]?items?)/i.test(base)) return base;
  if (count <= 1) return base;
  return `${base}+${count - 1}`;
}

function countStoreLineItems(ocrData, fullText = "", totalAmount = "") {
  const lines = getLines(ocrData);
  const words = getWords(ocrData);
  const scope = findAppleCommerceScope(lines);
  if (!scope.found && !hasCommerceSectionHint(fullText)) return 1;

  const total = parseMoneyValue(totalAmount);
  const sectionKinds = scope.found ? detectCommerceSectionKinds(lines, scope) : new Set();
  const directItems = scope.found ? compactLineItems(extractAppleLineItems(lines, words, scope, totalAmount)) : [];
  const direct = directItems.length;

  if (scope.found) {
    const scoped = lines.slice(scope.begin, scope.end);
    const reportRowsScoped = countHintRowsDedup(scoped, /問題を報告|report a problem/i);
    const reviewRowsScoped = countHintRowsDedup(scoped, /レビューを書く|write a review/i);
    const distinctNames = extractDistinctItemNames(directItems).length;
    const uniquePrices = uniqueNonEmpty(collectPriceTokensFromLines(lines, scope, total).map((v) => String(v))).length;
    const iconRows = countIconLikeItemRows(lines, scope, totalAmount);
    const amountWordRows = countAmountRowsFromWords(lines, words, scope, totalAmount);
    const anchorMax = Math.max(direct, reportRowsScoped, reviewRowsScoped, distinctNames, uniquePrices, iconRows, Math.min(amountWordRows, 3));
    const feedbackDuplicated = reportRowsScoped >= 2 || reviewRowsScoped >= 2;
    const feedbackCorroborated = distinctNames >= 2 || uniquePrices >= 2 || iconRows >= 2;
    const strongMultipleSmall = sectionKinds.size >= 2 || uniquePrices >= 2 || (feedbackDuplicated && feedbackCorroborated);

    if (iconRows === 1 && sectionKinds.size <= 1 && uniquePrices <= 1) return 1;

    // Stability guard: when all reliable anchors are in the small range, never let noisy OCR-composition
    // heuristics inflate to 5-9 items.
    if (anchorMax <= 3) {
      if (anchorMax >= 2 && strongMultipleSmall) return normalizeItemCount(anchorMax);
      return 1;
    }
  }

  const start = lines.findIndex((line) => isCommerceSectionStartLine(line.text || ""));
  const begin = start >= 0 ? start + 1 : (scope.found ? scope.begin + 1 : 0);
  let reportCount = 0;
  let reviewCount = 0;
  let updateCount = 0;

  for (let i = begin; i < lines.length; i += 1) {
    const raw = lines[i].text || "";
    if (!raw) continue;
    if (isCommerceFooterStartLine(raw)) break;
    if (/合計|total|小計|jct|税|請求とお支払い|お支払い情報|ご注文合計/i.test(raw)) continue;
    if (/問題を報告|report a problem/i.test(raw)) reportCount += 1;
    if (/レビューを書く|write a review/i.test(raw)) reviewCount += 1;
    if (/更新|updated/i.test(raw)) updateCount += 1;
  }
  // Apple receipt rows often include "問題を報告する" once per item.
  const amountCompositionCount = inferLineItemCountFromAmountComposition(lines, begin);
  const looseAmountCount = inferLineItemCountFromLooseAmounts(lines, begin, totalAmount);
  const mentionCount = Math.min(
    findCanonicalServicesInText(lines.map((x) => x.text).join(" ")).filter((x) => !["APP_STORE", "ICLOUD"].includes(x)).length,
    6
  );
  const linePriceValues = scope.found ? collectPriceTokensFromLines(lines, scope, total) : [];
  const linePriceSignal = linePriceValues.length;
  const lineResidualMissing = inferMissingCountFromResidual(total, linePriceValues);
  const lineResidualSignal = lineResidualMissing >= 1 ? linePriceSignal + lineResidualMissing : 0;

  const textNorm = normalizeDateText(fullText);
  const textReportCount = ((textNorm.match(/問題を報告|report a problem/gi) || []).length);
  const textReviewCount = ((textNorm.match(/レビューを書く|write a review/gi) || []).length);
  const textUpdateCount = ((textNorm.match(/更新\s*[：:]|updated\s*[:]/gi) || []).length);
  const textMentionCount = Math.min(
    findCanonicalServicesInText(textNorm).filter((x) => !["APP_STORE", "ICLOUD"].includes(x)).length,
    6
  );
  const textMoneyCount = inferLineItemCountFromTextAmounts(textNorm, totalAmount);
  const ocrLooseCount = inferLineItemCountFromLooseTextNumbers(textNorm, totalAmount);

  const feedbackSignal = Math.max(reportCount, reviewCount, textReportCount, textReviewCount);
  const updateSignal = Math.max(updateCount, textUpdateCount);
  const mentionSignal = Math.max(mentionCount, textMentionCount);
  const amountSignal = Math.max(amountCompositionCount, looseAmountCount, textMoneyCount, ocrLooseCount);
  const sectionSignal = sectionKinds.size >= 2 ? 2 : (!scope.found && hasMixedCommerceSections(fullText) ? 2 : 0);
  const residualSignal = Math.max(lineResidualSignal, amountSignal);

  const signals = [direct, feedbackSignal, updateSignal, mentionSignal, amountSignal, linePriceSignal, lineResidualSignal, sectionSignal]
    .filter((n) => Number.isFinite(n) && n >= 1);
  if (!signals.length) return 1;

  const candidates = uniqueNonEmpty(signals.map(String))
    .map((x) => Number(x))
    .filter((n) => Number.isFinite(n) && n >= 2)
    .sort((a, b) => b - a);
  for (const candidate of candidates) {
    const tolerance = candidate >= 4 ? 1 : 0;
    let support = 0;
    if (direct >= 1 && Math.abs(direct - candidate) <= tolerance) support += 1;
    if (feedbackSignal >= 1 && Math.abs(feedbackSignal - candidate) <= tolerance) support += 1;
    if (updateSignal >= 1 && Math.abs(updateSignal - candidate) <= tolerance) support += 1;
    if (mentionSignal >= 1 && Math.abs(mentionSignal - candidate) <= tolerance) support += 1;
    if (amountSignal >= 1 && Math.abs(amountSignal - candidate) <= tolerance) support += 1;
    if (linePriceSignal >= 1 && Math.abs(linePriceSignal - candidate) <= tolerance) support += 1;
    if (lineResidualSignal >= 2 && Math.abs(lineResidualSignal - candidate) <= tolerance) support += 1;
    if (sectionSignal >= 2 && candidate >= 2) support += 1;
    if (support >= 2) return normalizeItemCount(candidate);
  }

  if (direct >= 2) return normalizeItemCount(direct);
  const feedbackSupported = direct >= 2 || amountSignal >= 2 || mentionSignal >= 2 || linePriceSignal >= 2 || sectionSignal >= 2;
  if (feedbackSignal >= 2 && feedbackSupported) return normalizeItemCount(feedbackSignal);
  if (feedbackSignal >= 3 && direct >= 1) return normalizeItemCount(feedbackSignal);
  if (lineResidualSignal >= 2) return normalizeItemCount(lineResidualSignal);
  if (linePriceSignal >= 2) return normalizeItemCount(linePriceSignal);
  if (amountSignal >= 2) return normalizeItemCount(amountSignal);
  if (residualSignal >= 2) return normalizeItemCount(residualSignal);
  if (sectionSignal >= 2) return 2;
  return 1;
}

function estimateAppleItemCount(text, ocrData, totalAmount = "") {
  const t = (text || "").toLowerCase();
  if (!hasCommerceSectionHint(t)) return 1;
  return normalizeItemCount(countStoreLineItems(ocrData, text, totalAmount));
}

function inferLineItemCountFromTextAmounts(text, totalAmount = "") {
  const total = parseMoneyValue(totalAmount);
  const values = parseMoneyListFromTextAll(text)
    .map((x) => parseMoneyValue(x))
    .filter((v) => Number.isFinite(v) && v > 0 && v <= 500000);
  if (!values.length) return 0;

  if (total > 0) {
    const parts = values.filter((v) => v < total && v >= 80).sort((a, b) => b - a).slice(0, 24);
    const k = findCombinationCountWithDuplicates(parts, total, 2, 12, 2);
    if (k >= 2) return k;
  }

  // Fallback: count medium/high price tokens with duplicates preserved.
  const medium = values.filter((v) => v >= 300 && (total <= 0 || v < total)).length;
  return Math.min(medium, 12);
}

function inferLineItemCountFromLooseTextNumbers(text, totalAmount = "") {
  const total = parseMoneyValue(totalAmount);
  if (!total) return 0;

  const nums = extractLooseNumericCandidates(text)
    .filter((v) => v >= 80 && v < total)
    .sort((a, b) => b - a)
    .slice(0, 12);
  if (nums.length < 2) return 0;
  return findCombinationCount(nums, total, 2, 4);
}

function inferLineItemCountFromAmountComposition(lines, begin = 0) {
  const candidates = [];
  let explicitTotal = 0;

  for (let i = begin; i < lines.length; i += 1) {
    const raw = lines[i]?.text || "";
    if (!raw) continue;
    if (/(jct|tax|小計|subtotal|含む|included|注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex)/i.test(raw)) {
      continue;
    }
    const amounts = parseMoneyListFromText(raw).map((x) => parseMoneyValue(x)).filter((v) => v > 0);
    if (!amounts.length) continue;

    if (/ご注文合計|合計|total/i.test(raw)) {
      explicitTotal = Math.max(explicitTotal, ...amounts);
    }
    candidates.push(...amounts);
  }

  if (!candidates.length) return 0;
  const values = candidates.map((v) => Number(v)).filter((v) => Number.isFinite(v) && v > 0);
  if (values.length < 2) return 0;

  const total = explicitTotal || Math.max(...values);
  const parts = values.filter((v) => v < total && v >= 80).sort((a, b) => b - a).slice(0, 24);
  if (parts.length < 2) return 0;
  return findCombinationCountWithDuplicates(parts, total, 2, 12, 2);
}

function inferLineItemCountFromLooseAmounts(lines, begin = 0, totalAmount = "") {
  const total = parseMoneyValue(totalAmount);
  if (!total) return 0;

  const vals = [];
  for (let i = begin; i < lines.length; i += 1) {
    const raw = normalizeDateText(lines[i]?.text || "");
    if (!raw) continue;
    if (/請求とお支払い|お支払い情報|payment|billing/i.test(raw)) break;
    if (/(注文番号|書類番号|app[l1]e\s*acc[o0]unt|請求先|american express|visa|master|amex|jct|tax|小計|合計|total)/i.test(raw)) {
      continue;
    }
    const nums = raw.match(/\b\d{3,5}\b/g) || [];
    for (const s of nums) {
      const v = Number(s);
      if (!Number.isFinite(v)) continue;
      if (v >= 90000) continue;
      if (v >= 80 && v < total) vals.push(v);
    }
  }

  const uniq = uniqueNonEmpty(vals.map((v) => String(v))).map((x) => Number(x)).sort((a, b) => b - a).slice(0, 10);
  if (uniq.length < 2) return 0;
  return findCombinationCount(uniq, total, 2, 4);
}

function findCombinationCount(values, target, minSize = 2, maxSize = 4) {
  let best = 0;
  function dfs(start, sum, size) {
    if (sum === target) {
      if (size >= minSize && size <= maxSize) best = Math.max(best, size);
      return;
    }
    if (sum > target || size >= maxSize) return;
    for (let i = start; i < values.length; i += 1) {
      dfs(i + 1, sum + values[i], size + 1);
    }
  }
  dfs(0, 0, 0);
  return best;
}

function findCombinationCountWithDuplicates(values, target, minSize = 2, maxSize = 12, tolerance = 2) {
  const nums = (values || [])
    .map((v) => Number(v))
    .filter((v) => Number.isFinite(v) && v > 0)
    .sort((a, b) => b - a);
  if (!target || !nums.length) return 0;

  const maxCount = Math.min(Math.max(minSize, maxSize), nums.length);
  const sumsByCount = Array.from({ length: maxCount + 1 }, () => new Set());
  sumsByCount[0].add(0);

  for (const v of nums) {
    for (let c = maxCount - 1; c >= 0; c -= 1) {
      if (!sumsByCount[c].size) continue;
      for (const sum of sumsByCount[c]) {
        const next = sum + v;
        if (next > target + tolerance) continue;
        sumsByCount[c + 1].add(next);
      }
    }
  }

  for (let c = maxCount; c >= minSize; c -= 1) {
    for (const sum of sumsByCount[c]) {
      if (Math.abs(sum - target) <= tolerance) return c;
    }
  }
  return 0;
}

function hasLongNoisyTail(value) {
  const v = normalizeDateText(value || "");
  if (!v) return false;
  if (v.length < 20) return false;
  if (/[^\p{L}\p{N}\s+\-]/u.test(v)) return true;
  const words = v.split(/\s+/).filter(Boolean);
  const shortWordRatio = words.length ? words.filter((w) => w.length <= 1).length / words.length : 0;
  return shortWordRatio > 0.35;
}

function canonicalServiceName(value) {
  const v = normalizeOcrNameText(value || "");
  for (const entry of SERVICE_CATALOG.aliases) {
    if (entry.re.test(v)) return entry.canonical;
  }
  return "";
}

function normalizeOcrNameText(value) {
  const src = normalizeDateText(value || "");
  // Repair common OCR confusions in product names: L1NE -> LINE, C0in -> Coin.
  return src
    .replace(/([A-Za-z])0(?=[A-Za-z])/g, "$1O")
    .replace(/([A-Za-z])1(?=[A-Za-z])/g, "$1I")
    .replace(/([A-Za-z])5(?=[A-Za-z])/g, "$1S")
    .replace(/\bAApp\b/gi, "App");
}

function normalizeLatinTokenCase(value) {
  return String(value || "").replace(/\b[A-Za-z][A-Za-z0-9]{1,24}\b/g, (token) => {
    if (/[0-9]/.test(token)) return token;
    if (token === token.toUpperCase() && token.length <= 5) return token;
    return token.charAt(0).toUpperCase() + token.slice(1).toLowerCase();
  });
}

function findCanonicalServicesInText(value) {
  const v = value || "";
  const out = [];
  for (const entry of SERVICE_CATALOG.aliases) {
    if (entry.re.test(v)) out.push(entry.canonical);
  }
  return uniqueNonEmpty(out);
}

function displayServiceName(canonical) {
  if (!canonical) return "";
  return SERVICE_CATALOG.display[canonical] || canonical;
}

function trimServiceSuffixes(value) {
  let v = normalizeJapaneseSpacing(value || "");
  if (!v) return "";

  v = v
    .replace(/\s*[-–—]\s*(スマホ|お金|残高|口座|管理).*/i, "")
    .replace(/\s*[-–—]\s*(track|premium|features).*/i, "")
    .replace(/\s*(\(|（).*(月額|年額|subscription|premium).*(\)|）)/i, "")
    .replace(/\b(?:premium|month|monthly|plan|features)\b.*$/i, "")
    .replace(/\s+/g, " ")
    .trim();

  if (v.length > 28 && /[ぁ-んァ-ヶ一-龠]/.test(v)) {
    v = v.slice(0, 28).replace(/[-\s]+$/g, "");
  }
  return v;
}

function shortenServiceName(value) {
  let v = normalizeJapaneseSpacing(value || "");
  if (!v) return "";
  v = v
    .replace(/[｜|].*$/g, "")
    .replace(/[：:].*$/g, "")
    .replace(/\s+[-–—]\s+(スマホ|お金|残高|口座|管理|新しい|未来|できる).*/i, "")
    .replace(/\s{2,}/g, " ")
    .trim();
  if (v.length > 30) v = v.slice(0, 30).replace(/[-\s]+$/g, "");
  return v;
}

function isLikelyValidAppCandidate(value) {
  const v = normalizeDateText(value || "");
  if (!v) return false;
  if (v.length < 2 || v.length > 64) return false;
  if (/^[\d\s\-+¥￥$.,/]+$/.test(v)) return false;
  if (/^[a-z0-9\s]{1,10}$/i.test(v)) return false;
  if (/(注文番号|書類番号|app[l1]e\s*acc[o0]unt|american express|visa|master|amex|請求先|請求とお支払い)/i.test(v)) return false;

  const alphaWords = v.match(/[A-Za-z]{3,}/g) || [];
  const jpChunk = v.match(/[ぁ-んァ-ヶ一-龠々ー]{2,}/g) || [];
  return alphaWords.length > 0 || jpChunk.length > 0;
}

function normalizeServiceName(value) {
  const normalized = (value || "")
    .normalize("NFKC")
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/^app\s*store[-\s]+/i, "")
    .replace(/[¥￥$]\s?\d[\d,]*(?:\.\d+)?/g, " ")
    .replace(/領収\s*書|請求\s*書|合計|日付|書類番号|注文番号|請求先/gi, " ")
    .replace(/問題を報告する?|更新日?|月額|年額|サブスク|subscription/gi, " ")
    .replace(/\b(?:jct|tax|total|subtotal)\b/gi, " ")
    .replace(/\s+/g, " ")
    .trim();

  return normalizeJapaneseSpacing(normalized).slice(0, 48);
}

function parseMoneyFromText(value) {
  const list = parseMoneyListFromText(value);
  if (!list.length) return "";
  return list.sort((a, b) => parseMoneyValue(b) - parseMoneyValue(a))[0];
}

function parseMoneyListFromText(value) {
  return uniqueNonEmpty(parseMoneyListFromTextAll(value));
}

function parseMoneyListFromTextAll(value) {
  const raw = normalizeDateText(value || "");
  // Strict money parser for amount extraction: require currency marker or 円.
  const matches = raw.match(/([¥￥]\s?[0-9OIl][0-9OIl,]*(?:\.[0-9OIl]+)?|\$\s?[0-9OIl][0-9OIl,]*(?:\.[0-9OIl]+)?|[0-9OIl][0-9OIl,]*(?:\.[0-9OIl]+)?\s?円)/gi) || [];
  const out = [];
  for (const m of matches) {
    const normalized = normalizeAmountText(m);
    const n = parseMoneyValue(normalized);
    if (!Number.isFinite(n) || n <= 0) continue;
    if (n > 500000) continue;
    out.push(normalized);
  }
  return out;
}

function parseMoneyValue(value) {
  const n = Number((value || "").replace(/[^\d.]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function isNearSubtotalTaxSum(total, subtotalVals, taxVals) {
  for (const s of subtotalVals) {
    for (const t of taxVals) {
      const sum = s + t;
      if (Math.abs(sum - total) <= 2) return true;
    }
  }
  return false;
}

function normalizeAmountKey(value) {
  return String(parseMoneyValue(value));
}

function normalizeAmountText(value) {
  return (value || "")
    .normalize("NFKC")
    .replace(/[OoＯｏ]/g, "0")
    .replace(/[Il｜ｌＩ]/g, "1")
    .replace(/\s+/g, "")
    .replace("￥", "¥")
    .replace(/円/g, "")
    .replace(/(\d)\.(\d{3})(?!\d)/g, "$1,$2")
    .replace(/^(\d)/, "¥$1")
    .replace(/^(\$)/, "$1")
    .slice(0, 24);
}

function extractLooseNumericCandidates(text) {
  const raw = normalizeDateText(text || "");
  const matches = raw.match(/\b[0-9]{3,6}\b/g) || [];
  const vals = [];
  for (const s of matches) {
    const n = Number(s);
    if (!Number.isFinite(n)) continue;
    if (n >= 1900 && n <= 2100) continue;
    if (n >= 90000) continue;
    vals.push(n);
  }
  return uniqueNonEmpty(vals.map(String)).map(Number);
}

function extractDate(text, yearHints = [], opts = {}) {
  const allowTokenFallback = opts?.allowTokenFallback !== false;
  const n = normalizeDateParseText(text);
  const dense = squeezeDateDigits(n);
  const compact = dense.match(/(20\d{2}|19\d{2})\s*(\d{2})\s*(\d{2})/);
  if (compact) {
    const out = buildPlausibleYmd(compact[1], compact[2], compact[3]);
    if (out) return out;
  }

  const ymd = dense.match(/(20\d{2}|19\d{2})\s*[\/.\-年生午]\s*(\d{1,2})\s*[\/.\-月]\s*(\d{1,2})(?:\s*[日曰目白])?/);
  if (ymd) {
    const out = buildPlausibleYmd(ymd[1], ymd[2], ymd[3]);
    if (out) return out;
  }

  const ymdLoose = dense.match(/(20\d{2}|19\d{2})\s*[\/.\-年生午]?\s*(\d{1,2})\s*(?:[\/.\-月]|\s)\s*(\d{1,2})(?:\s*[日曰目白])?/);
  if (ymdLoose) {
    const out = buildPlausibleYmd(ymdLoose[1], ymdLoose[2], ymdLoose[3]);
    if (out) return out;
  }

  // Tolerate partial OCR year like "0年6月25日" by inferring year from hints.
  const partialYmd = dense.match(/(\d{1,4})\s*[年生午]\s*(\d{1,2})\s*月\s*(\d{1,2})(?:\s*[日曰目白])?/);
  if (partialYmd) {
    const inferredYear = inferYearFromToken(partialYmd[1], yearHints);
    if (inferredYear) {
      const out = buildPlausibleYmd(inferredYear, partialYmd[2], partialYmd[3]);
      if (out) return out;
    }
  }

  // Month-day only on date-labeled rows.
  const mdOnly = dense.match(/(?:日[付什]|[曰目白][付什]|注文[日曰目白]|ご注文[日曰目白]|date)?\s*(\d{1,2})\s*月?\s*(\d{1,2})(?:\s*[日曰目白])?/i);
  if (mdOnly) {
    const inferredYear = inferYearFromToken("", yearHints);
    if (inferredYear) {
      const out = buildPlausibleYmd(inferredYear, mdOnly[1], mdOnly[2]);
      if (out) return out;
    }
  }

  // OCR sometimes drops separators and leaves space-delimited Y M D.
  const spacedYmd = dense.match(/(20\d{2}|19\d{2})\s+(\d{1,2})\s+(\d{1,2})/);
  if (spacedYmd) {
    const out = buildPlausibleYmd(spacedYmd[1], spacedYmd[2], spacedYmd[3]);
    if (out) return out;
  }

  const dmy = dense.match(/(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(20\d{2}|19\d{2})/);
  if (dmy) {
    const out = buildPlausibleYmd(dmy[3], dmy[1], dmy[2]);
    if (out) return out;
  }

  if (allowTokenFallback) {
    const fromTokens = extractYmdFromNumericTokens(dense, yearHints);
    if (fromTokens) return fromTokens;
  }

  return "";
}

function normalizeDateParseText(value) {
  return normalizeDateText(value || "")
    .replace(/([0-9])\s*[生午]\s*(?=[0-9])/g, "$1年")
    .replace(/([0-9])\s*[SsＳｓ]\s*(?=\s*(?:[\/.\-年月日曰目白]|$))/g, "$15")
    .replace(/([0-9])\s*[曰目白]\s*(?=[^0-9]|$)/g, "$1日");
}

function squeezeDateDigits(value) {
  return (value || "").replace(/(\d)\s+(?=\d)/g, "$1");
}

function collectYearHints(text, lines = []) {
  const out = [];
  const raw = [text, ...lines.map((x) => x.text || "")].join(" ");
  const m = normalizeDateText(raw).match(/(?:19|20)\d{2}/g) || [];
  for (const y of m) {
    const n = Number(y);
    if (Number.isFinite(n) && n >= 2010 && n <= 2038) out.push(n);
  }
  return uniqueNonEmpty(out.map(String)).map(Number);
}

function inferYearFromToken(token, yearHints = []) {
  const currentYear = new Date().getFullYear();
  const t = String(token || "").replace(/\D/g, "");
  if (t.length === 4) {
    const y = Number(t);
    if (y >= 2010 && y <= 2038) return y;
    return 0;
  }

  const candidates = [];
  for (let y = 2010; y <= 2038; y += 1) candidates.push(y);
  for (const y of yearHints || []) candidates.push(y);
  const uniq = uniqueNonEmpty(candidates.map(String)).map(Number);

  let matched = uniq;
  if (t.length >= 1 && t.length <= 3) {
    matched = uniq.filter((y) => String(y).endsWith(t));
  } else if (t.length === 0 && (yearHints || []).length) {
    matched = yearHints;
  }

  if (!matched.length) {
    if (t.length === 2) {
      const y = 2000 + Number(t);
      if (y >= 2010 && y <= 2038) return y;
    }
    return 0;
  }

  // Prefer years not in far future.
  const bounded = matched.filter((y) => y <= currentYear + 1);
  const pool = bounded.length ? bounded : matched;
  pool.sort((a, b) => Math.abs(a - currentYear) - Math.abs(b - currentYear));
  return pool[0] || 0;
}

function extractYmdFromNumericTokens(value, yearHints = []) {
  const raw = normalizeDateText(value || "");
  const tokens = raw.match(/\d{1,4}/g) || [];
  if (tokens.length < 3) return "";

  for (let i = 0; i <= tokens.length - 3; i += 1) {
    const yToken = String(tokens[i] || "");
    const mToken = String(tokens[i + 1] || "");
    const dToken = String(tokens[i + 2] || "");
    if (!mToken || !dToken) continue;
    if (mToken.length > 2 || dToken.length > 2) continue;

    let year = 0;
    if (yToken.length === 4) {
      const y = Number(yToken);
      if (Number.isFinite(y) && y >= 2010 && y <= 2038) year = y;
    } else if (yToken.length >= 1 && yToken.length <= 3) {
      year = inferYearFromToken(yToken, yearHints);
    }
    if (!year) continue;

    const month = Number(mToken);
    const day = Number(dToken);
    if (!Number.isFinite(month) || !Number.isFinite(day)) continue;
    if (month < 1 || month > 12) continue;
    if (day < 1 || day > 31) continue;
    return `${year}${pad2(month)}${pad2(day)}`;
  }
  return "";
}

function hasAnchoredAppleDateEvidence(ocrData, dateValue = "") {
  const date = normalizeDateValue(dateValue);
  if (!isPlausibleDate(date)) return false;

  const lines = getLines(ocrData);
  if (!lines.length) return false;
  const pageBottom = lines.reduce((max, line) => Math.max(max, line.y0 || 0), 1);
  const labelBottomLimit = pageBottom * 0.62;
  const yearHints = collectYearHints(ocrData?.text || "", lines);

  const anchoredLines = lines.filter((line, idx) => {
    if ((line?.y0 || 0) > labelBottomLimit) return false;
    const text = normalizeDateText(line?.text || "");
    if (!text) return false;
    if (isDateLabelLike(text)) return true;
    if (idx > 0 && isDateLabelLike(lines[idx - 1]?.text || "")) return true;
    return false;
  });

  for (const line of anchoredLines) {
    const parsed = extractDate(line?.text || "", yearHints, { allowTokenFallback: false });
    if (normalizeDateValue(parsed) === date) return true;
  }
  return false;
}

function normalizeDateValue(value) {
  const only = (value || "").replace(/\D/g, "");
  if (only.length >= 8) return only.slice(0, 8);
  return "";
}

function normalizeItemCount(value) {
  const n = Number(String(value || "").replace(/[^\d]/g, ""));
  if (!Number.isFinite(n) || n < 1) return 1;
  return Math.min(Math.floor(n), 9);
}

function isPlausibleDate(value) {
  const d = normalizeDateValue(value);
  if (d.length !== 8) return false;
  const y = Number(d.slice(0, 4));
  const m = Number(d.slice(4, 6));
  const day = Number(d.slice(6, 8));
  if (y < 2010 || y > 2038) return false;
  if (m < 1 || m > 12) return false;
  if (day < 1 || day > 31) return false;
  return true;
}

function extractStorageCapacity(text) {
  const n = normalizeDateText(text).toLowerCase();
  const m = n.match(/(\d+)\s*(gb|tb)/i);
  if (!m) return "";
  return `${m[1]}${m[2].toUpperCase()}`;
}

function extractOrderProductName(ocrData, text) {
  const lines = getLines(ocrData);
  const start = lines.findIndex((line) => /ダウンロード製品|download/i.test(line.text || ""));
  const stopWords = /お支払い情報|請求とお支払い|ご注文合計|合計|小計|american express|visa|master/i;
  const picks = [];

  if (start >= 0) {
    for (let i = start + 1; i < lines.length; i += 1) {
      const raw = lines[i].text || "";
      if (!raw) continue;
      if (stopWords.test(raw)) break;
      if (/[¥￥$]\s?\d|\d+\s?円/.test(raw)) {
        const fromMoneyLine = extractNameFromMoneyLine(raw);
        if (fromMoneyLine) picks.push(fromMoneyLine);
        continue;
      }
      if (/ご注文番号|注文日|app[l1]e\s*acc[o0]unt|gmail|請求先住所|請求連絡先/i.test(raw)) continue;
      if (!isLikelyValidAppCandidate(raw)) continue;
      const clean = normalizeAppName(raw);
      if (clean && clean.length >= 3) picks.push(clean);
    }
  }

  if (!picks.length) {
    const mapped = canonicalServiceName(text);
    if (mapped) picks.push(displayServiceName(mapped));
  }
  return uniqueNonEmpty(picks).slice(0, 2).join("+");
}

function makeLearningKey(text, fallbackName) {
  const norm = (text || "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .slice(0, 800);
  const seed = norm || (fallbackName || "fallback");
  return `v3:${hashString(seed)}`;
}

function hashString(value) {
  let h = 5381;
  for (let i = 0; i < value.length; i += 1) {
    h = (h * 33) ^ value.charCodeAt(i);
  }
  return (h >>> 0).toString(16);
}

function loadLearnedOverrides(key) {
  if (!key) return null;
  try {
    const raw = localStorage.getItem(LEARN_STORE_KEY);
    if (!raw) return null;
    const map = JSON.parse(raw);
    const row = map[key];
    if (!row) return null;
    const loaded = {
      date: normalizeDateValue(row.date || ""),
      amount: normalizeAmountText(row.amount || ""),
      itemCount: normalizeItemCount(row.itemCount || 1),
      service: normalizeServiceName(row.service || ""),
    };
    if (isLikelyGarbageServiceName(loaded.service)) loaded.service = "";
    return loaded;
  } catch {
    return null;
  }
}

function shouldApplyLearnedOverrides(extracted, learned) {
  const extService = normalizeServiceName(extracted?.service || "");
  const learnedService = normalizeServiceName(learned?.service || "");
  const extCount = normalizeItemCount(extracted?.itemCount || 1);
  const learnedCount = normalizeItemCount(learned?.itemCount || 1);

  // Never overwrite safe fallback labels for mixed/uncertain receipts.
  if (/複数|他複数|複数明細|multi[-\s]?items?/i.test(extService)) return false;
  // Avoid applying stale low-count overrides over clearly multi-item extraction.
  if (extCount >= 4 && learnedCount <= 3) return false;
  // Never re-apply a learned "multiple" label as a hard override.
  if (isStrictMultipleServiceLabel(learnedService)) return false;
  // Avoid applying stale specific service over uncertain fallback context.
  if (/app[l1]e\s*service|app\s*store/i.test(extService) && learnedService && learnedService !== extService) return false;
  return true;
}

function saveLearnedOverrides(key, fields) {
  if (!ENABLE_LEARN_OVERRIDES) return;
  if (!key) return;
  try {
    const service = normalizeServiceName(fields.service || "");
    if (isLikelyGarbageServiceName(service)) return;
    if (isStrictMultipleServiceLabel(service)) return;
    const raw = localStorage.getItem(LEARN_STORE_KEY);
    const map = raw ? JSON.parse(raw) : {};
    map[key] = {
      date: normalizeDateValue(fields.date || ""),
      amount: normalizeAmountText(fields.amount || ""),
      itemCount: normalizeItemCount(fields.itemCount || 1),
      service,
      updatedAt: Date.now(),
    };

    // Keep only recent entries.
    const entries = Object.entries(map)
      .sort((a, b) => (b[1].updatedAt || 0) - (a[1].updatedAt || 0))
      .slice(0, 200);
    const trimmed = Object.fromEntries(entries);
    localStorage.setItem(LEARN_STORE_KEY, JSON.stringify(trimmed));
  } catch {
    // ignore storage errors
  }
}

function normalizeJapaneseSpacing(value) {
  let out = value || "";
  out = out.replace(/([ぁ-んァ-ヶ一-龠々ー])\s+(?=[ぁ-んァ-ヶ一-龠々ー])/g, "$1");
  out = out.replace(/みんな\s+の\s+銀行/gi, "みんなの銀行");
  out = out.replace(/アベマ/g, "ABEMA");
  return out;
}

function isScreenshotLike(value) {
  return /スクリーンショット|screenshot|image|img[_-]?\d/i.test(value || "");
}

function isLikelyToolUiCapture(value) {
  const text = normalizeDateText(value || "");
  if (!text) return false;
  const dense = text.toLowerCase().replace(/\s+/g, "");

  const uiMarkers = [
    /zip\s*に\s*含める|include\s*in\s*zip/i,
    /明細件数|line\s*items?/i,
    /サービス名\s*\/\s*アプリ名|service\s*\/\s*app/i,
    /最終ファイル名|final\s*filename/i,
    /この名前で単体保存|download\s*this\s*file/i,
    /虫眼鏡|loupe/i,
    /判定\s*:\s*app[l1]e領収書固定|mode\s*:\s*app[l1]e\s*receipt\s*preset/i,
    /元ファイル|original\s*:/i,
    /\[build\s*:\s*20\d{6,}\w*\]/i,
  ];
  const uiHits = uiMarkers.filter((re) => re.test(text)).length;

  const receiptMarkers = [
    /領収書|請求書|receipt|invoice/i,
    /app[l1]e\s*id|app[l1]e\s*acc[o0]unt/i,
    /注文番号|ご注文番号|order\s*(number|no)/i,
    /書類番号|document\s*no/i,
    /app\s*st[o0]re|app[l1]e\s*b[o0]{2}ks|i[t1]unes/i,
    /ご注文合計|合計|total/i,
  ];
  const receiptHits = receiptMarkers.filter((re) => re.test(text)).length;

  let denseHits = 0;
  if (dense.includes("zipに含める") || dense.includes("includeinzip")) denseHits += 1;
  if (dense.includes("明細件数") || dense.includes("lineitems")) denseHits += 1;
  if (dense.includes("サービス名/アプリ名") || dense.includes("service/app")) denseHits += 1;
  if (dense.includes("最終ファイル名") || dense.includes("finalfilename")) denseHits += 1;
  if (dense.includes("この名前で単体保存") || dense.includes("downloadthisfile")) denseHits += 1;
  if (dense.includes("判定:app1e領収書固定") || dense.includes("mode:app1ereceiptpreset")) denseHits += 1;
  if (dense.includes("[build:20")) denseHits += 1;

  if (denseHits >= 3) return true;
  if (denseHits >= 2 && uiHits >= 2) return true;
  if (uiHits >= 4) return true;
  if (uiHits >= 3 && receiptHits <= 2) return true;
  if (uiHits >= 2 && /ocr[:：]|判定[:：]|mode[:：]/i.test(text)) return true;
  return false;
}

function isLikelyGarbageServiceName(value) {
  const v = (value || "").trim();
  if (!v) return true;
  if (isScreenshotLike(v)) return true;
  if (isGenericServiceLabel(v)) return true;
  if (v.length < 2) return true;
  if (/[\\[\\]{}]/.test(v)) return true;
  if (/(etmn|snete|sneteeee|og\\s+af|os\\s*月|problem|report)/i.test(v)) return true;
  const tokens = v.split(/\s+/).filter(Boolean);
  const singleTokenRate = tokens.length ? tokens.filter((x) => x.length <= 1).length / tokens.length : 0;
  if (tokens.length >= 5 && singleTokenRate >= 0.5) return true;
  const weirdChars = (v.match(/[^A-Za-z0-9ぁ-んァ-ヶ一-龠々ー+\\-\\s]/g) || []).length;
  if (weirdChars >= 4) return true;
  return false;
}

function sanitizeServiceFinal(value, kind = "general") {
  let v = normalizeServiceName(value || "");
  v = v.replace(/^AppStore[-\s]*/i, "").trim();
  const canonical = canonicalServiceName(v);
  if (canonical) v = displayServiceName(canonical);
  v = trimServiceSuffixes(v);
  if (!v || isLikelyGarbageServiceName(v)) {
    if (kind === "appstore") return "AppStore";
    if (kind === "icloud") return "iCloud";
    return "AppleService";
  }
  return v;
}

function normalizeDateText(value) {
  return (value || "")
    .normalize("NFKC")
    .replace(/[OoＯｏ]/g, "0")
    .replace(/[Il｜ｌＩ]/g, "1")
    .replace(/\s+/g, " ")
    .trim();
}

function monthFolderFromDate(dateValue) {
  const d = normalizeDateValue(dateValue);
  if (d.length !== 8) return "unknown-month";
  return `${d.slice(0, 4)}-${d.slice(4, 6)}`;
}

function getLines(ocrData = { text: "", lines: [] }) {
  if (Array.isArray(ocrData.lines) && ocrData.lines.length) return ocrData.lines;
  return (ocrData.text || "")
    .split("\n")
    .map((line, idx) => ({ text: line.trim(), y0: idx * 10, x0: 0 }))
    .filter((line) => line.text);
}

function getWords(ocrData = { words: [], lines: [], text: "" }) {
  if (Array.isArray(ocrData.words) && ocrData.words.length) return ocrData.words;
  const out = [];
  for (const line of getLines(ocrData)) {
    const parts = String(line.text || "").split(/\s+/).filter(Boolean);
    let cursor = line.x0 || 0;
    for (const p of parts) {
      out.push({ text: p, y0: line.y0 || 0, x0: cursor });
      cursor += p.length * 10;
    }
  }
  return out;
}

function uniqueNonEmpty(values) {
  const seen = new Set();
  const out = [];
  for (const v of values) {
    const x = (v || "").trim();
    if (!x || seen.has(x)) continue;
    seen.add(x);
    out.push(x);
  }
  return out;
}

function getUniqueFilename(filename, counter) {
  const current = counter.get(filename) || 0;
  counter.set(filename, current + 1);
  if (current === 0) return filename;
  const dot = filename.lastIndexOf(".");
  const hasExt = dot > 0;
  const base = hasExt ? filename.slice(0, dot) : filename;
  const ext = hasExt ? filename.slice(dot) : "";
  return `${base}-${current + 1}${ext}`;
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.style.color = isError ? "var(--danger)" : "var(--primary-2)";
}

function withBuildTag(message) {
  return `${message} [build:${BUILD_ID}]`;
}

function dateStamp() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${y}${m}${day}-${hh}${mm}`;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function extFromName(name) {
  const idx = name.lastIndexOf(".");
  return idx > -1 ? name.slice(idx + 1).toLowerCase() : "";
}

function extFromMime(type) {
  const map = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/webp": "webp",
    "image/gif": "gif",
  };
  return map[type] || "";
}

function removeExt(name) {
  const idx = name.lastIndexOf(".");
  return idx > -1 ? name.slice(0, idx) : name;
}

function normalizeCompanyName(value) {
  return (value || "OTHER")
    .normalize("NFKC")
    .replace(/[^\p{L}\p{N}\s_-]/gu, " ")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toUpperCase()
    .slice(0, 24) || "OTHER";
}

function toSafeName(input) {
  return (input || "")
    .normalize("NFKC")
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 90);
}

function truncate(value, length) {
  if ((value || "").length <= length) return value;
  return `${value.slice(0, length)}...`;
}

function preventDefaults(e) {
  e.preventDefault();
}
