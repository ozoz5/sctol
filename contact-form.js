const CONTACT_FORM_ENDPOINT_FALLBACK = "";
const PLACEHOLDER_TOKEN = "REPLACE_WITH_YOUR_FORM_ID";

const CONTACT_FORM_MESSAGES = {
  ja: {
    invalid: "必須項目を入力してください。",
    sending: "送信中です...",
    success: "送信しました。内容を確認のうえ、必要に応じて返信します。",
    configMissing: "フォーム送信先が未設定です。contact.html / contact-en.html の data-endpoint を設定してください。",
    failed: "送信に失敗しました。時間をおいて再試行してください。",
  },
  en: {
    invalid: "Please fill out required fields.",
    sending: "Sending...",
    success: "Your message was sent. We will reply if needed.",
    configMissing: "Form endpoint is not configured. Set data-endpoint in contact.html / contact-en.html.",
    failed: "Failed to send. Please try again later.",
  },
};

function getContactMessage(lang, key) {
  const dict = CONTACT_FORM_MESSAGES[lang] || CONTACT_FORM_MESSAGES.en;
  return dict[key] || CONTACT_FORM_MESSAGES.en[key] || "";
}

function resolveEndpoint(form) {
  const endpoint = (form.dataset.endpoint || CONTACT_FORM_ENDPOINT_FALLBACK || "").trim();
  if (!endpoint) return "";
  if (endpoint.includes(PLACEHOLDER_TOKEN)) return "";
  return endpoint;
}

function setFormStatus(form, type, message) {
  const statusEl = form.querySelector(".contact-status");
  if (!statusEl) return;
  statusEl.textContent = message;
  statusEl.classList.remove("ok", "error", "pending");
  if (type) statusEl.classList.add(type);
}

function setFormBusy(form, busy) {
  const submitBtn = form.querySelector(".contact-submit");
  if (!submitBtn) return;
  submitBtn.disabled = busy;
}

async function submitContactForm(form) {
  const lang = (form.dataset.lang || "en").toLowerCase();
  const endpoint = resolveEndpoint(form);
  if (!endpoint) {
    setFormStatus(form, "error", getContactMessage(lang, "configMissing"));
    return;
  }

  if (!form.checkValidity()) {
    form.reportValidity();
    setFormStatus(form, "error", getContactMessage(lang, "invalid"));
    return;
  }

  const formData = new FormData(form);
  const honeypot = String(formData.get("_gotcha") || "").trim();
  if (honeypot) return;

  const payload = {
    name: String(formData.get("name") || "").trim(),
    reply_to: String(formData.get("reply_to") || "").trim(),
    category: String(formData.get("category") || "").trim(),
    subject: String(formData.get("subject") || "").trim(),
    message: String(formData.get("message") || "").trim(),
    consent: Boolean(formData.get("consent")),
    page: location.pathname,
    userAgent: navigator.userAgent,
    lang,
  };

  setFormBusy(form, true);
  setFormStatus(form, "pending", getContactMessage(lang, "sending"));
  try {
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }
    form.reset();
    setFormStatus(form, "ok", getContactMessage(lang, "success"));
  } catch (_error) {
    setFormStatus(form, "error", getContactMessage(lang, "failed"));
  } finally {
    setFormBusy(form, false);
  }
}

function setupContactForms() {
  const forms = document.querySelectorAll(".js-contact-form");
  if (!forms.length) return;
  forms.forEach((form) => {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      await submitContactForm(form);
    });
  });
}

setupContactForms();
