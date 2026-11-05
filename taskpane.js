/* global Office, Word */

// ==========================================
// DEVELOPER & SUPPORT CONFIGURATION
// (আপনার হোয়াটসঅ্যাপ নম্বর ও ইমেইল এখানে পরিবর্তন করতে পারেন)
// ==========================================
const SUPPORT_CONFIG = {
    whatsappNumber: "8801641034759", // আপনার দেশের কোডসহ WhatsApp নম্বর দিন
    supportEmail: "masudparvezpmk@gmail.com", // আপনার সাপোর্ট ইমেইল দিন
    appVersion: "v2.4.0"
};

let currentImageBase64 = null;
let currentImageMime = null;
let reportImageName = null;

Office.onReady(() => {
    initApp();
});

function initApp() {
    // 1. Account & Subscription Initialization
    initAccountSystem();

    // 2. Load saved API keys & AI Model (3.8 Flash, 3.1 Pro, 4 Argon)
    const savedKeys = localStorage.getItem("parvez_gemini_keys") || "";
    const savedModel = localStorage.getItem("parvez_gemini_model") || "3.8-flash";
    const savedThinking = localStorage.getItem("parvez_thinking_level") || "medium";

    document.getElementById("apiKeysInput").value = savedKeys;
    const modelSelect = document.getElementById("modelSelect");
    if ([...modelSelect.options].some(o => o.value === savedModel)) {
        modelSelect.value = savedModel;
    }
    document.getElementById("thinkingSelect").value = savedThinking;
    updateModelSummaryLabel();

    if (!savedKeys) {
        document.getElementById("apiSettings").open = true;
    }

    modelSelect.onchange = updateModelSummaryLabel;

    document.getElementById("saveKeysBtn").onclick = () => {
        localStorage.setItem("parvez_gemini_keys", document.getElementById("apiKeysInput").value.trim());
        localStorage.setItem("parvez_gemini_model", modelSelect.value);
        localStorage.setItem("parvez_thinking_level", document.getElementById("thinkingSelect").value);
        updateModelSummaryLabel();
        setStatus("✅ API & Model Settings saved!", "#d1fae5", "#065f46");
        document.getElementById("apiSettings").open = false;
    };

    // 3. Image Upload & Paste Handling
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");

    dropZone.onclick = () => fileInput.click();
    fileInput.onchange = (e) => {
        if (e.target.files && e.target.files[0]) {
            loadImageFile(e.target.files[0]);
        }
    };

    window.addEventListener("paste", (e) => {
        // Don't hijack paste if typing in textareas/inputs
        const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : "";
        if (activeTag === "textarea" || activeTag === "input") return;

        if (e.clipboardData && e.clipboardData.items) {
            for (let item of e.clipboardData.items) {
                if (item.type.indexOf("image") !== -1) {
                    const file = item.getAsFile();
                    loadImageFile(file);
                    e.preventDefault();
                    break;
                }
            }
        }
    });

    document.getElementById("clearImgBtn").onclick = clearImage;
    document.getElementById("extractBtn").onclick = extractFromImageWithAI;

    // 4. Refresh / Clear Text Button & Copy Prompt Button
    document.getElementById("clearTextBtn").onclick = () => {
        document.getElementById("rawTextInput").value = "";
        setStatus("🧹 টেক্সট বক্স খালি করা হয়েছে।", "#e2e8f0", "#334155");
    };
    document.getElementById("copyPromptBtn").onclick = copyPromptToClipboard;
    document.getElementById("insertFormattedBtn").onclick = formatAndInsertToWord;

    // 5. Error Report & Support Buttons
    initSupportAndReportSystem();
}

function updateModelSummaryLabel() {
    const val = document.getElementById("modelSelect").value;
    const names = {
        "3.8-flash": "3.8 Flash",
        "3.1-pro": "3.1 Pro",
        "4-argon": "4 Argon"
    };
    document.getElementById("activeModelLabel").innerText = names[val] || "3.8 Flash";
}

// ==========================================
// ACCOUNT & SUBSCRIPTION SYSTEM
// ==========================================
function initAccountSystem() {
    const savedEmail = localStorage.getItem("parvez_user_email") || "masudparvez@gmail.com";
    const savedPlan = localStorage.getItem("parvez_user_plan") || "Diamond (1 Year)";
    const savedExpiry = localStorage.getItem("parvez_user_expiry") || "Expires 10/29/2027 05:31 PM";

    const loggedInView = document.getElementById("loggedInView");
    const loggedOutView = document.getElementById("loggedOutView");

    function renderAccount() {
        const currentEmail = localStorage.getItem("parvez_user_email");
        if (currentEmail) {
            document.getElementById("userEmailDisplay").innerText = currentEmail;
            document.getElementById("subPlanName").innerText = localStorage.getItem("parvez_user_plan") || savedPlan;
            document.getElementById("subExpiryDate").innerText = localStorage.getItem("parvez_user_expiry") || savedExpiry;
            loggedInView.classList.remove("hidden");
            loggedOutView.classList.add("hidden");
        } else {
            loggedInView.classList.add("hidden");
            loggedOutView.classList.remove("hidden");
        }
    }

    if (!localStorage.getItem("parvez_user_init")) {
        localStorage.setItem("parvez_user_email", savedEmail);
        localStorage.setItem("parvez_user_plan", savedPlan);
        localStorage.setItem("parvez_user_expiry", savedExpiry);
        localStorage.setItem("parvez_user_init", "1");
    }

    document.getElementById("signOutBtn").onclick = () => {
        localStorage.removeItem("parvez_user_email");
        renderAccount();
        setStatus("Signed out.", "#e2e8f0", "#334155");
    };

    document.getElementById("signInBtn").onclick = () => {
        const email = document.getElementById("loginEmailInput").value.trim();
        if (!email || !email.includes("@")) {
            setStatus("⚠️ অনুগ্রহ করে সঠিক ইমেইল দিন!", "#fee2e2", "#991b1b");
            return;
        }
        localStorage.setItem("parvez_user_email", email);
        localStorage.setItem("parvez_user_plan", "Diamond (1 Year)");
        localStorage.setItem("parvez_user_expiry", "Expires 10/29/2027 05:31 PM");
        renderAccount();
        setStatus("✅ Signed in as " + email, "#d1fae5", "#065f46");
    };

    renderAccount();
}

// ==========================================
// ERROR REPORT, WHATSAPP & EMAIL SUPPORT
// ==========================================
function initSupportAndReportSystem() {
    const reportFileInput = document.getElementById("reportFileInput");
    const attachStatus = document.getElementById("reportAttachmentStatus");

    document.getElementById("reportImgBtn").onclick = () => reportFileInput.click();

    reportFileInput.onchange = (e) => {
        if (e.target.files && e.target.files[0]) {
            const file = e.target.files[0];
            reportImageName = file.name;
            attachStatus.innerText = "📎 যুক্ত করা হয়েছে: " + file.name;
            attachStatus.classList.remove("hidden");

            // Copy image to clipboard if supported so user can Ctrl+V in WhatsApp/Email
            try {
                const reader = new FileReader();
                reader.onload = () => {
                    setStatus("🖼️ স্ক্রিনশট যুক্ত হয়েছে। Send (➤) চাপুন।", "#dbeafe", "#1e40af");
                };
                reader.readAsDataURL(file);
            } catch (err) {}
        }
    };

    document.getElementById("reportSendBtn").onclick = () => {
        const issueText = document.getElementById("reportTextInput").value.trim();
        const currentEq = document.getElementById("rawTextInput").value.trim();
        const userEmail = localStorage.getItem("parvez_user_email") || "Unknown User";

        if (!issueText && !reportImageName) {
            setStatus("⚠️ সমস্যার বিবরণ/ফিল্ড কোড লিখুন অথবা ছবি যুক্ত করুন!", "#fee2e2", "#991b1b");
            return;
        }

        const reportBody =
            `*Parvez Math Pro (${SUPPORT_CONFIG.appVersion}) - Error Report*\n` +
            `User: ${userEmail}\n` +
            `Issue / Unconverted Code: ${issueText || "See attached image"}\n` +
            (reportImageName ? `Attached Image: ${reportImageName} (Please paste/attach image)\n` : "") +
            (currentEq ? `\n--- Current EQ Box Content ---\n${currentEq.substring(0, 500)}` : "");

        const waUrl = `https://wa.me/${SUPPORT_CONFIG.whatsappNumber}?text=${encodeURIComponent(reportBody)}`;
        window.open(waUrl, "_blank");

        document.getElementById("reportTextInput").value = "";
        reportImageName = null;
        attachStatus.classList.add("hidden");
        setStatus("📨 রিপোর্ট পাঠানোর জন্য WhatsApp ওপেন করা হয়েছে!", "#d1fae5", "#065f46");
    };

    document.getElementById("whatsappSupportBtn").onclick = () => {
        const userEmail = localStorage.getItem("parvez_user_email") || "";
        const msg = `হ্যালো Masud Parvej ভাই, আমি Parvez Math Pro (${SUPPORT_CONFIG.appVersion}) ব্যবহার করছি (${userEmail})। আমার কিছু সাহায্য প্রয়োজন।`;
        window.open(`https://wa.me/${SUPPORT_CONFIG.whatsappNumber}?text=${encodeURIComponent(msg)}`, "_blank");
    };

    document.getElementById("emailSupportBtn").onclick = () => {
        const userEmail = localStorage.getItem("parvez_user_email") || "";
        const subject = `Parvez Math Pro (${SUPPORT_CONFIG.appVersion}) Support - ${userEmail}`;
        const body = `Hello Masud Parvej,\n\nI need support with Parvez Math Pro:\n\n`;
        window.open(`mailto:${SUPPORT_CONFIG.supportEmail}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`, "_blank");
    };
}

function setStatus(msg, bg = "#e2e8f0", color = "#334155") {
    const box = document.getElementById("statusBox");
    box.innerText = msg;
    box.style.backgroundColor = bg;
    box.style.color = color;
}

function loadImageFile(file) {
    const reader = new FileReader();
    reader.onload = (evt) => {
        const dataUrl = evt.target.result;
        currentImageMime = file.type || "image/png";
        currentImageBase64 = dataUrl.split(",")[1];

        const img = document.getElementById("imagePreview");
        img.src = dataUrl;
        img.classList.remove("hidden");
        document.getElementById("dropText").classList.add("hidden");
        document.getElementById("clearImgBtn").classList.remove("hidden");
        document.getElementById("extractBtn").disabled = false;
        setStatus("🖼️ Image loaded. Click 'Extract EQ from Image'.", "#dbeafe", "#1e40af");
    };
    reader.readAsDataURL(file);
}

function clearImage() {
    currentImageBase64 = null;
    currentImageMime = null;
    document.getElementById("fileInput").value = "";
    document.getElementById("imagePreview").classList.add("hidden");
    document.getElementById("dropText").classList.remove("hidden");
    document.getElementById("clearImgBtn").classList.add("hidden");
    document.getElementById("extractBtn").disabled = true;
    setStatus("Ready.");
}

// ==========================================
// PROMPT BUILDER (Dynamic Mode: Questions vs Full Page)
// ==========================================
function getSystemPrompt() {
    const mode = document.getElementById("extractMode").value;
    const modeRule = mode === "questions"
        ? "OUTPUT REQUIREMENT: Extract ONLY the questions. Do NOT include any solutions, proofs, or answers."
        : "OUTPUT REQUIREMENT: Extract EVERYTHING from the image including all questions, full step-by-step solutions, proofs, side-notes, and answers.";

    return `Extract text from the given image into a single plain text response. Do NOT wrap in markdown code blocks. NO citations.
${modeRule}

RULES FOR MATH FORMATTING (MS Word EQ Fields):
1. EQ SCOPE & NESTING: Use \`{ EQ }\` ONLY for fractions (\`\\f\`), roots (\`\\r\`), matrices, vectors, limits, factorials, complex powers, or tall brackets/integrals. STRICTLY NO nested \`{ EQ }\` tags. NEVER type a literal \`}\` inside an EQ field before the very end.
2. OUTSIDE EQ: Plain text. Simple powers/subs = \`\\s\\up0(...)\` & \`\\s\\do0(...)\`. 
3. INSIDE EQ (Powers & Subs): Superscripts = \`\\s\\up4(...)\`. Subscripts = \`\\s\\do3(...)\` (Strictly for 7.5 pt font size rendering). NEVER use \`\\s\\up0\` or \`\\s\\do0\` inside \`{ EQ }\` tags. Nested superscripts = \`\\s\\up7(...)\`. Fractional powers on plain bases = \`x{ EQ \\s\\up10(\\f(1,3))}\`.
4. AUTO-SCALING BRACKETS: NEVER use plain keyboard brackets \`[ ]\`, \`{ }\` for tall content. 
   - Parentheses ( ): \`\\b(...)\` (NEVER use \`\\b\\bc\\(\`)
   - Braces { }: \`\\b\\bc\\{(...)\`
   - Square [ ]: \`\\b\\bc\\[(...)\`
   - Absolute | |: \`\\b\\bc\\|(...)\`
   - Single side: \`\\b\\lc\\{(...)\`, \`\\b\\lc\\[(...)\`, etc.
   - PARENTHESIS MATCHING: Every switch like \`\\b...(\` needs exactly ONE matching \`)\`. NEVER add literal \`}\`, \`]\`, or extra \`)\`.
5. COMMA & TEXT HANDLING (CRITICAL):
   - Commas: A literal comma \`,\` inside an EQ field acts as a syntax separator. To use a comma in math (like coordinates), wrap it in quotes: \`{ EQ \\b("a, -" \\f(π,2)) }\`.
   - Text: NEVER use LaTeX text commands like \`\\text{}\`. Write plain text (including Bengali words) directly: \`{ EQ \\b(৪২০০ এর \\f(৫,২১)) }\`.
6. INTEGRALS: 
   - Definite: Use \`{ EQ \\s(upper,∫,lower) }\` (e.g., \`{ EQ \\s(4,∫,2) }\`). 
   - Indefinite (tall content): Wrap the entire expression in \`{ EQ \\b\\lc\\∫(...) }\`.
7. SYMBOLS: Real Unicode ONLY (θ, α, ·, ×, ∴, ⇒, ⟶, ∫). STRICTLY NO LaTeX commands. Prime = \`'\`.
8. ADVANCED & MATRIX ORDER: Matrices = \`\\a\\hs10\\coN(...)\`. MATRIX ORDER: Keep dimensions inside \`{ EQ }\` immediately after the bracket using \`\\s\\do\` based on row count (2 rows=\`\\s\\do8\`, 3=\`\\s\\do14\`, 4=\`\\s\\do18\`, 5=\`\\s\\do24\`; e.g., \`{ EQ \\b\\bc\\[((\\a...))\\s\\do24(5×4) }\`). Vectors = \`{ EQ \\o(BD,\\s\\up7(⟶))}\`. Limits = \`{ EQ \\o(lim,\\s\\do8(x⟶0))}\`. Factorial = \`{ EQ \\x\\le\\bo(n)}\`.`;
}

function copyPromptToClipboard() {
    navigator.clipboard.writeText(getSystemPrompt());
    setStatus("📋 Prompt copied to clipboard!", "#d1fae5", "#065f46");
}

// ==========================================
// SMART MODEL RESOLVER (3.8 Flash, 3.1 Pro, 4 Argon)
// ==========================================
function getCandidateEndpoints(selectedTier) {
    const tierMap = {
        "3.8-flash": [
            "gemini-3.8-flash",
            "gemini-3.0-flash",
            "gemini-2.5-flash",
            "gemini-2.0-flash"
        ],
        "3.1-pro": [
            "gemini-3.1-pro",
            "gemini-3.1-pro-preview",
            "gemini-2.5-pro",
            "gemini-2.5-flash"
        ],
        "4-argon": [
            "gemini-4-argon",
            "gemini-exp-1206",
            "gemini-3.1-pro",
            "gemini-2.5-pro",
            "gemini-2.5-flash"
        ]
    };
    return tierMap[selectedTier] || tierMap["3.8-flash"];
}

async function extractFromImageWithAI() {
    if (!currentImageBase64) return;

    const rawKeys = document.getElementById("apiKeysInput").value.trim();
    const keys = rawKeys.split(/[\n,]+/).map(k => k.trim()).filter(Boolean);

    if (keys.length === 0) {
        setStatus("⚠️ অনুগ্রহ করে উপরে Settings-এ আপনার Gemini API Key দিন!", "#fee2e2", "#991b1b");
        document.getElementById("apiSettings").open = true;
        return;
    }

    const selectedTier = document.getElementById("modelSelect").value;
    const tierDisplayName = document.getElementById("activeModelLabel").innerText;
    const candidateModels = getCandidateEndpoints(selectedTier);

    const extractBtn = document.getElementById("extractBtn");
    extractBtn.disabled = true;

    const promptText = getSystemPrompt();

    for (let kIndex = 0; kIndex < keys.length; kIndex++) {
        const apiKey = keys[kIndex];
        for (let mIndex = 0; mIndex < candidateModels.length; mIndex++) {
            const modelId = candidateModels[mIndex];
            setStatus(`⏳ Extracting via ${tierDisplayName} (Key #${kIndex + 1})...`, "#fef3c7", "#92400e");

            try {
                const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${apiKey}`;
                const response = await fetch(url, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        contents: [{
                            parts: [
                                { text: promptText },
                                { inline_data: { mime_type: currentImageMime, data: currentImageBase64 } }
                            ]
                        }],
                        generationConfig: { temperature: 0.1 }
                    })
                });

                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}`);
                }

                const data = await response.json();
                let extracted = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
                extracted = extracted.replace(/^```[a-zA-Z]*\n?/gm, "").replace(/```$/gm, "").trim();

                document.getElementById("rawTextInput").value = extracted;
                setStatus(`✅ Extracted with ${tierDisplayName}! Inserting...`, "#d1fae5", "#065f46");
                extractBtn.disabled = false;

                await formatAndInsertToWord();
                return;

            } catch (err) {
                console.warn(`Attempt failed for ${modelId} (Key ${kIndex + 1}):`, err);
            }
        }
    }

    extractBtn.disabled = false;
    setStatus("❌ সব API Key ব্যস্ত বা লিমিট শেষ! ম্যানুয়ালি পেস্ট করুন।", "#fee2e2", "#991b1b");
}

// ==========================================
// PARVEZ VBA VERSION 2 -> OOXML ENGINE (100% EXACT)
// ==========================================
function xmlEscape(str) {
    return str
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function buildRunProps(opts = {}) {
    const forceTNR = opts.forceTNR || false;
    const csFont = forceTNR ? "Times New Roman" : "Kalpurush";
    const sizeHalfPts = Math.round((opts.size || 10.5) * 2);
    const italicTag = opts.italic ? "<w:i/><w:iCs/>" : "";
    const posTag = opts.position ? `<w:position w:val="${Math.round(opts.position * 2)}"/>` : "";

    return `<w:rPr>` +
        `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="${csFont}"/>` +
        italicTag +
        posTag +
        `<w:sz w:val="${sizeHalfPts}"/>` +
        `<w:szCs w:val="${sizeHalfPts}"/>` +
        `</w:rPr>`;
}

function generateTextRuns(text, baseOpts = {}, isInstrText = false) {
    if (!text) return "";

    const uprightIndices = new Set();
    const funcRegex = /\b(lim|log|ln)\b/g;
    let match;
    while ((match = funcRegex.exec(text)) !== null) {
        for (let i = match.index; i < match.index + match[0].length; i++) {
            uprightIndices.add(i);
        }
    }

    const groups = [];
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        const isItalicL = (ch === "l" && !uprightIndices.has(i));
        const isSpecialTNR = (ch === "\u2013" || ch === "\u00D7" || ch === "\u222B");

        const italic = isItalicL ? true : Boolean(baseOpts.italic);
        const forceTNR = isSpecialTNR ? true : Boolean(baseOpts.forceTNR);
        const size = baseOpts.size || 10.5;
        const position = baseOpts.position || 0;

        const key = `${size}|${italic}|${forceTNR}|${position}`;
        if (groups.length > 0 && groups[groups.length - 1].key === key) {
            groups[groups.length - 1].chars += ch;
        } else {
            groups.push({ key, chars: ch, opts: { size, italic, forceTNR, position } });
        }
    }

    const tag = isInstrText ? "w:instrText" : "w:t";
    return groups.map(g => {
        return `<w:r>${buildRunProps(g.opts)}<${tag} xml:space="preserve">${xmlEscape(g.chars)}</${tag}></w:r>`;
    }).join("");
}

function buildEqFieldOoxml(eqCode) {
    const n = eqCode.length;
    const charSize = new Array(n).fill(10.5);
    const charForceTNR = new Array(n).fill(false);

    const shiftRegex = /\\s\\(?:up(?:4|7|10)|do(?:2|3|8|14|18|24))\([^)]*\)/gi;
    let m;
    while ((m = shiftRegex.exec(eqCode)) !== null) {
        for (let i = m.index; i < m.index + m[0].length; i++) {
            charSize[i] = 7.5;
        }
    }

    let searchPos = 0;
    while (searchPos < n) {
        const idx = eqCode.toLowerCase().indexOf("\\s(", searchPos);
        if (idx === -1) break;
        const openParen = idx + 2;
        let pDepth = 0;
        let closeParen = -1;

        for (let i = openParen; i < n; i++) {
            if (eqCode[i] === "(") pDepth++;
            else if (eqCode[i] === ")") {
                pDepth--;
                if (pDepth === 0) {
                    closeParen = i;
                    break;
                }
            }
        }

        if (closeParen > openParen) {
            const innerStr = eqCode.substring(openParen + 1, closeParen);
            if (innerStr.includes("\u222B")) {
                for (let i = openParen + 1; i < closeParen; i++) {
                    if (eqCode[i] === "\u222B") {
                        charSize[i] = 13.5;
                        charForceTNR[i] = true;
                    } else {
                        charSize[i] = 7.5;
                    }
                }
            }
            searchPos = closeParen + 1;
        } else {
            break;
        }
    }

    let instrRuns = "";
    let segStart = 0;
    for (let i = 1; i <= n; i++) {
        if (i === n || charSize[i] !== charSize[segStart] || charForceTNR[i] !== charForceTNR[segStart]) {
            const substr = eqCode.substring(segStart, i);
            instrRuns += generateTextRuns(substr, {
                size: charSize[segStart],
                forceTNR: charForceTNR[segStart]
            }, true);
            segStart = i;
        }
    }

    return `<w:r><w:fldChar w:fldCharType="begin" w:fldLock="true" w:dirty="true"/></w:r>` +
           instrRuns +
           `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
           `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

function buildOutsideTextOoxml(text) {
    if (!text) return "";
    const tokenRegex = /\\s\\(up0|do0)\(([^()]*)\)/g;
    let lastIdx = 0;
    let outXml = "";
    let m;

    while ((m = tokenRegex.exec(text)) !== null) {
        if (m.index > lastIdx) {
            outXml += generateTextRuns(text.substring(lastIdx, m.index), { size: 10.5 });
        }
        const type = m[1].toLowerCase();
        const content = m[2];
        if (type === "up0") {
            outXml += generateTextRuns(content, { size: 7.5, position: 3.5 });
        } else {
            outXml += generateTextRuns(content, { size: 7.5, position: -2 });
        }
        lastIdx = m.index + m[0].length;
    }

    if (lastIdx < text.length) {
        outXml += generateTextRuns(text.substring(lastIdx), { size: 10.5 });
    }
    return outXml;
}

function convertRawTextToWordOoxml(rawText) {
    let text = rawText.replace(/-/g, "\u2013");
    text = text.replace(/['\u2019]/g, "\u2032");
    text = text.replace(/\\c?dot/g, "\u00B7");
    text = text.replace(/\{\s*eq\s+/gi, "{ EQ ");

    const lines = text.split(/\r?\n/);
    let paragraphsXml = "";

    for (let line of lines) {
        if (!line.trim()) {
            paragraphsXml += `<w:p><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:p>`;
            continue;
        }

        let pRuns = "";
        const eqRegex = /\{\s*EQ\s+([^}]+)\}/g;
        let lastIndex = 0;
        let match;

        while ((match = eqRegex.exec(line)) !== null) {
            if (match.index > lastIndex) {
                pRuns += buildOutsideTextOoxml(line.substring(lastIndex, match.index));
            }
            const eqInner = " EQ " + match[1] + " ";
            pRuns += buildEqFieldOoxml(eqInner);
            lastIndex = match.index + match[0].length;
        }

        if (lastIndex < line.length) {
            pRuns += buildOutsideTextOoxml(line.substring(lastIndex));
        }

        paragraphsXml += `<w:p><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>${pRuns}</w:p>`;
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>${paragraphsXml}</w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

async function formatAndInsertToWord() {
    const rawText = document.getElementById("rawTextInput").value;
    if (!rawText.trim()) {
        setStatus("⚠️ বক্সে কোনো টেক্সট নেই!", "#fee2e2", "#991b1b");
        return;
    }

    try {
        setStatus("⚙️ Formatting & Inserting into Word...", "#fef3c7", "#92400e");
        const ooxml = convertRawTextToWordOoxml(rawText);

        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.insertOoxml(ooxml, Word.InsertLocation.replace);
            await context.sync();
        });

        setStatus("🎉 সফলভাবে MS Word-এ ফরম্যাট হয়ে বসে গেছে!", "#d1fae5", "#065f46");
    } catch (error) {
        console.error(error);
        setStatus("❌ Word Error: " + error.message, "#fee2e2", "#991b1b");
    }
}