/* global Office, Word */

let currentImageBase64 = null;
let currentImageMime = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initApp();
    } else {
        initApp(); // Fallback for browser testing
    }
});

function initApp() {
    // Load saved API keys
    const savedKeys = localStorage.getItem("parvez_gemini_keys") || "";
    const savedModel = localStorage.getItem("parvez_gemini_model") || "gemini-2.5-flash";
    document.getElementById("apiKeysInput").value = savedKeys;
    document.getElementById("modelSelect").value = savedModel;

    if (!savedKeys) {
        document.getElementById("apiSettings").open = true;
    }

    document.getElementById("saveKeysBtn").onclick = () => {
        localStorage.setItem("parvez_gemini_keys", document.getElementById("apiKeysInput").value.trim());
        localStorage.setItem("parvez_gemini_model", document.getElementById("modelSelect").value);
        setStatus("✅ API Settings saved!", "#d1fae5", "#065f46");
        document.getElementById("apiSettings").open = false;
    };

    // Image Upload & Paste Handling
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");

    dropZone.onclick = () => fileInput.click();
    fileInput.onchange = (e) => {
        if (e.target.files && e.target.files[0]) {
            loadImageFile(e.target.files[0]);
        }
    };

    // Global Paste (Ctrl+V) for screenshots
    window.addEventListener("paste", (e) => {
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
    document.getElementById("copyPromptBtn").onclick = copyPromptToClipboard;
    document.getElementById("insertFormattedBtn").onclick = formatAndInsertToWord;
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
// MULTI-KEY & MULTI-MODEL AI ENGINE
// ==========================================
async function extractFromImageWithAI() {
    if (!currentImageBase64) return;

    const rawKeys = document.getElementById("apiKeysInput").value.trim();
    const keys = rawKeys.split(/[\n,]+/).map(k => k.trim()).filter(Boolean);

    if (keys.length === 0) {
        setStatus("⚠️ অনুগ্রহ করে উপরে Settings-এ আপনার Gemini API Key দিন!", "#fee2e2", "#991b1b");
        document.getElementById("apiSettings").open = true;
        return;
    }

    const primaryModel = document.getElementById("modelSelect").value;
    const fallbackModels = [primaryModel, "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"].filter((v, i, a) => a.indexOf(v) === i);

    const extractBtn = document.getElementById("extractBtn");
    extractBtn.disabled = true;

    const promptText = getSystemPrompt();

    for (let mIndex = 0; mIndex < fallbackModels.length; mIndex++) {
        const model = fallbackModels[mIndex];
        for (let kIndex = 0; kIndex < keys.length; kIndex++) {
            const apiKey = keys[kIndex];
            setStatus(`⏳ Extracting via ${model} (Key #${kIndex + 1})...`, "#fef3c7", "#92400e");

            try {
                const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
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
                // Clean markdown codeblocks if AI accidentally includes them
                extracted = extracted.replace(/^```[a-zA-Z]*\n?/gm, "").replace(/```$/gm, "").trim();

                document.getElementById("rawTextInput").value = extracted;
                setStatus("✅ Extracted! Formatting & Inserting into Word...", "#d1fae5", "#065f46");
                extractBtn.disabled = false;

                // Automatically format and insert into Word!
                await formatAndInsertToWord();
                return;

            } catch (err) {
                console.warn(`Failed with ${model} (Key ${kIndex + 1}):`, err);
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
    const sizeHalfPts = Math.round((opts.size || 10.5) * 2); // 10.5pt = 21, 7.5pt = 15, 13.5pt = 27
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

// Handles Step 4.1, 5.1, 5.2, 5.3, 5.4 (l italic, lim/log/ln upright, En-Dash, ×, ∫ force TNR)
function generateTextRuns(text, baseOpts = {}, isInstrText = false) {
    if (!text) return "";

    // Find all upright math functions ('lim', 'log', 'ln') so their 'l' is NOT italicized
    const uprightIndices = new Set();
    const funcRegex = /\b(lim|log|ln)\b/g;
    let match;
    while ((match = funcRegex.exec(text)) !== null) {
        for (let i = match.index; i < match.index + match[0].length; i++) {
            uprightIndices.add(i);
        }
    }

    // Group adjacent characters with identical formatting
    const groups = [];
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        const isItalicL = (ch === "l" && !uprightIndices.has(i));
        const isSpecialTNR = (ch === "\u2013" || ch === "\u00D7" || ch === "\u222B"); // –, ×, ∫

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

// Formats inside { EQ ... } according to Step 4.2, 4.3, and Step 5.5
function buildEqFieldOoxml(eqCode) {
    const n = eqCode.length;
    const charSize = new Array(n).fill(10.5);
    const charForceTNR = new Array(n).fill(false);

    // STEP 4.2 & 4.3: \s\up(4, 7, 10) and \s\do(2, 3, 8, 14, 18, 24) -> Set Size to 7.5
    const shiftRegex = /\\s\\(?:up(?:4|7|10)|do(?:2|3|8|14|18|24))\([^)]*\)/gi;
    let m;
    while ((m = shiftRegex.exec(eqCode)) !== null) {
        for (let i = m.index; i < m.index + m[0].length; i++) {
            charSize[i] = 7.5;
        }
    }

    // STEP 5.5: Definite integrals \s(upper,∫,lower) -> limits 7.5pt, integral sign 13.5pt
    let searchPos = 0;
    while (searchPos < n) {
        const idx = eqCode.toLowerCase().indexOf("\\s(", searchPos);
        if (idx === -1) break;
        const openParen = idx + 2; // Position of '('
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

    // Group segments by size and forceTNR, then pass to generateTextRuns
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

    // Step 8: Lock field (w:fldLock="true") and mark dirty="true" so Word renders it immediately
    return `<w:r><w:fldChar w:fldCharType="begin" w:fldLock="true" w:dirty="true"/></w:r>` +
           instrRuns +
           `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
           `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

// Formats outside EQ text according to Step 6 (\s\up0) & Step 7 (\s\do0)
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
            // Step 6: Size 7.5, Position +3.5pt
            outXml += generateTextRuns(content, { size: 7.5, position: 3.5 });
        } else {
            // Step 7: Size 7.5, Position -2pt
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
    // STEP 1: Replace standard hyphens (-) with En Dash (\u2013)
    let text = rawText.replace(/-/g, "\u2013");

    // STEP 1.1: Replace Single Quote / Prime (') and ’ with True Math Prime (\u2032)
    text = text.replace(/['\u2019]/g, "\u2032");

    // STEP 1.2: Remove Accidental LaTeX \dot and \cdot -> Unicode middle dot (\u00B7)
    text = text.replace(/\\c?dot/g, "\u00B7");

    // STEP 2.5: Normalize {eq or {EQ to { EQ
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