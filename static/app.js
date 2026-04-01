(() => {
    "use strict";

    const STORAGE_KEY = "montebello_ot_state";
    const THEME_KEY = "app_theme";
    const CAT_LABELS = { ot10: "OT 1.0", ot15: "OT 1.5", cte10: "CTE 1.0", cte15: "CTE 1.5" };
    const CATS = ["ot10", "ot15", "cte10", "cte15"];

    function entryHours(e) {
        if (e.kind === "weekBlock") {
            return CATS.reduce((s, c) => s + (Number(e[c]) || 0), 0);
        }
        return Number(e.hours) || 0;
    }
    function fmtWeekBlockLine(e) {
        const parts = [];
        CATS.forEach(c => {
            const v = Number(e[c]) || 0;
            if (v > 0) parts.push(`${CAT_LABELS[c]} ${v.toFixed(2)}`);
        });
        return parts.join(", ");
    }
    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }
    function buildOtPayloadFromEntries() {
        const otByEmp = {};
        state.otEntries.forEach(e => {
            if (!otByEmp[e.empNo]) otByEmp[e.empNo] = { entries: [], weekBlocks: [] };
            if (e.kind === "weekBlock") {
                const row = { week: e.week, rangeText: (e.rangeText || "").trim() };
                CATS.forEach(c => { row[c] = Math.round((Number(e[c]) || 0) * 100) / 100; });
                otByEmp[e.empNo].weekBlocks.push(row);
            } else {
                otByEmp[e.empNo].entries.push({
                    date: e.date, category: e.category, hours: e.hours,
                });
            }
        });
        return otByEmp;
    }
    function resetWeekBlockForm() {
        if (!$otRangeText) return;
        $otRangeText.value = "";
        const w1 = document.querySelector("#ot-week-1");
        if (w1) w1.checked = true;
        document.querySelectorAll(".ot-wb-hrs").forEach(el => { el.value = ""; });
    }


    // -----------------------------------------------------------------------
    // Theme
    // -----------------------------------------------------------------------
    function initTheme() {
        const saved = localStorage.getItem(THEME_KEY) || "default";
        applyTheme(saved);
    }
    function applyTheme(name) {
        document.documentElement.setAttribute("data-theme", name);
        localStorage.setItem(THEME_KEY, name);
        document.querySelectorAll(".theme-btn").forEach(b => {
            b.classList.toggle("active", b.dataset.theme === name);
        });
    }
    document.addEventListener("click", (e) => {
        const btn = e.target.closest(".theme-btn");
        if (btn) applyTheme(btn.dataset.theme);
    });

    // -----------------------------------------------------------------------
    // State
    // -----------------------------------------------------------------------
    let state = { employees: [], payPeriodEnd: "", otEntries: [] };
    let activeEmpNo = null;

    function saveState() {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    }
    function loadState() {
        try {
            const s = JSON.parse(localStorage.getItem(STORAGE_KEY));
            if (s) {
                state.employees = s.employees || [];
                state.payPeriodEnd = s.payPeriodEnd || "";
                state.otEntries = s.otEntries || [];
            }
        } catch (_) {}
    }

    // -----------------------------------------------------------------------
    // DOM
    // -----------------------------------------------------------------------
    const $ = (id) => document.getElementById(id);

    const $csvUpload = $("csv-upload");
    const $csvStatus = $("csv-status");
    const $btnClearCsv = $("btn-clear-csv");
    const $ppEndDate = $("pp-end-date");
    const $ppInfo = $("pp-info");
    const $empSummary = $("employee-summary");
    const $empCount = $("emp-count");

    const $tabs = document.querySelectorAll(".tab");
    const $tabContents = document.querySelectorAll(".tab-content");

    const $btnGenerateSlips = $("btn-generate-slips");
    const $slipsStatus = $("slips-status");

    const $empSearch = $("emp-search");
    const $empList = $("emp-list");
    const $empSelectedNo = $("emp-selected-no");
    const $otDate = $("ot-date");
    const $otCategory = $("ot-category");
    const $otHours = $("ot-hours");
    const $btnAddOt = $("btn-add-ot");
    const $otCardsWrap = $("ot-cards-wrap");
    const $otCards = $("ot-cards");
    const $otEntryCount = $("ot-entry-count");
    const $otSummaryWrap = $("ot-summary-wrap");
    const $otSummaryBody = $("ot-summary-body");
    const $wk1Range = $("wk1-range");
    const $wk2Range = $("wk2-range");
    const $otWeekInfo = $("ot-week-info");

    const $btnGenerateOt = $("btn-generate-ot");
    const $btnClearSession = $("btn-clear-session");
    const $otStatus = $("ot-status");
    const $activeCardStatus = $("active-card-status");
    const $importCorrections = $("import-corrections");
    const $importStatus = $("import-status");
    const $btnImportParse = $("btn-import-parse");
    const $btnImportGenerate = $("btn-import-generate");
    const $importSummary = $("import-summary");
    const $importResult = $("import-result");

    const $confirmModal = $("confirm-modal");
    const $modalMessage = $("modal-message");
    const $modalCancel = $("modal-cancel");
    const $modalConfirm = $("modal-confirm");

    const $activeOverlay = $("active-overlay");
    const $activeCard = $("active-card");
    const $activeName = $("active-emp-name");
    const $activeEntries = $("active-emp-entries");
    const $btnDoneEmp = $("btn-done-emp");
    const $entryFields = $("entry-fields-row");
    const $otRangeText = $("ot-range-text");
    const $btnAddWeekBlock = $("btn-add-week-block");


    // Collapsible toggle headers
    const $cardsToggle = $("cards-toggle");
    const $cardsBody = $("cards-body");
    const $summaryToggle = $("summary-toggle");
    const $summaryBody = $("summary-body");

    // -----------------------------------------------------------------------
    // Confirm modal
    // -----------------------------------------------------------------------
    let modalCallback = null;

    function showConfirm(message, btnText, onConfirm) {
        $modalMessage.textContent = message;
        $modalConfirm.textContent = btnText;
        modalCallback = onConfirm;
        $confirmModal.classList.remove("hidden");
    }
    $modalCancel.addEventListener("click", () => {
        $confirmModal.classList.add("hidden");
        modalCallback = null;
    });
    $modalConfirm.addEventListener("click", () => {
        $confirmModal.classList.add("hidden");
        if (modalCallback) modalCallback();
        modalCallback = null;
    });

    // -----------------------------------------------------------------------
    // Collapsible sections
    // -----------------------------------------------------------------------
    function initCollapsible(toggle, body) {
        let collapsed = false;
        toggle.addEventListener("click", () => {
            collapsed = !collapsed;
            body.classList.toggle("hidden", collapsed);
            toggle.classList.toggle("collapsed", collapsed);
        });
    }
    initCollapsible($cardsToggle, $cardsBody);
    initCollapsible($summaryToggle, $summaryBody);

    // About section — starts collapsed, uses .open class
    const $aboutToggle = $("about-toggle");
    const $aboutBody = $("about-body");
    if ($aboutToggle && $aboutBody) {
        let aboutOpen = false;
        $aboutToggle.addEventListener("click", () => {
            aboutOpen = !aboutOpen;
            $aboutBody.classList.toggle("open", aboutOpen);
            $aboutToggle.classList.toggle("collapsed", !aboutOpen);
        });
        $aboutToggle.classList.add("collapsed");
    }

    // -----------------------------------------------------------------------
    // Date helpers
    // -----------------------------------------------------------------------
    function payPeriodWeeks(endStr) {
        const end = new Date(endStr + "T00:00:00");
        const wk1Start = new Date(end); wk1Start.setDate(end.getDate() - 13);
        const wk1End   = new Date(end); wk1End.setDate(end.getDate() - 7);
        const wk2Start = new Date(end); wk2Start.setDate(end.getDate() - 6);
        return { wk1Start, wk1End, wk2Start, wk2End: end };
    }
    function fmtShort(d) { return `${d.getMonth()+1}/${d.getDate()}`; }
    function fmtFull(d) { return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`; }
    function toISO(d) { return d.toISOString().split("T")[0]; }

    function dateToWeek(dateStr, ppEnd) {
        if (!ppEnd) return "?";
        const { wk1Start, wk1End, wk2Start, wk2End } = payPeriodWeeks(ppEnd);
        const d = new Date(dateStr + "T00:00:00");
        if (d >= wk1Start && d <= wk1End) return "1";
        if (d >= wk2Start && d <= wk2End) return "2";
        return "?";
    }

    function uniqueEmpCount() {
        return new Set(state.otEntries.map(e => e.empNo)).size;
    }

    // -----------------------------------------------------------------------
    // Tabs
    // -----------------------------------------------------------------------
    $tabs.forEach(tab => {
        tab.addEventListener("click", () => {
            $tabs.forEach(t => t.classList.remove("active"));
            $tabContents.forEach(tc => tc.classList.remove("active"));
            tab.classList.add("active");
            $("tab-" + tab.dataset.tab).classList.add("active");
        });
    });

    // -----------------------------------------------------------------------
    // CSV upload + clear
    // -----------------------------------------------------------------------
    $csvUpload.addEventListener("change", async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        $csvStatus.textContent = "Uploading...";
        $csvStatus.className = "file-status";

        const formData = new FormData();
        formData.append("file", file);

        try {
            const resp = await fetch("/api/parse-csv", { method: "POST", body: formData });
            const data = await resp.json();
            if (data.error) throw new Error(data.error);
            state.employees = data.employees;
            saveState();
            renderEmployeeState();
        } catch (err) {
            $csvStatus.textContent = "Error: " + err.message;
        }
    });

    $btnClearCsv.addEventListener("click", () => {
        showConfirm(
            "Are you sure you want to remove the employee list? You can upload a new one after.",
            "Yes, Remove",
            () => {
                state.employees = [];
                $csvUpload.value = "";
                saveState();
                renderEmployeeState();
            }
        );
    });

    // -----------------------------------------------------------------------
    // Import Corrections (standalone tab — does NOT merge into manual OT)
    // -----------------------------------------------------------------------
    let importedEntries = [];

    $importCorrections.addEventListener("change", () => {
        const file = $importCorrections.files[0];
        if (file) {
            $importStatus.textContent = file.name;
            $btnImportParse.disabled = !(state.employees.length && state.payPeriodEnd);
        } else {
            $importStatus.textContent = "No file chosen";
            $btnImportParse.disabled = true;
            $btnImportGenerate.disabled = true;
            $importSummary.classList.add("hidden");
        }
        importedEntries = [];
        hideStatus($importResult);
    });

    $btnImportParse.addEventListener("click", async () => {
        const file = $importCorrections.files[0];
        if (!file || !state.employees.length || !state.payPeriodEnd) {
            showStatus($importResult, "Upload employee list and set pay period in Setup first.", "error");
            return;
        }
        showStatus($importResult, "Parsing...", "loading");
        $btnImportParse.disabled = true;
        try {
            const formData = new FormData();
            formData.append("file", file);
            formData.append("employees", JSON.stringify(state.employees));
            formData.append("payPeriodEnd", state.payPeriodEnd);
            const resp = await fetch("/api/import-corrections", { method: "POST", body: formData });
            const data = await resp.json();
            if (!resp.ok) throw new Error(data.error || "Parse failed");
            importedEntries = data.entries || [];
            const unmatched = data.unmatched || [];
            const empCount = new Set(importedEntries.map(e => e.empNo)).size;
            $importSummary.classList.remove("hidden");
            $importSummary.innerHTML = `<strong>${importedEntries.length}</strong> entries for <strong>${empCount}</strong> employees.` +
                (unmatched.length ? ` <br>${unmatched.length} without emp # (name + date + OT only): ${unmatched.slice(0, 5).join(", ")}${unmatched.length > 5 ? "…" : ""}.` : "");
            $btnImportGenerate.disabled = importedEntries.length === 0;
            showStatus($importResult, "Ready to generate. Click Generate OT Slips + Excel.", "success");
        } catch (err) {
            showStatus($importResult, "Error: " + err.message, "error");
        } finally {
            $btnImportParse.disabled = false;
        }
    });

    $btnImportGenerate.addEventListener("click", async () => {
        if (!importedEntries.length || !state.payPeriodEnd) return;
        startElapsedTimer($importResult, "Generating PDF and Excel...");
        $btnImportGenerate.disabled = true;
        try {
            const otByEmp = {};
            importedEntries.forEach(e => {
                if (!otByEmp[e.empNo]) otByEmp[e.empNo] = { entries: [] };
                otByEmp[e.empNo].entries.push({ date: e.date, category: e.category, hours: e.hours });
            });
            const employees = [...state.employees];
            const seenUm = new Set();
            importedEntries.forEach(e => {
                if (e.empNo.startsWith("__UM__") && !seenUm.has(e.empNo)) {
                    seenUm.add(e.empNo);
                    employees.push({ emp_no: e.empNo, last: e.last, first: e.first });
                }
            });
            const ctrl = new AbortController();
            const t = setTimeout(() => ctrl.abort(), 120000);
            const resp = await fetch("/api/generate-overtime", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ employees, payPeriodEnd: state.payPeriodEnd, otEntries: otByEmp }),
                signal: ctrl.signal,
            });
            clearTimeout(t);
            const data = await resp.json();
            if (data.error) throw new Error(data.error);
            const pdfBytes = Uint8Array.from(atob(data.pdf), c => c.charCodeAt(0));
            downloadBlob(new Blob([pdfBytes], { type: "application/pdf" }), data.pdfFilename);
            const xlBytes = Uint8Array.from(atob(data.excel), c => c.charCodeAt(0));
            downloadBlob(new Blob([xlBytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), data.excelFilename);
            stopElapsedTimer($importResult);
            showStatus($importResult, "PDF and Excel downloaded.", "success");
        } catch (err) {
            stopElapsedTimer($importResult);
            const msg = err.name === "AbortError" ? "Request timed out — try again." : err.message;
            showStatus($importResult, "Error: " + msg, "error");
        } finally {
            $btnImportGenerate.disabled = false;
        }
    });

    function renderEmployeeState() {
        if (state.employees.length === 0) {
            $csvStatus.textContent = "No file loaded";
            $csvStatus.className = "file-status";
            $empSummary.classList.add("hidden");
            $btnClearCsv.classList.add("hidden");
            setOtControlsEnabled(false);
            $btnGenerateSlips.disabled = true;
            return;
        }
        $csvStatus.textContent = `${state.employees.length} employees loaded`;
        $csvStatus.className = "file-status loaded";
        $empCount.textContent = state.employees.length;
        $empSummary.classList.remove("hidden");
        $btnClearCsv.classList.remove("hidden");
        updateButtonStates();
    }

    // -----------------------------------------------------------------------
    // Pay period date
    // -----------------------------------------------------------------------
    $ppEndDate.addEventListener("change", () => {
        state.payPeriodEnd = $ppEndDate.value;
        saveState();
        updatePayPeriodInfo();
        updateButtonStates();
        renderAll();
    });

    function updatePayPeriodInfo() {
        if (!state.payPeriodEnd) {
            $ppInfo.textContent = "";
            $otWeekInfo.classList.add("hidden");
            return;
        }
        const { wk1Start, wk1End, wk2Start, wk2End } = payPeriodWeeks(state.payPeriodEnd);
        const d = new Date(state.payPeriodEnd + "T00:00:00");
        if (d.getDay() !== 6) {
            $ppInfo.textContent = "Note: this date is not a Saturday.";
            $ppInfo.style.color = "#dc2626";
        } else {
            $ppInfo.textContent = "";
            $ppInfo.style.color = "";
        }
        $wk1Range.textContent = `${fmtShort(wk1Start)} – ${fmtFull(wk1End)}`;
        $wk2Range.textContent = `${fmtShort(wk2Start)} – ${fmtFull(wk2End)}`;
        $otWeekInfo.classList.remove("hidden");
        $otDate.min = toISO(wk1Start);
        $otDate.max = state.payPeriodEnd;
    }

    // -----------------------------------------------------------------------
    // Combobox (searchable employee select)
    // -----------------------------------------------------------------------
    let comboHighlight = -1;
    let filteredEmps = [];

    $empSearch.addEventListener("focus", () => { $empSearch.select(); openCombo(); });
    $empSearch.addEventListener("input", () => openCombo());

    $empSearch.addEventListener("keydown", (e) => {
        if (!$empList.classList.contains("hidden")) {
            if (e.key === "ArrowDown") {
                e.preventDefault();
                comboHighlight = Math.min(comboHighlight + 1, filteredEmps.length - 1);
                renderComboHighlight();
                return;
            } else if (e.key === "ArrowUp") {
                e.preventDefault();
                comboHighlight = Math.max(comboHighlight - 1, 0);
                renderComboHighlight();
                return;
            } else if (e.key === "Escape") {
                closeCombo();
                return;
            }
        }

        if (e.key === "Enter") {
            e.preventDefault();
            // If arrow-highlighted, select that one; otherwise select first match
            if (comboHighlight >= 0 && comboHighlight < filteredEmps.length) {
                selectEmployee(filteredEmps[comboHighlight]);
            } else if (filteredEmps.length > 0) {
                selectEmployee(filteredEmps[0]);
            }
        }
    });

    document.addEventListener("click", (e) => {
        if (!e.target.closest("#emp-combobox")) closeCombo();
    });

    function openCombo() {
        const q = $empSearch.value.toLowerCase();
        filteredEmps = state.employees.filter(emp => {
            const full = `${emp.last}, ${emp.first} ${emp.emp_no}`.toLowerCase();
            return full.includes(q);
        });
        comboHighlight = -1;

        $empList.innerHTML = filteredEmps.slice(0, 50).map((emp, i) => {
            const hasEntries = state.otEntries.some(e => e.empNo === emp.emp_no);
            const checkmark = hasEntries ? '<span class="emp-done">&#10003;</span>' : '';
            return `<div class="combobox-item" data-idx="${i}">` +
                `${checkmark}${emp.last}, ${emp.first} <span class="emp-num">#${emp.emp_no}</span></div>`;
        }).join("");

        $empList.querySelectorAll(".combobox-item").forEach(el => {
            el.addEventListener("mousedown", (e) => {
                e.preventDefault();
                selectEmployee(filteredEmps[parseInt(el.dataset.idx)]);
            });
        });

        $empList.classList.remove("hidden");
    }

    function closeCombo() { $empList.classList.add("hidden"); }

    function renderComboHighlight() {
        $empList.querySelectorAll(".combobox-item").forEach((el, i) => {
            el.classList.toggle("highlighted", i === comboHighlight);
        });
        const active = $empList.querySelector(".highlighted");
        if (active) active.scrollIntoView({ block: "nearest" });
    }

    function selectEmployee(emp) {
        $empSearch.value = `${emp.last}, ${emp.first} (#${emp.emp_no})`;
        $empSelectedNo.value = emp.emp_no;
        activeEmpNo = emp.emp_no;
        closeCombo();
        showActiveCard();
    }

    // -----------------------------------------------------------------------
    // Active employee card (overlay pop-out)
    // -----------------------------------------------------------------------
    function showActiveCard() {
        if (!activeEmpNo) return;

        const emp = state.employees.find(e => e.emp_no === activeEmpNo);
        if (!emp) return;

        $activeName.innerHTML = `<strong>${emp.last}, ${emp.first}</strong> <span class="emp-card-meta">#${emp.emp_no}</span>`;
        if (state.payPeriodEnd && !$otDate.value) {
            const { wk1Start } = payPeriodWeeks(state.payPeriodEnd);
            $otDate.value = toISO(wk1Start);
        }
        renderActiveEntries();
        hideStatus($activeCardStatus);
        $activeOverlay.classList.remove("hidden");

        setTimeout(() => ($otRangeText || $otDate).focus(), 50);
    }

    function closeActiveCard() {
        activeEmpNo = null;
        $empSelectedNo.value = "";
        $empSearch.value = "";
        $activeOverlay.classList.add("hidden");
        $otDate.value = "";
        $otHours.value = "";
        resetWeekBlockForm();
        hideStatus($activeCardStatus);
        renderCompletedCards();
        $empSearch.focus();
    }

    function renderActiveEntries() {
        const entries = state.otEntries
            .map((e, idx) => ({ ...e, _idx: idx }))
            .filter(e => e.empNo === activeEmpNo)
            .sort((a, b) => {
                const ka = a.kind === "weekBlock" ? `0-${a.week}-${a.rangeText}` : `1-${a.date}-${a.category}`;
                const kb = b.kind === "weekBlock" ? `0-${b.week}-${b.rangeText}` : `1-${b.date}-${b.category}`;
                return ka.localeCompare(kb);
            });

        if (entries.length === 0) {
            $activeEntries.innerHTML = '<p class="active-empty">No entries yet. Add overtime below.</p>';
            return;
        }

        const totalHrs = entries.reduce((s, e) => s + entryHours(e), 0);

        $activeEntries.innerHTML = entries.map(e => {
            if (e.kind === "weekBlock") {
                return `<div class="active-entry-row active-entry-weekblock">
                <span class="row-wk">Wk ${e.week}</span>
                <span class="row-date range-text">${escapeHtml(e.rangeText || "")}</span>
                <span class="row-cat">${fmtWeekBlockLine(e)}</span>
                <span class="row-hrs">${entryHours(e).toFixed(2)} hrs</span>
                <button class="btn-remove" data-idx="${e._idx}" title="Remove">&times;</button>
            </div>`;
            }
            const d = new Date(e.date + "T00:00:00");
            const wk = dateToWeek(e.date, state.payPeriodEnd);
            return `<div class="active-entry-row">
                <span class="row-wk">Wk ${wk}</span>
                <span class="row-date">${fmtShort(d)}</span>
                <span class="row-cat">${CAT_LABELS[e.category]}</span>
                <span class="row-hrs">${Number(e.hours).toFixed(2)} hrs</span>
                <button class="btn-remove" data-idx="${e._idx}" title="Remove">&times;</button>
            </div>`;
        }).join("") + `<div class="active-total">Total: ${Number(totalHrs).toFixed(2)} hrs</div>`;

        $activeEntries.querySelectorAll(".btn-remove").forEach(btn => {
            btn.addEventListener("click", () => {
                removeEntry(parseInt(btn.dataset.idx));
                renderActiveEntries();
            });
        });
    }

    $btnDoneEmp.addEventListener("click", closeActiveCard);

    $activeOverlay.addEventListener("click", (e) => {
        if (e.target === $activeOverlay) closeActiveCard();
    });

    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && !$activeOverlay.classList.contains("hidden")) {
            closeActiveCard();
            return;
        }

        const mod = e.metaKey || e.ctrlKey;

        // Ctrl/Cmd+Enter — Done with employee
        if (mod && e.key === "Enter" && !$activeOverlay.classList.contains("hidden")) {
            e.preventDefault();
            closeActiveCard();
            return;
        }

    });

    // -----------------------------------------------------------------------
    // OT controls
    // -----------------------------------------------------------------------
    function setOtControlsEnabled(enabled) {
        $empSearch.disabled = !enabled;
        $otDate.disabled = !enabled;
        $otCategory.disabled = !enabled;
        $otHours.disabled = !enabled;
        $btnAddOt.disabled = !enabled;
        if ($otRangeText) $otRangeText.disabled = !enabled;
        if ($btnAddWeekBlock) $btnAddWeekBlock.disabled = !enabled;
        document.querySelectorAll('input[name="ot-week"]').forEach(r => { r.disabled = !enabled; });
        document.querySelectorAll(".ot-wb-hrs").forEach(el => { el.disabled = !enabled; });
    }

    function updateButtonStates() {
        const ready = state.employees.length > 0 && state.payPeriodEnd;
        $btnGenerateSlips.disabled = !ready;
        setOtControlsEnabled(!!ready);
        $btnGenerateOt.disabled = !(ready && state.otEntries.length > 0);
        $btnClearSession.disabled = state.otEntries.length === 0;
        if ($btnImportParse && $importCorrections && $importCorrections.files[0]) {
            $btnImportParse.disabled = !ready;
        }
    }

    // -----------------------------------------------------------------------
    // Add OT entry
    // -----------------------------------------------------------------------
    $btnAddOt.addEventListener("click", addOtEntry);
    $otHours.addEventListener("keydown", (e) => { if (e.key === "Enter") addOtEntry(); });
    if ($btnAddWeekBlock) $btnAddWeekBlock.addEventListener("click", addWeekBlockEntry);
    document.querySelectorAll(".ot-wb-hrs").forEach(el => {
        el.addEventListener("keydown", (e) => { if (e.key === "Enter") addWeekBlockEntry(); });
    });

    let cardStatusTimer = null;
    function showCardStatus(msg, type) {
        showStatus($activeCardStatus, msg, type);
        clearTimeout(cardStatusTimer);
        cardStatusTimer = setTimeout(() => hideStatus($activeCardStatus), 3000);
    }

    function addWeekBlockEntry() {
        const empNo = $empSelectedNo.value;
        const weekEl = document.querySelector('input[name="ot-week"]:checked');
        const week = weekEl ? parseInt(weekEl.value, 10) : 1;
        const rangeText = ($otRangeText && $otRangeText.value || "").trim();
        const ot10 = parseFloat($("ot-wb-ot10") && $("ot-wb-ot10").value) || 0;
        const ot15 = parseFloat($("ot-wb-ot15") && $("ot-wb-ot15").value) || 0;
        const cte10 = parseFloat($("ot-wb-cte10") && $("ot-wb-cte10").value) || 0;
        const cte15 = parseFloat($("ot-wb-cte15") && $("ot-wb-cte15").value) || 0;
        const totalH = ot10 + ot15 + cte10 + cte15;

        if (!empNo) return showCardStatus("Search and select an employee first.", "error");
        if (!rangeText) return showCardStatus("Enter a date range (e.g. 3/23–3/25).", "error");
        if (!totalH || totalH <= 0) return showCardStatus("Enter hours in at least one category.", "error");

        const emp = state.employees.find(e => e.emp_no === empNo);
        const round2 = (x) => Math.round(x * 100) / 100;

        state.otEntries.push({
            kind: "weekBlock",
            empNo, last: emp.last, first: emp.first,
            week,
            rangeText,
            ot10: round2(ot10), ot15: round2(ot15), cte10: round2(cte10), cte15: round2(cte15),
        });

        saveState();
        hideStatus($activeCardStatus);
        renderActiveEntries();
        renderCompletedCards();
        renderSummary();
        updateButtonStates();

        resetWeekBlockForm();
        if ($otRangeText) $otRangeText.focus();
    }

    function addOtEntry() {
        const empNo = $empSelectedNo.value;
        const date = $otDate.value;
        const category = $otCategory.value;
        const hours = parseFloat($otHours.value);

        if (!empNo) return showCardStatus("Search and select an employee first.", "error");
        if (!date) return showCardStatus("Select a date.", "error");
        if (!hours || hours <= 0) return showCardStatus("Enter hours greater than 0.", "error");

        const emp = state.employees.find(e => e.emp_no === empNo);

        const hrs = Math.round(hours * 100) / 100;
        state.otEntries.push({
            empNo, last: emp.last, first: emp.first,
            date, category, hours: hrs,
        });

        saveState();
        hideStatus($activeCardStatus);
        renderActiveEntries();
        renderCompletedCards();
        renderSummary();
        updateButtonStates();

        $otHours.value = "";
        $otDate.focus();
    }

    // -----------------------------------------------------------------------
    // Undo / remove
    // -----------------------------------------------------------------------
    function removeEntry(idx) {
        state.otEntries.splice(idx, 1);
        saveState();
        renderCompletedCards();
        renderSummary();
        updateButtonStates();
    }

    // -----------------------------------------------------------------------
    // Render
    // -----------------------------------------------------------------------
    function renderAll() {
        if (activeEmpNo) renderActiveEntries();
        renderCompletedCards();
        renderSummary();
    }

    function renderCompletedCards() {
        const groups = {};
        state.otEntries.forEach((e, idx) => {
            if (!groups[e.empNo]) groups[e.empNo] = { emp: e, entries: [], totalHrs: 0 };
            groups[e.empNo].entries.push({ ...e, _idx: idx });
            groups[e.empNo].totalHrs += entryHours(e);
        });

        const sorted = Object.values(groups).sort((a, b) =>
            `${a.emp.last}${a.emp.first}`.localeCompare(`${b.emp.last}${b.emp.first}`)
        );

        if (sorted.length === 0) {
            $otCardsWrap.classList.add("hidden");
            $otSummaryWrap.classList.add("hidden");
            return;
        }

        $otCardsWrap.classList.remove("hidden");
        $otSummaryWrap.classList.remove("hidden");
        $otEntryCount.textContent = uniqueEmpCount();

        $otCards.innerHTML = sorted.map(g => {
            const isActive = g.emp.empNo === activeEmpNo;
            return `<div class="completed-row ${isActive ? 'active-highlight' : ''}" data-empno="${g.emp.empNo}">
                <div class="completed-row-header">
                    <span class="completed-name">${g.emp.last}, ${g.emp.first}</span>
                    <span class="completed-meta">
                        #${g.emp.empNo} &middot; ${g.entries.length} entries &middot; ${Number(g.totalHrs).toFixed(2)} hrs
                    </span>
                    <button class="btn-edit-emp btn-small btn-ghost" data-empno="${g.emp.empNo}">Edit</button>
                </div>
            </div>`;
        }).join("");

        $otCards.querySelectorAll(".btn-edit-emp").forEach(btn => {
            btn.addEventListener("click", () => {
                const empNo = btn.dataset.empno;
                const emp = state.employees.find(e => e.emp_no === empNo);
                if (emp) selectEmployee(emp);
            });
        });
    }

    function renderSummary() {
        if (state.otEntries.length === 0) {
            $otSummaryWrap.classList.add("hidden");
            return;
        }
        $otSummaryWrap.classList.remove("hidden");

        const byEmp = {};
        state.otEntries.forEach(e => {
            if (!byEmp[e.empNo]) {
                byEmp[e.empNo] = {
                    last: e.last, first: e.first, empNo: e.empNo,
                    wk: {
                        1: { cats: {}, dates: new Set(), ranges: [] },
                        2: { cats: {}, dates: new Set(), ranges: [] },
                    },
                };
            }
            if (e.kind === "weekBlock") {
                const wn = e.week === 2 || e.week === "2" ? 2 : 1;
                const w = byEmp[e.empNo].wk[wn];
                CATS.forEach(c => {
                    const v = Number(e[c]) || 0;
                    if (v > 0) w.cats[c] = (w.cats[c] || 0) + v;
                });
                if (e.rangeText) w.ranges.push(String(e.rangeText).trim());
            } else {
                const week = dateToWeek(e.date, state.payPeriodEnd);
                if (week === "1" || week === "2") {
                    const w = byEmp[e.empNo].wk[week];
                    w.cats[e.category] = (w.cats[e.category] || 0) + e.hours;
                    w.dates.add(e.date);
                }
            }
        });

        const empList = Object.values(byEmp).sort((a, b) =>
            `${a.last}${a.first}`.localeCompare(`${b.last}${b.first}`)
        );

        function fmtDatesSet(dateSet) {
            if (dateSet.size === 0) return "";
            const sorted = [...dateSet].sort();
            return sorted.map(ds => {
                const d = new Date(ds + "T00:00:00");
                return fmtShort(d);
            }).join(", ");
        }

        let html = "";
        let grandTotal = 0;

        empList.forEach(emp => {
            let empTotal = 0;

            const wk1Dates = fmtDatesSet(emp.wk["1"].dates);
            const wk1Ranges = emp.wk["1"].ranges.length ? emp.wk["1"].ranges.join("; ") : "";
            let wk1Total = 0;
            const wk1Cells = CATS.map(c => {
                const v = emp.wk["1"].cats[c] || 0; wk1Total += v;
                return `<td>${v ? Number(v).toFixed(2) : ""}</td>`;
            }).join("");
            empTotal += wk1Total;

            const wk2Dates = fmtDatesSet(emp.wk["2"].dates);
            const wk2Ranges = emp.wk["2"].ranges.length ? emp.wk["2"].ranges.join("; ") : "";
            let wk2Total = 0;
            const wk2Cells = CATS.map(c => {
                const v = emp.wk["2"].cats[c] || 0; wk2Total += v;
                return `<td>${v ? Number(v).toFixed(2) : ""}</td>`;
            }).join("");
            empTotal += wk2Total;
            grandTotal += empTotal;

            const wk1LabelText = [wk1Dates, wk1Ranges].filter(Boolean).join("; ");
            const wk2LabelText = [wk2Dates, wk2Ranges].filter(Boolean).join("; ");
            const wk1Label = wk1LabelText ? `Wk 1: ${wk1LabelText}` : "Wk 1";
            const wk2Label = wk2LabelText ? `Wk 2: ${wk2LabelText}` : "Wk 2";

            html += `<tr>
                <td class="emp-name-cell" rowspan="2">${emp.last}, ${emp.first}<br><span style="font-size:0.78rem;color:var(--text-muted)">#${emp.empNo}</span></td>
                <td class="wk-label">${wk1Label}</td>${wk1Cells}
                <td>${wk1Total ? Number(wk1Total).toFixed(2) : ""}</td>
            </tr>
            <tr class="wk-row"><td class="wk-label">${wk2Label}</td>${wk2Cells}
                <td>${wk2Total ? Number(wk2Total).toFixed(2) : ""}</td>
            </tr>
            <tr><td colspan="7" class="emp-total-cell" style="text-align:right;padding-right:1rem;">
                Employee Total: <strong>${Number(empTotal).toFixed(2)}</strong>
            </td></tr>`;
        });

        html += `<tr class="grand-total">
            <td colspan="6" style="text-align:right;padding-right:1rem;">GRAND TOTAL</td>
            <td><strong>${Number(grandTotal).toFixed(2)}</strong></td>
        </tr>`;

        $otSummaryBody.innerHTML = html;
    }

    // -----------------------------------------------------------------------
    // Generate blank slips (Feature 1)
    // -----------------------------------------------------------------------
    $btnGenerateSlips.addEventListener("click", async () => {
        startElapsedTimer($slipsStatus, "Generating PDF binder...");
        $btnGenerateSlips.disabled = true;
        try {
            const ctrl1 = new AbortController();
            const t1 = setTimeout(() => ctrl1.abort(), 120000);
            const resp = await fetch("/api/generate-slips", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ employees: state.employees, payPeriodEnd: state.payPeriodEnd }),
                signal: ctrl1.signal,
            });
            clearTimeout(t1);
            if (!resp.ok) { const err = await resp.json(); throw new Error(err.error || "Server error"); }
            const blob = await resp.blob();
            downloadBlob(blob, resp.headers.get("Content-Disposition")?.split("filename=")[1] || "slips.pdf");
            stopElapsedTimer($slipsStatus);
            showStatus($slipsStatus, "PDF binder downloaded successfully.", "success");
        } catch (err) {
            stopElapsedTimer($slipsStatus);
            const msg = err.name === "AbortError" ? "Request timed out — try again, it may take a moment for large employee lists." : err.message;
            showStatus($slipsStatus, "Error: " + msg, "error");
        } finally { updateButtonStates(); }
    });

    // -----------------------------------------------------------------------
    // Generate OT slips + Excel (Feature 2)
    // -----------------------------------------------------------------------
    $btnGenerateOt.addEventListener("click", async () => {
        startElapsedTimer($otStatus, "Generating OT slips and Excel...");
        $btnGenerateOt.disabled = true;

        const otByEmp = buildOtPayloadFromEntries();

        try {
            const ctrl2 = new AbortController();
            const t2 = setTimeout(() => ctrl2.abort(), 120000);
            const resp = await fetch("/api/generate-overtime", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ employees: state.employees, payPeriodEnd: state.payPeriodEnd, otEntries: otByEmp }),
                signal: ctrl2.signal,
            });
            clearTimeout(t2);
            const data = await resp.json();
            if (data.error) throw new Error(data.error);

            const pdfBytes = Uint8Array.from(atob(data.pdf), c => c.charCodeAt(0));
            downloadBlob(new Blob([pdfBytes], { type: "application/pdf" }), data.pdfFilename);

            const xlBytes = Uint8Array.from(atob(data.excel), c => c.charCodeAt(0));
            downloadBlob(new Blob([xlBytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), data.excelFilename);

            stopElapsedTimer($otStatus);
            showStatus($otStatus, "OT slips PDF and Excel summary downloaded.", "success");
        } catch (err) {
            stopElapsedTimer($otStatus);
            const msg = err.name === "AbortError" ? "Request timed out — try again, it may take a moment." : err.message;
            showStatus($otStatus, "Error: " + msg, "error");
        } finally { updateButtonStates(); }
    });

    // -----------------------------------------------------------------------
    // Clear session
    // -----------------------------------------------------------------------
    $btnClearSession.addEventListener("click", () => {
        showConfirm(
            "Are you sure you want to clear all overtime data and start a new pay period?",
            "Yes, Clear Everything",
            () => {
                state.otEntries = [];
                state.payPeriodEnd = "";
                activeEmpNo = null;
                $ppEndDate.value = "";
                $empSearch.value = "";
                $empSelectedNo.value = "";
                $activeOverlay.classList.add("hidden");
                saveState();
                updatePayPeriodInfo();
                renderAll();
                updateButtonStates();
                hideStatus($otStatus);
                showStatus($otStatus, "Session cleared. Ready for a new pay period.", "success");
            }
        );
    });

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------
    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = filename || "download";
        document.body.appendChild(a); a.click(); a.remove();
        URL.revokeObjectURL(url);
    }
    function showStatus(el, msg, type) {
        el.innerHTML = msg;
        el.className = `status-msg ${type}`;
        el.classList.remove("hidden");
    }
    function hideStatus(el) { el.classList.add("hidden"); el._timer = null; }

    function startElapsedTimer(el, baseMsg) {
        const start = Date.now();
        el.innerHTML = `<div class="progress-bar"><div class="progress-fill"></div></div>${baseMsg} 0s elapsed`;
        el.className = "status-msg loading";
        el.classList.remove("hidden");
        el._timer = setInterval(() => {
            const sec = Math.round((Date.now() - start) / 1000);
            el.innerHTML = `<div class="progress-bar"><div class="progress-fill"></div></div>${baseMsg} ${sec}s elapsed`;
        }, 1000);
    }
    function stopElapsedTimer(el) {
        if (el._timer) { clearInterval(el._timer); el._timer = null; }
    }

    // -----------------------------------------------------------------------
    // Init
    // -----------------------------------------------------------------------
    initTheme();
    loadState();
    if (state.payPeriodEnd) {
        $ppEndDate.value = state.payPeriodEnd;
        updatePayPeriodInfo();
    }
    renderEmployeeState();
    renderCompletedCards();
    renderSummary();
    updateButtonStates();
})();
