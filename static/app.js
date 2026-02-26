(() => {
    "use strict";

    const STORAGE_KEY = "montebello_ot_state";
    const THEME_KEY = "app_theme";
    const CAT_LABELS = { ot10: "OT 1.0", ot15: "OT 1.5", cte10: "CTE 1.0", cte15: "CTE 1.5" };
    const CATS = ["ot10", "ot15", "cte10", "cte15"];

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
        renderActiveEntries();
        hideStatus($activeCardStatus);
        $activeOverlay.classList.remove("hidden");

        setTimeout(() => $otDate.focus(), 50);
    }

    function closeActiveCard() {
        activeEmpNo = null;
        $empSelectedNo.value = "";
        $empSearch.value = "";
        $activeOverlay.classList.add("hidden");
        $otDate.value = "";
        $otHours.value = "";
        hideStatus($activeCardStatus);
        renderCompletedCards();
        $empSearch.focus();
    }

    function renderActiveEntries() {
        const entries = state.otEntries
            .map((e, idx) => ({ ...e, _idx: idx }))
            .filter(e => e.empNo === activeEmpNo)
            .sort((a, b) => a.date.localeCompare(b.date));

        if (entries.length === 0) {
            $activeEntries.innerHTML = '<p class="active-empty">No entries yet. Add overtime below.</p>';
            return;
        }

        const totalHrs = entries.reduce((s, e) => s + e.hours, 0);

        $activeEntries.innerHTML = entries.map(e => {
            const d = new Date(e.date + "T00:00:00");
            const wk = dateToWeek(e.date, state.payPeriodEnd);
            return `<div class="active-entry-row">
                <span class="row-wk">Wk ${wk}</span>
                <span class="row-date">${fmtShort(d)}</span>
                <span class="row-cat">${CAT_LABELS[e.category]}</span>
                <span class="row-hrs">${e.hours.toFixed(1)} hrs</span>
                <button class="btn-remove" data-idx="${e._idx}" title="Remove">&times;</button>
            </div>`;
        }).join("") + `<div class="active-total">Total: ${totalHrs.toFixed(1)} hrs</div>`;

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
    }

    function updateButtonStates() {
        const ready = state.employees.length > 0 && state.payPeriodEnd;
        $btnGenerateSlips.disabled = !ready;
        setOtControlsEnabled(!!ready);
        $btnGenerateOt.disabled = !(ready && state.otEntries.length > 0);
        $btnClearSession.disabled = state.otEntries.length === 0;
    }

    // -----------------------------------------------------------------------
    // Add OT entry
    // -----------------------------------------------------------------------
    $btnAddOt.addEventListener("click", addOtEntry);
    $otHours.addEventListener("keydown", (e) => { if (e.key === "Enter") addOtEntry(); });

    let cardStatusTimer = null;
    function showCardStatus(msg, type) {
        showStatus($activeCardStatus, msg, type);
        clearTimeout(cardStatusTimer);
        cardStatusTimer = setTimeout(() => hideStatus($activeCardStatus), 3000);
    }

    function addOtEntry() {
        const empNo = $empSelectedNo.value;
        const date = $otDate.value;
        const category = $otCategory.value;
        const hours = parseFloat($otHours.value);

        if (!empNo) return showCardStatus("Search and select an employee first.", "error");
        if (!date) return showCardStatus("Select a date.", "error");
        if (!hours || hours <= 0) return showCardStatus("Enter hours greater than 0.", "error");

        const week = dateToWeek(date, state.payPeriodEnd);
        if (week === "?") return showCardStatus("Date is outside the selected pay period.", "error");

        const emp = state.employees.find(e => e.emp_no === empNo);

        state.otEntries.push({
            empNo, last: emp.last, first: emp.first,
            date, category, hours,
        });

        saveState();
        hideStatus($activeCardStatus);
        renderActiveEntries();
        renderCompletedCards();
        renderSummary();
        updateButtonStates();

        $otHours.value = "";
        $otDate.value = "";
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
            groups[e.empNo].totalHrs += e.hours;
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
                        #${g.emp.empNo} &middot; ${g.entries.length} entries &middot; ${g.totalHrs.toFixed(1)} hrs
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
                byEmp[e.empNo] = { last: e.last, first: e.first, empNo: e.empNo,
                    wk: { 1: { cats: {}, dates: new Set() }, 2: { cats: {}, dates: new Set() } } };
            }
            const week = dateToWeek(e.date, state.payPeriodEnd);
            if (week === "1" || week === "2") {
                const w = byEmp[e.empNo].wk[week];
                w.cats[e.category] = (w.cats[e.category] || 0) + e.hours;
                w.dates.add(e.date);
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
            let wk1Total = 0;
            const wk1Cells = CATS.map(c => {
                const v = emp.wk["1"].cats[c] || 0; wk1Total += v;
                return `<td>${v ? v.toFixed(1) : ""}</td>`;
            }).join("");
            empTotal += wk1Total;

            const wk2Dates = fmtDatesSet(emp.wk["2"].dates);
            let wk2Total = 0;
            const wk2Cells = CATS.map(c => {
                const v = emp.wk["2"].cats[c] || 0; wk2Total += v;
                return `<td>${v ? v.toFixed(1) : ""}</td>`;
            }).join("");
            empTotal += wk2Total;
            grandTotal += empTotal;

            const wk1Label = wk1Dates ? `Wk 1: ${wk1Dates}` : "Wk 1";
            const wk2Label = wk2Dates ? `Wk 2: ${wk2Dates}` : "Wk 2";

            html += `<tr>
                <td class="emp-name-cell" rowspan="2">${emp.last}, ${emp.first}<br><span style="font-size:0.78rem;color:var(--text-muted)">#${emp.empNo}</span></td>
                <td class="wk-label">${wk1Label}</td>${wk1Cells}
                <td>${wk1Total ? wk1Total.toFixed(1) : ""}</td>
            </tr>
            <tr class="wk-row"><td class="wk-label">${wk2Label}</td>${wk2Cells}
                <td>${wk2Total ? wk2Total.toFixed(1) : ""}</td>
            </tr>
            <tr><td colspan="7" class="emp-total-cell" style="text-align:right;padding-right:1rem;">
                Employee Total: <strong>${empTotal.toFixed(1)}</strong>
            </td></tr>`;
        });

        html += `<tr class="grand-total">
            <td colspan="6" style="text-align:right;padding-right:1rem;">GRAND TOTAL</td>
            <td><strong>${grandTotal.toFixed(1)}</strong></td>
        </tr>`;

        $otSummaryBody.innerHTML = html;
    }

    // -----------------------------------------------------------------------
    // Generate blank slips (Feature 1)
    // -----------------------------------------------------------------------
    $btnGenerateSlips.addEventListener("click", async () => {
        showStatus($slipsStatus, "Generating PDF binder...", "loading");
        $btnGenerateSlips.disabled = true;
        try {
            const resp = await fetch("/api/generate-slips", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ employees: state.employees, payPeriodEnd: state.payPeriodEnd }),
            });
            if (!resp.ok) { const err = await resp.json(); throw new Error(err.error || "Server error"); }
            const blob = await resp.blob();
            downloadBlob(blob, resp.headers.get("Content-Disposition")?.split("filename=")[1] || "slips.pdf");
            showStatus($slipsStatus, "PDF binder downloaded successfully.", "success");
        } catch (err) {
            showStatus($slipsStatus, "Error: " + err.message, "error");
        } finally { updateButtonStates(); }
    });

    // -----------------------------------------------------------------------
    // Generate OT slips + Excel (Feature 2)
    // -----------------------------------------------------------------------
    $btnGenerateOt.addEventListener("click", async () => {
        showStatus($otStatus, "Generating OT slips and Excel...", "loading");
        $btnGenerateOt.disabled = true;

        const otByEmp = {};
        state.otEntries.forEach(e => {
            if (!otByEmp[e.empNo]) otByEmp[e.empNo] = { entries: [] };
            otByEmp[e.empNo].entries.push({ date: e.date, category: e.category, hours: e.hours });
        });

        try {
            const resp = await fetch("/api/generate-overtime", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ employees: state.employees, payPeriodEnd: state.payPeriodEnd, otEntries: otByEmp }),
            });
            const data = await resp.json();
            if (data.error) throw new Error(data.error);

            const pdfBytes = Uint8Array.from(atob(data.pdf), c => c.charCodeAt(0));
            downloadBlob(new Blob([pdfBytes], { type: "application/pdf" }), data.pdfFilename);

            const xlBytes = Uint8Array.from(atob(data.excel), c => c.charCodeAt(0));
            downloadBlob(new Blob([xlBytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), data.excelFilename);

            showStatus($otStatus, "OT slips PDF and Excel summary downloaded.", "success");
        } catch (err) {
            showStatus($otStatus, "Error: " + err.message, "error");
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
        el.textContent = msg;
        el.className = `status-msg ${type}`;
        el.classList.remove("hidden");
    }
    function hideStatus(el) { el.classList.add("hidden"); }

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
