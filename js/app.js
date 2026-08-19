"use strict";

/* ============================================================
 *  Normalisation Excel - logique applicative
 * ============================================================ */

let sourceWorkbook = null;
let processedWorkbook = null;
let sourceFileName = "fichier";

const fileInput = document.getElementById("excelFile");
const processButton = document.getElementById("processButton");
const downloadButton = document.getElementById("downloadButton");
const sheetSelector = document.getElementById("sheetSelector");

fileInput.addEventListener("change", loadWorkbook);
processButton.addEventListener("click", processWorkbook);
downloadButton.addEventListener("click", downloadWorkbook);
sheetSelector.addEventListener("change", updatePreview);

/* Règles métier par défaut. */
addRule("RH", "RHD", true, false, true);
addRule("LH", "LHD", true, false, true);
addRule("left hand drive", "LHD", false, false, true);
addRule("right hand drive", "RHD", false, false, true);

/* ------------------------------------------------------------
 *  Chargement du fichier
 * ---------------------------------------------------------- */
async function loadWorkbook(event) {
    const file = event.target.files[0];
    if (!file) { return; }

    try {
        setStatus("Chargement du fichier...", "info");
        sourceFileName = file.name.replace(/\.[^.]+$/, "");

        const buffer = await file.arrayBuffer();
        sourceWorkbook = XLSX.read(buffer, {
            type: "array",
            cellDates: true,
            cellFormula: true
        });

        processedWorkbook = null;
        downloadButton.disabled = true;

        buildScopePanel();
        document.getElementById("scopeCard").style.display = "block";
        document.getElementById("previewCard").style.display = "none";

        setStatus(`${sourceWorkbook.SheetNames.length} feuille(s) chargée(s).`, "success");
    } catch (error) {
        console.error(error);
        setStatus("Impossible de lire le fichier.", "error");
    }
}

/* ------------------------------------------------------------
 *  Construction du panneau de sélection des colonnes
 * ---------------------------------------------------------- */
function buildScopePanel() {
    const panel = document.getElementById("scopePanel");
    panel.innerHTML = "";

    sourceWorkbook.SheetNames.forEach((sheetName, sheetIndex) => {
        const sheet = sourceWorkbook.Sheets[sheetName];

        const details = document.createElement("details");
        details.dataset.sheetName = sheetName;
        if (sheetIndex === 0) { details.open = true; }

        const summary = document.createElement("summary");
        summary.textContent = sheetName;
        details.appendChild(summary);

        const controls = document.createElement("div");
        controls.className = "sheet-controls";

        const sheetLabel = document.createElement("label");
        sheetLabel.className = "inline-checkbox";

        const sheetCheckbox = document.createElement("input");
        sheetCheckbox.type = "checkbox";
        sheetCheckbox.checked = true;
        sheetCheckbox.className = "sheet-enabled";

        sheetLabel.appendChild(sheetCheckbox);
        sheetLabel.append(" Feuille active");
        controls.appendChild(sheetLabel);

        const allButton = document.createElement("button");
        allButton.type = "button";
        allButton.className = "secondary-button";
        allButton.textContent = "Toutes les colonnes";
        allButton.onclick = () => selectSheetColumns(details, true);

        const noneButton = document.createElement("button");
        noneButton.type = "button";
        noneButton.className = "secondary-button";
        noneButton.textContent = "Aucune colonne";
        noneButton.onclick = () => selectSheetColumns(details, false);

        controls.appendChild(allButton);
        controls.appendChild(noneButton);
        details.appendChild(controls);

        if (!sheet["!ref"]) {
            const emptyMessage = document.createElement("p");
            emptyMessage.className = "hint";
            emptyMessage.style.padding = "0 12px 10px";
            emptyMessage.textContent = "Feuille vide.";
            details.appendChild(emptyMessage);
            panel.appendChild(details);
            return;
        }

        const range = XLSX.utils.decode_range(sheet["!ref"]);

        const wrapper = document.createElement("div");
        wrapper.className = "table-wrapper";

        const table = document.createElement("table");
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Modifier</th>
                    <th>Colonne</th>
                    <th>Nom détecté</th>
                    <th>Type détecté</th>
                    <th>Type à appliquer</th>
                    <th>Majuscules / minuscules</th>
                </tr>
            </thead>
        `;

        const tbody = document.createElement("tbody");
        for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex++) {
            tbody.appendChild(createColumnRow(sheet, range, columnIndex));
        }

        table.appendChild(tbody);
        wrapper.appendChild(table);
        details.appendChild(wrapper);
        panel.appendChild(details);
    });
}

function createColumnRow(sheet, range, columnIndex) {
    const row = document.createElement("tr");
    row.dataset.columnIndex = columnIndex;

    const firstCellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: columnIndex });
    const firstCell = sheet[firstCellAddress];
    const headerValue = firstCell ? String(firstCell.w ?? firstCell.v ?? "") : "";

    const detectedType = detectColumnType(sheet, columnIndex, range);

    /* Colonne à modifier (décochée par défaut). */
    const enabledCell = document.createElement("td");
    const enabledCheckbox = document.createElement("input");
    enabledCheckbox.type = "checkbox";
    enabledCheckbox.checked = false;
    enabledCheckbox.className = "column-enabled";
    enabledCell.appendChild(enabledCheckbox);

    /* Lettre de la colonne. */
    const letterCell = document.createElement("td");
    letterCell.textContent = XLSX.utils.encode_col(columnIndex);

    /* Nom détecté. */
    const nameCell = document.createElement("td");
    nameCell.className = "column-name";
    nameCell.textContent = headerValue || "(sans en-tête)";
    nameCell.title = headerValue || "(sans en-tête)";

    /* Type détecté. */
    const detectedTypeCell = document.createElement("td");
    detectedTypeCell.textContent = detectedType.label;

    /* Type à appliquer. */
    const targetTypeCell = document.createElement("td");
    const typeSelect = document.createElement("select");
    typeSelect.className = "column-type";
    typeSelect.innerHTML = `
        <option value="auto">Automatique</option>
        <option value="text">Texte</option>
        <option value="number">Nombre</option>
        <option value="date">Date</option>
        <option value="boolean">Booléen</option>
    `;
    typeSelect.value = "auto";
    targetTypeCell.appendChild(typeSelect);

    /* Casse. */
    const caseCell = document.createElement("td");
    const caseSelect = document.createElement("select");
    caseSelect.className = "column-case";
    caseSelect.innerHTML = `
        <option value="unchanged">Conserver la casse</option>
        <option value="lowercase">Tout en minuscules</option>
        <option value="uppercase">Tout en MAJUSCULES</option>
    `;
    caseCell.appendChild(caseSelect);

    row.appendChild(enabledCell);
    row.appendChild(letterCell);
    row.appendChild(nameCell);
    row.appendChild(detectedTypeCell);
    row.appendChild(targetTypeCell);
    row.appendChild(caseCell);
    return row;
}

function detectColumnType(sheet, columnIndex, range) {
    const counts = { text: 0, number: 0, date: 0, boolean: 0 };
    let inspected = 0;

    for (let row = range.s.r + 1; row <= range.e.r && inspected < 100; row++) {
        const address = XLSX.utils.encode_cell({ r: row, c: columnIndex });
        const cell = sheet[address];

        if (!cell || cell.v === "" || cell.v === null || cell.v === undefined) {
            continue;
        }
        inspected++;

        if (cell.t === "d" || cell.v instanceof Date) {
            counts.date++;
        } else if (cell.t === "n") {
            counts.number++;
        } else if (cell.t === "b") {
            counts.boolean++;
        } else {
            counts.text++;
        }
    }

    if (inspected === 0) {
        return { value: "auto", label: "Indéterminé" };
    }

    const dominantType = Object.entries(counts)
        .sort((first, second) => second[1] - first[1])[0][0];

    const labels = { text: "Texte", number: "Nombre", date: "Date", boolean: "Booléen" };
    return { value: dominantType, label: labels[dominantType] };
}

/* ------------------------------------------------------------
 *  Sélections globales
 * ---------------------------------------------------------- */
function selectAllSheets(selected) {
    document.querySelectorAll(".sheet-enabled").forEach(checkbox => {
        checkbox.checked = selected;
    });
}

function selectSheetColumns(sheetSection, selected) {
    sheetSection.querySelectorAll(".column-enabled").forEach(checkbox => {
        checkbox.checked = selected;
    });
}

function getScopeConfiguration() {
    const configuration = {};

    document.querySelectorAll("#scopePanel details").forEach(details => {
        const sheetName = details.dataset.sheetName;
        const sheetCheckbox = details.querySelector(".sheet-enabled");

        configuration[sheetName] = {
            enabled: sheetCheckbox ? sheetCheckbox.checked : false,
            columns: {}
        };

        details.querySelectorAll("tbody tr").forEach(row => {
            const columnIndex = Number(row.dataset.columnIndex);
            configuration[sheetName].columns[columnIndex] = {
                enabled: row.querySelector(".column-enabled").checked,
                type: row.querySelector(".column-type").value,
                caseMode: row.querySelector(".column-case").value
            };
        });
    });

    return configuration;
}

/* ------------------------------------------------------------
 *  Traitement du classeur
 * ---------------------------------------------------------- */
function processWorkbook() {
    if (!sourceWorkbook) {
        setStatus("Veuillez sélectionner un fichier.", "error");
        return;
    }

    try {
        const scope = getScopeConfiguration();
        const selectedColumns = countSelectedColumns(scope);

        if (selectedColumns === 0) {
            setStatus("Sélectionnez au moins une colonne à modifier.", "error");
            return;
        }

        const settings = getSettings();
        const rules = getCustomRules();

        processedWorkbook = XLSX.utils.book_new();
        let modifiedCells = 0;
        let conversionErrors = 0;

        sourceWorkbook.SheetNames.forEach(sheetName => {
            const sourceSheet = sourceWorkbook.Sheets[sheetName];
            const targetSheet = cloneWorksheet(sourceSheet);
            const sheetConfiguration = scope[sheetName];

            if (sheetConfiguration && sheetConfiguration.enabled && targetSheet["!ref"]) {
                const range = XLSX.utils.decode_range(targetSheet["!ref"]);

                for (let row = range.s.r; row <= range.e.r; row++) {
                    if (settings.ignoreHeaders && row === range.s.r) {
                        continue;
                    }

                    for (let column = range.s.c; column <= range.e.c; column++) {
                        const columnConfiguration = sheetConfiguration.columns[column];
                        if (!columnConfiguration || !columnConfiguration.enabled) {
                            continue;
                        }

                        const address = XLSX.utils.encode_cell({ r: row, c: column });
                        const cell = targetSheet[address];
                        if (!cell || cell.f) {
                            continue;
                        }

                        const originalType = cell.t;
                        const originalValue = cloneCellValue(cell.v);

                        const result = transformCell(cell, columnConfiguration, settings, rules);
                        if (!result.success) {
                            conversionErrors++;
                            continue;
                        }

                        if (originalType !== cell.t || !valuesAreEqual(originalValue, cell.v)) {
                            modifiedCells++;
                            delete cell.w;
                        }
                    }
                }
            }

            XLSX.utils.book_append_sheet(processedWorkbook, targetSheet, sheetName);
        });

        populateSheetSelector();
        updatePreview();
        downloadButton.disabled = false;
        document.getElementById("previewCard").style.display = "block";

        let message = `${selectedColumns} colonne(s) sélectionnée(s). ` +
            `${modifiedCells} cellule(s) modifiée(s).`;
        if (conversionErrors > 0) {
            message += ` ${conversionErrors} valeur(s) non convertie(s).`;
        }

        setStatus(message, conversionErrors > 0 ? "info" : "success");
    } catch (error) {
        console.error(error);
        setStatus("Erreur pendant le traitement.", "error");
    }
}

function countSelectedColumns(scope) {
    let count = 0;
    Object.values(scope).forEach(sheetConfiguration => {
        if (!sheetConfiguration.enabled) { return; }
        Object.values(sheetConfiguration.columns).forEach(columnConfiguration => {
            if (columnConfiguration.enabled) { count++; }
        });
    });
    return count;
}

/* ------------------------------------------------------------
 *  Transformation d'une cellule
 * ---------------------------------------------------------- */
function transformCell(cell, columnConfiguration, settings, rules) {
    if (cell.v === null || cell.v === undefined || cell.v === "") {
        return { success: true };
    }

    const targetType = columnConfiguration.type;
    const caseMode = columnConfiguration.caseMode;

    if (targetType === "auto") {
        if (cell.t === "s" || cell.t === "str") {
            cell.v = normalizeText(String(cell.v), caseMode, settings, rules);
            cell.t = "s";
        }
        return { success: true };
    }

    if (targetType === "text") {
        let textValue;
        if (cell.v instanceof Date) {
            textValue = formatDate(cell.v);
        } else if (cell.t === "b") {
            textValue = cell.v ? "Oui" : "Non";
        } else {
            textValue = String(cell.v);
        }
        cell.v = normalizeText(textValue, caseMode, settings, rules);
        cell.t = "s";
        return { success: true };
    }

    if (targetType === "number") {
        const numberValue = parseNumber(cell.v);
        if (numberValue === null) { return { success: false }; }
        cell.v = numberValue;
        cell.t = "n";
        return { success: true };
    }

    if (targetType === "date") {
        const dateValue = parseDate(cell.v, cell.t);
        if (!dateValue) { return { success: false }; }
        cell.v = dateValue;
        cell.t = "d";
        cell.z = "dd/mm/yyyy";
        return { success: true };
    }

    if (targetType === "boolean") {
        const booleanValue = parseBoolean(cell.v);
        if (booleanValue === null) { return { success: false }; }
        cell.v = booleanValue;
        cell.t = "b";
        return { success: true };
    }

    return { success: true };
}

function normalizeText(value, caseMode, settings, rules) {
    let result = String(value);

    if (settings.spaceMode === "trim") {
        result = result.trim();
    }
    if (settings.spaceMode === "underscore") {
        result = result.trim().replace(/\s+/g, "_");
    }

    /* Casse appliquée uniquement à la colonne sélectionnée. */
    if (caseMode === "lowercase") {
        result = result.toLocaleLowerCase("fr-FR");
    }
    if (caseMode === "uppercase") {
        result = result.toLocaleUpperCase("fr-FR");
    }

    /* Règles les plus longues d'abord (évite les collisions RH / RHD). */
    const orderedRules = [...rules].sort(
        (firstRule, secondRule) => secondRule.search.length - firstRule.search.length
    );
    orderedRules.forEach(rule => {
        result = executeReplacementRule(result, rule);
    });

    settings.protectedTerms.forEach(term => {
        const cleanTerm = term.trim();
        if (!cleanTerm) { return; }
        result = executeReplacementRule(result, {
            enabled: true,
            search: cleanTerm,
            replacement: cleanTerm.toLocaleUpperCase("fr-FR"),
            wholeWord: true,
            caseSensitive: false
        });
    });

    if (settings.capitalizeFirst) {
        result = capitalizeFirstLetter(result);
    }

    return result;
}

function executeReplacementRule(input, rule) {
    if (!rule.enabled || !rule.search.trim()) {
        return input;
    }

    const escapedSearch = escapeRegExp(rule.search.trim());
    const flags = rule.caseSensitive ? "gu" : "giu";

    if (rule.wholeWord) {
        const pattern = `(^|[^\\p{L}\\p{N}])(${escapedSearch})(?=$|[^\\p{L}\\p{N}])`;
        const regex = new RegExp(pattern, flags);
        return input.replace(regex, (completeMatch, prefix) => prefix + rule.replacement);
    }

    const regex = new RegExp(escapedSearch, flags);
    return input.replace(regex, () => rule.replacement);
}

/* ------------------------------------------------------------
 *  Options et règles
 * ---------------------------------------------------------- */
function getSettings() {
    const protectedTerms = document.getElementById("protectedTerms").value
        .split(",")
        .map(value => value.trim())
        .filter(Boolean);

    return {
        spaceMode: document.getElementById("spaceMode").value,
        capitalizeFirst: document.getElementById("capitalizeFirst").checked,
        ignoreHeaders: document.getElementById("ignoreHeaders").checked,
        protectedTerms
    };
}

function addRule(searchValue = "", replacementValue = "", wholeWord = true, caseSensitive = false, enabled = true) {
    const row = document.createElement("div");
    row.className = "rule-row";
    row.innerHTML = `
        <div class="rule-checkbox">
            <input class="rule-enabled" type="checkbox" title="Activer la règle">
        </div>
        <input class="rule-search" type="text" placeholder="Texte recherché">
        <input class="rule-replacement" type="text" placeholder="Remplacement">
        <div class="rule-checkbox">
            <input class="rule-whole-word" type="checkbox" title="Mot entier">
        </div>
        <div class="rule-checkbox">
            <input class="rule-case-sensitive" type="checkbox" title="Respecter la casse">
        </div>
        <button type="button" class="danger-button">Supprimer</button>
    `;

    row.querySelector(".rule-enabled").checked = enabled;
    row.querySelector(".rule-search").value = searchValue;
    row.querySelector(".rule-replacement").value = replacementValue;
    row.querySelector(".rule-whole-word").checked = wholeWord;
    row.querySelector(".rule-case-sensitive").checked = caseSensitive;
    row.querySelector(".danger-button").addEventListener("click", () => row.remove());

    document.getElementById("customRules").appendChild(row);
}

function getCustomRules() {
    return Array.from(document.querySelectorAll(".rule-row"))
        .map(row => ({
            enabled: row.querySelector(".rule-enabled").checked,
            search: row.querySelector(".rule-search").value,
            replacement: row.querySelector(".rule-replacement").value,
            wholeWord: row.querySelector(".rule-whole-word").checked,
            caseSensitive: row.querySelector(".rule-case-sensitive").checked
        }))
        .filter(rule => rule.search.trim() !== "");
}

/* ------------------------------------------------------------
 *  Convertisseurs de valeurs
 * ---------------------------------------------------------- */
function parseNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
        return value;
    }

    let normalized = String(value).trim().replace(/\s/g, "").replace(/[€$£%]/g, "");

    if (normalized.includes(",") && normalized.includes(".")) {
        if (normalized.lastIndexOf(",") > normalized.lastIndexOf(".")) {
            normalized = normalized.replace(/\./g, "").replace(",", ".");
        } else {
            normalized = normalized.replace(/,/g, "");
        }
    } else {
        normalized = normalized.replace(",", ".");
    }

    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
}

function parseDate(value, cellType) {
    if (value instanceof Date && !isNaN(value)) {
        return value;
    }

    if (cellType === "n" && typeof value === "number") {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (parsed) {
            return new Date(
                parsed.y,
                parsed.m - 1,
                parsed.d,
                parsed.H || 0,
                parsed.M || 0,
                Math.floor(parsed.S || 0)
            );
        }
    }

    const text = String(value).trim();

    let match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
    if (match) {
        return createValidatedDate(Number(match[3]), Number(match[2]), Number(match[1]));
    }

    match = text.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
    if (match) {
        return createValidatedDate(Number(match[1]), Number(match[2]), Number(match[3]));
    }

    return null;
}

function createValidatedDate(year, month, day) {
    const date = new Date(year, month - 1, day);
    if (
        date.getFullYear() !== year ||
        date.getMonth() !== month - 1 ||
        date.getDate() !== day
    ) {
        return null;
    }
    return date;
}

function parseBoolean(value) {
    if (typeof value === "boolean") { return value; }
    if (value === 1) { return true; }
    if (value === 0) { return false; }

    const normalized = String(value).trim().toLocaleLowerCase("fr-FR");
    const trueValues = ["oui", "yes", "true", "vrai", "1", "x"];
    const falseValues = ["non", "no", "false", "faux", "0"];

    if (trueValues.includes(normalized)) { return true; }
    if (falseValues.includes(normalized)) { return false; }
    return null;
}

function capitalizeFirstLetter(value) {
    const characters = Array.from(value);
    for (let index = 0; index < characters.length; index++) {
        if (/\p{L}/u.test(characters[index])) {
            characters[index] = characters[index].toLocaleUpperCase("fr-FR");
            break;
        }
    }
    return characters.join("");
}

/* ------------------------------------------------------------
 *  Prévisualisation
 * ---------------------------------------------------------- */
function populateSheetSelector() {
    sheetSelector.innerHTML = "";
    processedWorkbook.SheetNames.forEach(sheetName => {
        const option = document.createElement("option");
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelector.appendChild(option);
    });
}

function updatePreview() {
    if (!processedWorkbook) { return; }

    const sheetName = sheetSelector.value || processedWorkbook.SheetNames[0];
    const sheet = processedWorkbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false
    });

    renderPreview(rows.slice(0, 100));
    document.getElementById("previewCounter").textContent =
        `${rows.length} ligne(s), aperçu limité à 100 lignes.`;
}

function renderPreview(rows) {
    const container = document.getElementById("previewContent");

    if (!rows.length) {
        container.innerHTML = "<p style='padding:12px;'>Feuille vide.</p>";
        return;
    }

    const maxColumns = Math.max(...rows.map(row => row.length), 1);

    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    for (let column = 0; column < maxColumns; column++) {
        const th = document.createElement("th");
        th.textContent = XLSX.utils.encode_col(column);
        headerRow.appendChild(th);
    }

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.forEach(row => {
        const tr = document.createElement("tr");
        for (let column = 0; column < maxColumns; column++) {
            const td = document.createElement("td");
            const value = row[column] ?? "";
            td.textContent = value;
            td.title = String(value);
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    container.innerHTML = "";
    container.appendChild(table);
}

/* ------------------------------------------------------------
 *  Export
 * ---------------------------------------------------------- */
function downloadWorkbook() {
    if (!processedWorkbook) {
        setStatus("Aucun fichier traité.", "error");
        return;
    }

    const outputName = `${sourceFileName}_normalise.xlsx`;
    XLSX.writeFile(processedWorkbook, outputName, { compression: true });
    setStatus(`Le fichier ${outputName} a été généré.`, "success");
}

/* ------------------------------------------------------------
 *  Utilitaires
 * ---------------------------------------------------------- */
function cloneWorksheet(sourceSheet) {
    const clone = {};

    Object.keys(sourceSheet).forEach(key => {
        const value = sourceSheet[key];

        if (value instanceof Date) {
            clone[key] = new Date(value.getTime());
            return;
        }
        if (Array.isArray(value)) {
            clone[key] = value.map(item =>
                (item && typeof item === "object") ? { ...item } : item
            );
            return;
        }
        if (value && typeof value === "object") {
            clone[key] = { ...value };
            if (value.v instanceof Date) {
                clone[key].v = new Date(value.v.getTime());
            }
            return;
        }
        clone[key] = value;
    });

    return clone;
}

function cloneCellValue(value) {
    if (value instanceof Date) {
        return new Date(value.getTime());
    }
    return value;
}

function valuesAreEqual(first, second) {
    if (first instanceof Date && second instanceof Date) {
        return first.getTime() === second.getTime();
    }
    return first === second;
}

function formatDate(date) {
    return new Intl.DateTimeFormat("fr-FR").format(date);
}

function escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function setStatus(message, type) {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = `status ${type}`;
}
