/* app.js — Slide Creator main application logic */

// ── State ──────────────────────────────────────────────────────
const state = {
    designSystem: null,       // "slick" | "colorful" | "bold" | "editorial" | "noir"
    deckTitle: "",
    deckAuthor: "",
    sections: [],             // [{id, name, color}]
    slides: [],               // [{id, type, sectionId, data}]
    currentScreen: "setup",
    editingSlideId: null,
    editingIsNew: false,      // true = came from gallery (show "Add to Deck"), false = editing existing
};

function genId() {
    return "id_" + Date.now() + "_" + Math.random().toString(36).substr(2, 6);
}

// ── Navigation ─────────────────────────────────────────────────
const SCREENS = ["setup", "sections", "gallery", "editor", "filmstrip"];
const SCREEN_LABELS = { setup: "Design", sections: "Sections", gallery: "Slides", editor: "Editor", filmstrip: "Deck" };

function navigateTo(screen) {
    document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
    const el = document.getElementById("screen-" + screen);
    if (el) el.classList.add("active");
    state.currentScreen = screen;
    updateBreadcrumb();
    if (screen === "setup") renderSetup();
    if (screen === "sections") renderSections();
    if (screen === "gallery") renderGallery();
    if (screen === "editor") renderEditor();
    if (screen === "filmstrip") renderFilmstrip();
}

function updateBreadcrumb() {
    const bc = document.getElementById("breadcrumb");
    const crumbs = [];
    for (const s of SCREENS) {
        if (s === "editor") continue; // editor isn't in breadcrumb
        const active = s === state.currentScreen || (state.currentScreen === "editor" && s === "gallery");
        const cls = active ? "active" : "";
        crumbs.push(`<span class="${cls}" data-screen="${s}">${SCREEN_LABELS[s]}</span>`);
    }
    bc.innerHTML = crumbs.join('<span class="sep">/</span>');
    bc.querySelectorAll("span[data-screen]").forEach(el => {
        el.addEventListener("click", () => {
            const target = el.dataset.screen;
            if (target === "gallery" || target === "filmstrip") {
                if (!state.designSystem) return;
                if (state.sections.length === 0 && target !== "setup") return;
            }
            navigateTo(target);
        });
    });
}

// ── Toast ──────────────────────────────────────────────────────
function showToast(msg) {
    const t = document.getElementById("toast");
    t.textContent = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 2500);
}

// ── Setup Screen ───────────────────────────────────────────────
function renderSetup() {
    const picker = document.getElementById("designPicker");
    const systems = [
        { id: "slick", label: "Slick Minimal", desc: "White background, green accent bars, serif headlines." },
        { id: "colorful", label: "Colorful", desc: "Colored header bars, multi-color cards, sans-serif." },
        { id: "bold", label: "Bold", desc: "Dark titles, warm gray content, bold geometric accents." },
        { id: "editorial", label: "Editorial", desc: "Serif headlines, warm paper background, refined and restrained." },
        { id: "noir", label: "Noir", desc: "Full dark mode, vivid electric accents, commanding presence." },
        { id: "editorial_v2", label: "Editorial v2", desc: "Purpose-driven layouts, progressive reveals, magazine design language." },
    ];
    picker.innerHTML = systems.map(ds => {
        const sel = state.designSystem === ds.id ? " selected" : "";
        const previewData = { title: "Sample Presentation", subtitle: "A preview of this design system", author: "Your Name", date: "2025" };
        const previewHtml = typeof PREVIEW_RENDERERS !== "undefined" && PREVIEW_RENDERERS.title
            ? PREVIEW_RENDERERS.title(previewData, ds.id, "green") : '<div style="width:100%;height:100%;background:#eee;display:flex;align-items:center;justify-content:center;color:#999">Preview</div>';
        return `<div class="design-card${sel}" data-ds="${ds.id}">
            <div class="design-card-label">${ds.label}</div>
            <div class="design-card-desc">${ds.desc}</div>
            <div class="design-card-preview"><div style="width:100%;aspect-ratio:16/9;overflow:hidden;position:relative">
                <div style="transform:scale(0.34);transform-origin:top left;width:1000px;height:562px">${previewHtml}</div>
            </div></div>
        </div>`;
    }).join("");

    picker.querySelectorAll(".design-card").forEach(card => {
        card.addEventListener("click", () => {
            state.designSystem = card.dataset.ds;
            picker.querySelectorAll(".design-card").forEach(c => c.classList.remove("selected"));
            card.classList.add("selected");
            document.getElementById("metaFields").style.display = "";
        });
    });

    if (state.designSystem) document.getElementById("metaFields").style.display = "";
    document.getElementById("metaTitle").value = state.deckTitle;
    document.getElementById("metaAuthor").value = state.deckAuthor;
}

// ── Sections Screen ────────────────────────────────────────────
function renderSections() {
    if (state.sections.length === 0) {
        state.sections.push({ id: genId(), name: "Introduction", color: "green" });
    }
    const list = document.getElementById("sectionsList");
    const colorOpts = SECTION_COLOR_OPTIONS[state.designSystem] || SECTION_COLOR_OPTIONS.slick;

    list.innerHTML = state.sections.map((sec, idx) => {
        const swatches = colorOpts.map(c => {
            const active = sec.color === c.value ? " active" : "";
            return `<div class="color-swatch${active}" style="background:${c.hex}" data-color="${c.value}" title="${c.label}"></div>`;
        }).join("");
        return `<div class="section-row" draggable="true" data-idx="${idx}" data-id="${sec.id}">
            <span class="drag-handle">\u2630</span>
            <span class="section-num">${String(idx + 1).padStart(2, "0")}</span>
            <input type="text" value="${esc(sec.name)}" data-idx="${idx}" placeholder="Section name">
            <div class="color-swatches">${swatches}</div>
            <button class="btn-remove" data-idx="${idx}" title="Remove">&times;</button>
        </div>`;
    }).join("");

    // Event: name input
    list.querySelectorAll("input[type=text]").forEach(inp => {
        inp.addEventListener("input", () => {
            state.sections[+inp.dataset.idx].name = inp.value;
        });
    });

    // Event: color swatch
    list.querySelectorAll(".color-swatch").forEach(sw => {
        sw.addEventListener("click", () => {
            const row = sw.closest(".section-row");
            const idx = +row.dataset.idx;
            state.sections[idx].color = sw.dataset.color;
            row.querySelectorAll(".color-swatch").forEach(s => s.classList.remove("active"));
            sw.classList.add("active");
        });
    });

    // Event: remove
    list.querySelectorAll(".btn-remove").forEach(btn => {
        btn.addEventListener("click", () => {
            if (state.sections.length <= 1) { showToast("Need at least one section"); return; }
            const idx = +btn.dataset.idx;
            const removed = state.sections.splice(idx, 1)[0];
            // Reassign orphaned slides to first section
            state.slides.forEach(sl => { if (sl.sectionId === removed.id) sl.sectionId = state.sections[0].id; });
            renderSections();
        });
    });

    // Drag reorder
    let dragIdx = null;
    list.querySelectorAll(".section-row").forEach(row => {
        row.addEventListener("dragstart", e => { dragIdx = +row.dataset.idx; row.classList.add("dragging"); });
        row.addEventListener("dragend", () => { row.classList.remove("dragging"); dragIdx = null; });
        row.addEventListener("dragover", e => { e.preventDefault(); });
        row.addEventListener("drop", e => {
            e.preventDefault();
            const targetIdx = +row.dataset.idx;
            if (dragIdx !== null && dragIdx !== targetIdx) {
                const [item] = state.sections.splice(dragIdx, 1);
                state.sections.splice(targetIdx, 0, item);
                renderSections();
            }
        });
    });
}

// ── Gallery Screen ─────────────────────────────────────────────
function renderGallery() {
    const grid = document.getElementById("galleryGrid");
    const catMap = {};
    for (const cat of SLIDE_CATEGORIES) catMap[cat.id] = { ...cat, types: [] };

    for (const [type, schema] of Object.entries(SLIDE_SCHEMAS)) {
        if (type === "section_divider") continue; // auto-generated
        const cat = catMap[schema.category] || catMap.content;
        cat.types.push({ type, schema });
    }

    grid.innerHTML = Object.values(catMap).filter(c => c.types.length > 0).map(cat => {
        const cards = cat.types.map(({ type, schema }) => {
            const data = buildDefaultData(type);
            const previewHtml = typeof PREVIEW_RENDERERS !== "undefined" && PREVIEW_RENDERERS[type]
                ? PREVIEW_RENDERERS[type](data, state.designSystem, "green") : '<div style="width:100%;height:100%;background:#f0f0f0"></div>';
            const badge = schema.generatesMultipleSlides
                ? `<div class="multi-slide-badge">${schema.generatesMultipleSlides === "N" ? "Multi-slide" : schema.generatesMultipleSlides + " slides"}</div>` : "";
            return `<div class="gallery-card" data-type="${type}">
                <div class="gallery-card-preview">
                    <div style="transform:scale(0.2);transform-origin:top left;width:1000px;height:562px;position:absolute">${previewHtml}</div>
                    ${badge}
                </div>
                <div class="gallery-card-label">${schema.label}</div>
                <div class="gallery-card-desc">${schema.description}</div>
            </div>`;
        }).join("");
        return `<div class="gallery-category">
            <div class="gallery-category-label">${cat.label}</div>
            <div class="gallery-category-items">${cards}</div>
        </div>`;
    }).join("");

    grid.querySelectorAll(".gallery-card").forEach(card => {
        card.addEventListener("click", () => {
            const type = card.dataset.type;
            const newSlide = {
                id: genId(),
                type: type,
                sectionId: state.sections[0]?.id || "",
                data: buildDefaultData(type),
            };
            state.slides.push(newSlide);
            state.editingSlideId = newSlide.id;
            state.editingIsNew = true;
            navigateTo("editor");
        });
    });
}

// ── Editor Screen ──────────────────────────────────────────────
function renderEditor() {
    const slide = state.slides.find(s => s.id === state.editingSlideId);
    if (!slide) { navigateTo("filmstrip"); return; }
    const schema = SLIDE_SCHEMAS[slide.type];
    if (!schema) return;

    // Header
    const header = document.getElementById("editorFormHeader");
    header.innerHTML = `<h3>${schema.label}</h3><p>${schema.description}</p>`;
    if (schema.generatesMultipleSlides) {
        const n = schema.generatesMultipleSlides === "N" ? "multiple" : schema.generatesMultipleSlides;
        header.innerHTML += `<p style="color:#5B2C8F;font-weight:600;font-size:12px;margin-top:4px">This creates ${n} slides in your final deck.</p>`;
    }

    // Form body
    const form = document.getElementById("editorForm");
    form.innerHTML = "";

    // Section picker
    const secPicker = document.createElement("div");
    secPicker.className = "section-picker";
    secPicker.innerHTML = `<label>Section</label>
        <select id="editorSectionSelect">${state.sections.map(s =>
            `<option value="${s.id}"${s.id === slide.sectionId ? " selected" : ""}>${s.name}</option>`
        ).join("")}</select>`;
    form.appendChild(secPicker);
    secPicker.querySelector("select").addEventListener("change", e => { slide.sectionId = e.target.value; });

    // Fields
    if (schema.layout === "wsn") {
        renderWSNLayout(form, schema, slide);
    } else if (schema.layout === "grid-2x2") {
        renderMatrixLayout(form, schema, slide);
    } else {
        for (const field of schema.fields) {
            form.appendChild(buildFieldElement(field, slide.data, slide));
        }
    }

    // Update done button text
    document.getElementById("editorDone").textContent = state.editingIsNew ? "Add to Deck" : "Save Changes";

    updatePreview();
}

function buildFieldElement(field, data, slide) {
    const wrap = document.createElement("div");
    wrap.className = "form-group";

    if (field.type === "text" || field.type === "number") {
        wrap.innerHTML = `<label>${field.label}</label>
            <input type="${field.type}" value="${esc(data[field.key] || "")}" data-key="${field.key}" placeholder="${field.label}">`;
        wrap.querySelector("input").addEventListener("input", e => {
            data[field.key] = e.target.value;
            updatePreview();
        });
        wrap.querySelector("input").addEventListener("focus", () => highlightPreviewField(field.key));
    }
    else if (field.type === "textarea") {
        wrap.innerHTML = `<label>${field.label}</label>
            <textarea data-key="${field.key}" placeholder="${field.label}">${esc(data[field.key] || "")}</textarea>`;
        wrap.querySelector("textarea").addEventListener("input", e => {
            data[field.key] = e.target.value;
            updatePreview();
        });
        wrap.querySelector("textarea").addEventListener("focus", () => highlightPreviewField(field.key));
    }
    else if (field.type === "select") {
        const opts = (field.options || []).map(o => `<option value="${o}"${data[field.key] === o ? " selected" : ""}>${o || "(none)"}</option>`).join("");
        wrap.innerHTML = `<label>${field.label}</label><select data-key="${field.key}">${opts}</select>`;
        wrap.querySelector("select").addEventListener("change", e => {
            data[field.key] = e.target.value;
            updatePreview();
        });
    }
    else if (field.type === "repeating") {
        wrap.className = "repeating-group";
        const items = data[field.key] || [];
        wrap.innerHTML = `<div class="repeating-group-header">
            <label>${field.label}</label>
            ${items.length < (field.maxItems || 99) ? `<button class="btn-add-item">+ Add</button>` : ""}
        </div><div class="repeating-items"></div>`;

        const renderItems = () => {
            const container = wrap.querySelector(".repeating-items");
            container.innerHTML = "";
            const currentItems = data[field.key] || [];
            currentItems.forEach((item, idx) => {
                const row = document.createElement("div");
                row.className = "repeating-item";
                const fieldsDiv = document.createElement("div");
                fieldsDiv.className = "item-fields" + (field.subfields.every(sf => sf.type === "text") && field.subfields.length <= 2 ? " inline" : "");
                for (const sf of field.subfields) {
                    const fg = document.createElement("div");
                    fg.className = "form-group";
                    if (sf.type === "textarea") {
                        fg.innerHTML = `<label>${sf.label}</label><textarea placeholder="${sf.label}">${esc(item[sf.key] || "")}</textarea>`;
                        fg.querySelector("textarea").addEventListener("input", e => { item[sf.key] = e.target.value; updatePreview(); });
                        fg.querySelector("textarea").addEventListener("focus", () => highlightPreviewField(field.key));
                    } else if (sf.type === "select") {
                        const opts = (sf.options || []).map(o => `<option value="${o}"${item[sf.key] === o ? " selected" : ""}>${o || "(none)"}</option>`).join("");
                        fg.innerHTML = `<label>${sf.label}</label><select>${opts}</select>`;
                        fg.querySelector("select").addEventListener("change", e => { item[sf.key] = e.target.value; updatePreview(); });
                    } else {
                        fg.innerHTML = `<label>${sf.label}</label><input type="text" value="${esc(item[sf.key] || "")}" placeholder="${sf.label}">`;
                        fg.querySelector("input").addEventListener("input", e => { item[sf.key] = e.target.value; updatePreview(); });
                        fg.querySelector("input").addEventListener("focus", () => highlightPreviewField(field.key));
                    }
                    fieldsDiv.appendChild(fg);
                }
                row.appendChild(fieldsDiv);
                if (currentItems.length > (field.minItems || 1)) {
                    const rmBtn = document.createElement("button");
                    rmBtn.className = "btn-remove-item";
                    rmBtn.textContent = "\u00d7";
                    rmBtn.addEventListener("click", () => { currentItems.splice(idx, 1); renderItems(); updatePreview(); });
                    row.appendChild(rmBtn);
                }
                container.appendChild(row);
            });
            // Update add button visibility
            const addBtn = wrap.querySelector(".btn-add-item");
            if (addBtn) addBtn.style.display = currentItems.length >= (field.maxItems || 99) ? "none" : "";
        };

        const addBtn = wrap.querySelector(".btn-add-item");
        if (addBtn) {
            addBtn.addEventListener("click", () => {
                const newItem = {};
                for (const sf of field.subfields) newItem[sf.key] = sf.default || "";
                data[field.key].push(newItem);
                renderItems();
                updatePreview();
            });
        }
        renderItems();
    }
    else if (field.type === "repeating-simple") {
        wrap.className = "repeating-group";
        const items = data[field.key] || [];
        wrap.innerHTML = `<div class="repeating-group-header">
            <label>${field.label}</label>
            ${items.length < (field.maxItems || 99) ? `<button class="btn-add-item">+ Add</button>` : ""}
        </div><div class="repeating-items"></div>`;

        const renderItems = () => {
            const container = wrap.querySelector(".repeating-items");
            container.innerHTML = "";
            const currentItems = data[field.key] || [];
            currentItems.forEach((val, idx) => {
                const row = document.createElement("div");
                row.className = "repeating-item";
                const fieldsDiv = document.createElement("div");
                fieldsDiv.className = "item-fields";
                const fg = document.createElement("div");
                fg.className = "form-group";
                fg.innerHTML = `<input type="text" value="${esc(val)}" placeholder="${field.itemLabel || "Item"} ${idx + 1}">`;
                fg.querySelector("input").addEventListener("input", e => { data[field.key][idx] = e.target.value; updatePreview(); });
                fg.querySelector("input").addEventListener("focus", () => highlightPreviewField(field.key));
                fieldsDiv.appendChild(fg);
                row.appendChild(fieldsDiv);
                if (currentItems.length > (field.minItems || 1)) {
                    const rmBtn = document.createElement("button");
                    rmBtn.className = "btn-remove-item";
                    rmBtn.textContent = "\u00d7";
                    rmBtn.addEventListener("click", () => { currentItems.splice(idx, 1); renderItems(); updatePreview(); });
                    row.appendChild(rmBtn);
                }
                container.appendChild(row);
            });
            const addBtn = wrap.querySelector(".btn-add-item");
            if (addBtn) addBtn.style.display = currentItems.length >= (field.maxItems || 99) ? "none" : "";
        };

        const addBtn = wrap.querySelector(".btn-add-item");
        if (addBtn) {
            addBtn.addEventListener("click", () => {
                data[field.key].push(field.default || "");
                renderItems();
                updatePreview();
            });
        }
        renderItems();
    }
    else if (field.type === "group") {
        wrap.className = "form-group";
        wrap.innerHTML = `<label class="group-label">${field.label}</label>`;
        if (!data[field.key]) data[field.key] = {};
        for (const sf of field.subfields) {
            wrap.appendChild(buildFieldElement(sf, data[field.key], slide));
        }
    }
    else if (field.type === "fixed-list") {
        // Handled by matrix/special layouts, but fallback
        wrap.className = "repeating-group";
        wrap.innerHTML = `<div class="repeating-group-header"><label>${field.label}</label></div><div class="repeating-items"></div>`;
        const container = wrap.querySelector(".repeating-items");
        const items = data[field.key] || [];
        items.forEach((item, idx) => {
            const row = document.createElement("div");
            row.className = "repeating-item";
            row.innerHTML = `<div class="item-fields"><strong style="font-size:12px;color:#555;margin-bottom:4px;display:block">${(field.itemLabels && field.itemLabels[idx]) || "Item " + (idx + 1)}</strong></div>`;
            const fieldsDiv = row.querySelector(".item-fields");
            for (const sf of field.subfields) {
                const fg = document.createElement("div");
                fg.className = "form-group";
                if (sf.type === "textarea") {
                    fg.innerHTML = `<label>${sf.label}</label><textarea placeholder="${sf.label}">${esc(item[sf.key] || "")}</textarea>`;
                    fg.querySelector("textarea").addEventListener("input", e => { item[sf.key] = e.target.value; updatePreview(); });
                } else {
                    fg.innerHTML = `<label>${sf.label}</label><input type="text" value="${esc(item[sf.key] || "")}" placeholder="${sf.label}">`;
                    fg.querySelector("input").addEventListener("input", e => { item[sf.key] = e.target.value; updatePreview(); });
                }
                fg.querySelector("input, textarea")?.addEventListener("focus", () => highlightPreviewField(field.key));
                fieldsDiv.appendChild(fg);
            }
            container.appendChild(row);
        });
    }

    return wrap;
}

// Special layouts
function renderWSNLayout(form, schema, slide) {
    // Title field first
    const titleField = schema.fields.find(f => f.key === "title");
    if (titleField) form.appendChild(buildFieldElement(titleField, slide.data, slide));

    const wsnContainer = document.createElement("div");
    wsnContainer.className = "wsn-groups";
    for (const groupKey of ["what", "soWhat", "nowWhat"]) {
        const field = schema.fields.find(f => f.key === groupKey);
        if (!field) continue;
        const col = document.createElement("div");
        col.className = "wsn-group";
        col.innerHTML = `<label class="group-label">${field.label}</label>`;
        if (!slide.data[groupKey]) slide.data[groupKey] = {};
        for (const sf of field.subfields) {
            col.appendChild(buildFieldElement(sf, slide.data[groupKey], slide));
        }
        wsnContainer.appendChild(col);
    }
    form.appendChild(wsnContainer);
}

function renderMatrixLayout(form, schema, slide) {
    // Title + axis fields first
    for (const field of schema.fields) {
        if (field.key === "quadrants") continue;
        form.appendChild(buildFieldElement(field, slide.data, slide));
    }
    const quadField = schema.fields.find(f => f.key === "quadrants");
    if (!quadField) return;
    const grid = document.createElement("div");
    grid.className = "matrix-grid";
    const items = slide.data.quadrants || [{}, {}, {}, {}];
    const labels = quadField.itemLabels || ["Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right"];
    items.forEach((item, idx) => {
        const cell = document.createElement("div");
        cell.className = "matrix-cell";
        cell.innerHTML = `<div class="matrix-cell-label">${labels[idx]}</div>`;
        for (const sf of quadField.subfields) {
            cell.appendChild(buildFieldElement(sf, item, slide));
        }
        grid.appendChild(cell);
    });
    form.appendChild(grid);
}

// ── Preview Update ─────────────────────────────────────────────
function updatePreview() {
    const slide = state.slides.find(s => s.id === state.editingSlideId);
    if (!slide) return;
    const container = document.getElementById("previewContainer");
    const section = state.sections.find(s => s.id === slide.sectionId);
    const sc = section ? section.color : "green";

    const renderer = typeof PREVIEW_RENDERERS !== "undefined" && PREVIEW_RENDERERS[slide.type];
    const html = renderer ? renderer(slide.data, state.designSystem, sc) : '<div class="preview-slide" style="display:flex;align-items:center;justify-content:center;color:#999">Preview not available</div>';

    container.innerHTML = `<div class="preview-scale-wrap" id="previewScaleWrap">${html}</div>`;
    scalePreview();
}

function scalePreview() {
    const container = document.getElementById("previewContainer");
    const wrap = document.getElementById("previewScaleWrap");
    if (!container || !wrap) return;
    const cw = container.clientWidth;
    const ch = container.clientHeight;
    const scale = Math.min(cw / 1000, ch / 562);
    wrap.style.transform = `scale(${scale})`;
    container.style.width = "100%";
    container.style.height = (562 * scale) + "px";
}

function highlightPreviewField(key) {
    document.querySelectorAll("[data-field].pv-highlight").forEach(el => el.classList.remove("pv-highlight"));
    const el = document.querySelector(`#previewContainer [data-field="${key}"]`);
    if (el) {
        el.classList.add("pv-highlight");
        setTimeout(() => el.classList.remove("pv-highlight"), 2000);
    }
}

// ── Filmstrip Screen ───────────────────────────────────────────
function renderFilmstrip() {
    const track = document.getElementById("filmstripTrack");
    const count = document.getElementById("filmstripCount");

    if (state.slides.length === 0) {
        count.textContent = "";
        track.innerHTML = `<div class="filmstrip-empty">
            <p><strong>No slides yet</strong></p>
            <p>Go to the slide gallery to add your first slide.</p>
            <button class="btn btn-primary" onclick="navigateTo('gallery')" style="margin-top:12px">+ Add Slide</button>
        </div>`;
        return;
    }

    // Count total output slides (wsn_reveal=3, progressive_reveal=N)
    let outputCount = 0;
    state.slides.forEach(sl => {
        const schema = SLIDE_SCHEMAS[sl.type];
        if (schema?.generatesMultipleSlides === "N") {
            outputCount += (sl.data.takeaways?.length || 1);
        } else if (schema?.generatesMultipleSlides) {
            outputCount += schema.generatesMultipleSlides;
        } else {
            outputCount += 1;
        }
    });
    outputCount += state.sections.length; // section dividers
    count.textContent = `${state.slides.length} slides configured \u2192 ${outputCount} slides in final deck`;

    track.innerHTML = "";
    let dragSrcId = null;

    state.sections.forEach((section, sIdx) => {
        // Section divider card
        const divider = document.createElement("div");
        divider.className = "filmstrip-divider";
        const colorHex = (SECTION_COLOR_OPTIONS[state.designSystem] || SECTION_COLOR_OPTIONS.slick)
            .find(c => c.value === section.color)?.hex || "#368727";
        divider.style.borderColor = colorHex;
        divider.innerHTML = `<div class="filmstrip-divider-num" style="color:${colorHex}">${String(sIdx + 1).padStart(2, "0")}</div>
            <div class="filmstrip-divider-title">${esc(section.name)}</div>`;
        track.appendChild(divider);

        // Slides in this section
        const sectionSlides = state.slides.filter(s => s.sectionId === section.id);
        sectionSlides.forEach(slide => {
            const card = document.createElement("div");
            card.className = "filmstrip-card";
            card.draggable = true;
            card.dataset.slideId = slide.id;

            const schema = SLIDE_SCHEMAS[slide.type];
            const renderer = typeof PREVIEW_RENDERERS !== "undefined" && PREVIEW_RENDERERS[slide.type];
            const previewHtml = renderer ? renderer(slide.data, state.designSystem, section.color) : "";

            card.innerHTML = `
                <div class="filmstrip-card-preview">
                    <div style="transform:scale(0.22);transform-origin:top left;width:1000px;height:562px;position:absolute">${previewHtml}</div>
                </div>
                <div class="filmstrip-card-info">
                    <span class="filmstrip-card-type">${schema?.label || slide.type}</span>
                    <span class="filmstrip-card-section" style="color:${colorHex}">${section.name}</span>
                </div>
                <div class="filmstrip-card-actions">
                    <button title="Edit" data-action="edit">\u270e</button>
                    <button title="Duplicate" data-action="dup">\u2398</button>
                    <button title="Delete" data-action="del">\u00d7</button>
                </div>`;

            // Events
            card.querySelector('[data-action="edit"]').addEventListener("click", e => {
                e.stopPropagation();
                state.editingSlideId = slide.id;
                state.editingIsNew = false;
                navigateTo("editor");
            });
            card.querySelector('[data-action="dup"]').addEventListener("click", e => {
                e.stopPropagation();
                const dup = { id: genId(), type: slide.type, sectionId: slide.sectionId, data: JSON.parse(JSON.stringify(slide.data)) };
                const idx = state.slides.indexOf(slide);
                state.slides.splice(idx + 1, 0, dup);
                renderFilmstrip();
                showToast("Slide duplicated");
            });
            card.querySelector('[data-action="del"]').addEventListener("click", e => {
                e.stopPropagation();
                const idx = state.slides.indexOf(slide);
                state.slides.splice(idx, 1);
                renderFilmstrip();
                showToast("Slide deleted");
            });
            card.addEventListener("click", () => {
                state.editingSlideId = slide.id;
                state.editingIsNew = false;
                navigateTo("editor");
            });

            // Drag and drop
            card.addEventListener("dragstart", e => {
                dragSrcId = slide.id;
                card.classList.add("dragging");
                e.dataTransfer.effectAllowed = "move";
            });
            card.addEventListener("dragend", () => { card.classList.remove("dragging"); dragSrcId = null; });
            card.addEventListener("dragover", e => {
                e.preventDefault();
                e.dataTransfer.dropEffect = "move";
                const rect = card.getBoundingClientRect();
                const mid = rect.left + rect.width / 2;
                card.classList.toggle("drag-over-left", e.clientX < mid);
                card.classList.toggle("drag-over-right", e.clientX >= mid);
            });
            card.addEventListener("dragleave", () => { card.classList.remove("drag-over-left", "drag-over-right"); });
            card.addEventListener("drop", e => {
                e.preventDefault();
                card.classList.remove("drag-over-left", "drag-over-right");
                if (!dragSrcId || dragSrcId === slide.id) return;
                const srcIdx = state.slides.findIndex(s => s.id === dragSrcId);
                const targetIdx = state.slides.findIndex(s => s.id === slide.id);
                if (srcIdx < 0 || targetIdx < 0) return;
                const [moved] = state.slides.splice(srcIdx, 1);
                // Also update section assignment to target's section
                moved.sectionId = slide.sectionId;
                const rect = card.getBoundingClientRect();
                const insertAfter = e.clientX >= rect.left + rect.width / 2;
                const newTargetIdx = state.slides.findIndex(s => s.id === slide.id);
                state.slides.splice(insertAfter ? newTargetIdx + 1 : newTargetIdx, 0, moved);
                renderFilmstrip();
            });

            track.appendChild(card);
        });
    });
}

// ── Build & Download ───────────────────────────────────────────
async function buildDeck() {
    if (state.slides.length === 0) { showToast("Add some slides first!"); return; }

    const slideConfigs = [];
    state.sections.forEach((section, idx) => {
        // Section divider
        slideConfigs.push(["section_divider", {
            title: section.name,
            subtitle: "",
            sectionNumber: idx + 1,
            sectionColor: section.color,
        }]);
        // Section slides
        const sectionSlides = state.slides.filter(s => s.sectionId === section.id);
        sectionSlides.forEach(slide => {
            slideConfigs.push([slide.type, { ...slide.data, sectionColor: section.color }]);
        });
    });

    showToast("Building deck...");
    try {
        const resp = await fetch("/build", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                designSystem: state.designSystem,
                deckTitle: state.deckTitle,
                slides: slideConfigs,
            }),
        });
        if (!resp.ok) throw new Error("Build failed");
        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = (state.deckTitle || "presentation") + ".pptx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        showToast("Deck downloaded!");
    } catch (err) {
        showToast("Error building deck: " + err.message);
    }
}

// ── Save / Load ────────────────────────────────────────────────
function saveProject() {
    const project = {
        designSystem: state.designSystem,
        deckTitle: state.deckTitle,
        deckAuthor: state.deckAuthor,
        sections: state.sections,
        slides: state.slides,
    };
    const json = JSON.stringify(project, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = (state.deckTitle || "project") + ".json";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showToast("Project saved!");
}

function loadProject(file) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const project = JSON.parse(e.target.result);
            if (!project.designSystem || !project.sections || !project.slides) {
                throw new Error("Invalid project file");
            }
            state.designSystem = project.designSystem;
            state.deckTitle = project.deckTitle || "";
            state.deckAuthor = project.deckAuthor || "";
            state.sections = project.sections;
            state.slides = project.slides;
            navigateTo("filmstrip");
            showToast("Project loaded!");
        } catch (err) {
            showToast("Error loading project: " + err.message);
        }
    };
    reader.readAsText(file);
}

// ── HTML escape ────────────────────────────────────────────────
function esc(s) {
    if (!s) return "";
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ── Wire up events ─────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
    // Setup screen
    document.getElementById("metaTitle").addEventListener("input", e => { state.deckTitle = e.target.value; });
    document.getElementById("metaAuthor").addEventListener("input", e => { state.deckAuthor = e.target.value; });
    document.getElementById("setupNext").addEventListener("click", () => {
        if (!state.designSystem) { showToast("Pick a design system first"); return; }
        navigateTo("sections");
    });

    // Sections screen
    document.getElementById("addSectionBtn").addEventListener("click", () => {
        const colors = SECTION_COLOR_OPTIONS[state.designSystem] || SECTION_COLOR_OPTIONS.slick;
        const color = colors[state.sections.length % colors.length].value;
        state.sections.push({ id: genId(), name: "Section " + (state.sections.length + 1), color });
        renderSections();
    });
    document.getElementById("sectionsBack").addEventListener("click", () => navigateTo("setup"));
    document.getElementById("sectionsNext").addEventListener("click", () => {
        if (state.sections.length === 0) { showToast("Add at least one section"); return; }
        if (state.slides.length > 0) navigateTo("filmstrip");
        else navigateTo("gallery");
    });

    // Gallery screen
    document.getElementById("galleryBack").addEventListener("click", () => navigateTo("sections"));
    document.getElementById("galleryToFilmstrip").addEventListener("click", () => navigateTo("filmstrip"));

    // Editor screen
    document.getElementById("editorCancel").addEventListener("click", () => {
        if (state.editingIsNew) {
            // Remove the slide we just created
            state.slides = state.slides.filter(s => s.id !== state.editingSlideId);
        }
        state.editingSlideId = null;
        if (state.slides.length > 0) navigateTo("filmstrip");
        else navigateTo("gallery");
    });
    document.getElementById("editorDone").addEventListener("click", () => {
        state.editingSlideId = null;
        navigateTo("filmstrip");
    });

    // Filmstrip screen
    document.getElementById("addSlideBtn").addEventListener("click", () => navigateTo("gallery"));
    document.getElementById("buildBtn").addEventListener("click", buildDeck);
    document.getElementById("filmstripBack").addEventListener("click", () => navigateTo("gallery"));

    // Save / Load
    document.getElementById("saveBtn").addEventListener("click", saveProject);
    document.getElementById("loadBtn").addEventListener("click", () => document.getElementById("loadFileInput").click());
    document.getElementById("loadFileInput").addEventListener("change", e => {
        if (e.target.files.length > 0) loadProject(e.target.files[0]);
        e.target.value = "";
    });

    // Window resize
    window.addEventListener("resize", () => {
        if (state.currentScreen === "editor") scalePreview();
    });

    // Initial render
    navigateTo("setup");
});
