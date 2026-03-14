/* previews.js -- HTML preview renderers for all 19 slide types */

function esc(s) {
    if (!s) return '';
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const PREVIEW_COLORS = ['#368727', '#3880F3', '#5B2C8F', '#04547C', '#D4A843'];
const PREVIEW_LIGHTS = ['#E2F0D9', '#E0EAFC', '#EDE4F5', '#E0EEF4', '#F5ECD4'];

const SECTION_COLOR_HEX = {
    green: '#368727', blue: '#3880F3', purple: '#5B2C8F', cobalt: '#04547C', gold: '#D4A843',
    red: '#C23B22', teal: '#1B7A6E', ochre: '#CC7A2E', slate: '#5A6A7A', plum: '#8E4585'
};

const EV2 = {
    DK_GREEN: '#044014', GOLD: '#D4A843', CHARCOAL: '#2D2D2D',
    BLACK: '#1A1A1A', MID: '#666', QUIET: '#999', FAINT: '#CCC',
    RULE_CLR: '#DDD', WARM: '#F0EDE6', COOL: '#EAEEF0',
    LIGHT_BOX: '#F7F6F4', COBALT: '#04547C', PURPLE: '#5B2C8F',
    CREAM: '#E8E0CC', LIGHT_GREEN: '#A8C4A0',
    ACCENTS: ['#044014', '#04547C', '#5B2C8F', '#D4A843', '#1B7A6E', '#C23B22', '#CC7A2E'],
    TF: "Georgia, 'Fidelity Slab', serif",
    BF: "Calibri, 'Fidelity Sans', sans-serif",
};

function _ev2TopRule(isDark) {
    return `<div style="position:absolute;left:0;top:0;width:100%;height:4px;background:${isDark ? EV2.GOLD : EV2.DK_GREEN}"></div>`;
}
function _ev2RunningHead(title) {
    return `<div style="position:absolute;left:65px;top:12px;width:700px;font-size:9px;font-weight:700;color:${EV2.QUIET};text-transform:uppercase;font-family:${EV2.BF};letter-spacing:0.5px">${esc(title)}</div>`;
}
function _ev2White(title) {
    return `<div class="preview-slide theme-editorial_v2" style="background:#fff">${_ev2TopRule(false)}${_ev2RunningHead(title)}`;
}
function _ev2Dark() {
    return `<div class="preview-slide theme-editorial_v2" style="background:${EV2.DK_GREEN}">${_ev2TopRule(true)}`;
}

/* ---------- tiny helpers ------------------------------------------------- */

function _accent(sc) {
    return SECTION_COLOR_HEX[sc] || '#368727';
}

function _slickBar(accent) {
    return `<div style="position:absolute;left:0;top:0;width:25px;height:100%;background:${accent}"></div>`;
}

function _slickTitle(title, accent, theme) {
    const isEditorial = theme === 'editorial';
    const bg = isEditorial ? '#FAF6F0' : '#fff';
    const font = isEditorial ? 'Georgia, serif' : '"Rockwell", Georgia, serif';
    const ruleColor = isEditorial ? '#044014' : accent;
    return `<div data-field="title" style="position:absolute;left:60px;top:22px;width:900px;font-size:20px;font-weight:700;color:#403F3E;font-family:${font}">${title}</div>`;
}

function _colorfulBar(accent) {
    return `<div class="pv-header-bar" style="background:${accent}"></div>`;
}

function _colorfulTitle(title) {
    return `<div data-field="title" style="position:absolute;left:30px;top:24px;width:940px;font-size:20px;font-weight:700;color:#fff">${title}</div>`;
}

function _slideOpen(theme, bg) {
    let bgStyle;
    if (bg) bgStyle = `background:${bg}`;
    else if (theme === 'editorial') bgStyle = 'background:#FAF6F0';
    else if (theme === 'editorial_v2') bgStyle = 'background:#fff';
    else if (theme === 'bold') bgStyle = 'background:#F2F0EC';
    else if (theme === 'noir') bgStyle = 'background:#0D0D0D';
    else bgStyle = 'background:#fff';
    return `<div class="preview-slide theme-${theme}" style="${bgStyle}">`;
}

function _placeholder(text) {
    return `<span style="color:#bbb;font-style:italic">${esc(text)}</span>`;
}

/* ---- Bold helpers ---- */
function _boldBar(accent) {
    return `<div style="position:absolute;left:0;top:0;width:30px;height:100%;background:${accent}"></div>`;
}

function _boldTitle(title, accent) {
    return `<div data-field="title" style="position:absolute;left:65px;top:22px;width:880px;font-size:21px;font-weight:800;color:#1A1A1A;font-family:'Trebuchet MS',sans-serif">${title}</div>`;
}

/* ---- Noir helpers ---- */
function _noirBar(accent) {
    return `<div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${accent}"></div>`;
}

function _noirTitle(title, accent) {
    return `<div data-field="title" style="position:absolute;left:45px;top:22px;width:920px;font-size:21px;font-weight:700;color:#F0F0F0;font-family:'Trebuchet MS',sans-serif">${title}</div>
        <div style="position:absolute;left:45px;top:57px;width:160px;height:4px;background:${accent}"></div>`;
}

/* ---------- stepper bar helper ------------------------------------------ */

function _stepperBar(labels, activeCount, accent, theme) {
    const n = labels.length;
    const colors = ['#368727', '#3880F3', '#5B2C8F', '#04547C', '#D4A843'];
    const startX = 80;
    const endX = 920;
    const spacing = n > 1 ? (endX - startX) / (n - 1) : 0;
    let html = '';
    for (let i = 0; i < n; i++) {
        const active = i < activeCount;
        const cx = startX + i * spacing;
        const color = active ? colors[i % 5] : '#ddd';
        const textColor = active ? '#fff' : '#999';
        html += `<div style="position:absolute;left:${cx - 14}px;top:0;width:28px;height:28px;border-radius:50%;background:${color};display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:${textColor}">${i + 1}</div>`;
        html += `<div style="position:absolute;left:${cx - 40}px;top:32px;width:80px;text-align:center;font-size:8px;color:${active ? color : '#999'};font-weight:${active ? '600' : '400'}">${esc(labels[i])}</div>`;
        if (i < n - 1) {
            const nx = startX + (i + 1) * spacing;
            html += `<div style="position:absolute;left:${cx + 16}px;top:13px;width:${nx - cx - 32}px;height:2px;background:#e5e5e5"></div>`;
        }
    }
    return html;
}

/* ---- Icon shape SVG helper --------------------------------------------- */
function _iconShape(idx, cx, cy, size, color) {
    const r = size / 2;
    switch (idx % 3) {
        case 0: /* circle */
            return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${color}"/>`;
        case 1: /* square */
            return `<rect x="${cx - r}" y="${cy - r}" width="${size}" height="${size}" rx="4" fill="${color}"/>`;
        case 2: /* triangle */
            return `<polygon points="${cx},${cy - r} ${cx + r},${cy + r} ${cx - r},${cy + r}" fill="${color}"/>`;
    }
}

/* ========================================================================= */
/*  PREVIEW_RENDERERS                                                        */
/* ========================================================================= */

const PREVIEW_RENDERERS = {

    /* --------------------------------------------------------------------- */
    /*  1. TITLE                                                             */
    /* --------------------------------------------------------------------- */
    title(data, theme, sc) {
        const t = esc(data.title || 'Presentation Title');
        const sub = esc(data.subtitle || '');
        const auth = esc(data.author || '');
        const dt = esc(data.date || '');
        const accent = _accent(sc);

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                <div style="position:absolute;left:0;top:0;width:12px;height:100%;background:${accent}"></div>
                <div data-field="title" style="position:absolute;left:50px;top:80px;width:900px;font-size:38px;font-weight:700;color:${accent}">${t}</div>
                <div style="position:absolute;left:50px;top:292px;display:flex;gap:0">
                    <div style="width:150px;height:5px;background:#368727"></div>
                    <div style="width:150px;height:5px;background:#3880F3;margin-left:4px"></div>
                    <div style="width:150px;height:5px;background:#5B2C8F;margin-left:4px"></div>
                    <div style="width:150px;height:5px;background:#04547C;margin-left:4px"></div>
                </div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:50px;top:310px;font-size:15px;color:#555">${sub}</div>` : ''}
                ${auth || dt ? `<div style="position:absolute;left:50px;top:420px;font-size:12px;color:#555">
                    ${auth ? `<span data-field="author">${auth}</span>` : ''}${auth && dt ? '<br>' : ''}${dt ? `<span data-field="date">${dt}</span>` : ''}
                </div>` : ''}
                ${data.imagePath ? `<img src="${esc(data.imagePath)}" style="position:absolute;right:20px;top:15px;width:200px;height:250px;object-fit:contain">` : `<div style="position:absolute;right:30px;top:20px;width:200px;height:260px;border:2px dashed #ccc;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#bbb;font-size:11px">Image</div>`}
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#044014">
                <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:#00FFC8"></div>
                <div data-field="title" style="position:absolute;left:90px;top:100px;width:860px;font-size:36px;font-weight:700;color:#F0F0F0;font-family:'Trebuchet MS',sans-serif">${t}</div>
                <div style="position:absolute;left:90px;top:265px;width:250px;height:4px;background:#00FFC8"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:90px;top:285px;font-size:15px;color:#ccc">${sub}</div>` : ''}
                ${auth ? `<div data-field="author" style="position:absolute;left:90px;top:410px;font-size:12px;color:#999;font-weight:600">${auth}</div>` : ''}
                ${dt ? `<div data-field="date" style="position:absolute;left:90px;top:445px;font-size:11px;color:#999">${dt}</div>` : ''}
                ${data.imagePath ? `<img src="${esc(data.imagePath)}" style="position:absolute;right:20px;top:15px;width:200px;height:250px;object-fit:contain">` : `<div style="position:absolute;right:30px;top:20px;width:200px;height:260px;border:2px dashed #ccc;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#bbb;font-size:11px">Image</div>`}
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                <div style="position:absolute;left:0;top:0;width:30px;height:100%;background:${accent}"></div>
                <div data-field="title" style="position:absolute;left:65px;top:100px;width:880px;font-size:36px;font-weight:800;color:#1A1A1A;font-family:'Trebuchet MS',sans-serif">${t}</div>
                <div style="position:absolute;left:65px;top:265px;width:250px;height:3px;background:${accent}"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:65px;top:280px;font-size:15px;color:#585858">${sub}</div>` : ''}
                ${auth ? `<div data-field="author" style="position:absolute;left:65px;top:400px;font-size:12px;color:#8A8A8A;font-weight:600">${auth}</div>` : ''}
                ${dt ? `<div data-field="date" style="position:absolute;left:65px;top:432px;font-size:11px;color:#8A8A8A">${dt}</div>` : ''}
                ${data.imagePath ? `<img src="${esc(data.imagePath)}" style="position:absolute;right:20px;top:15px;width:200px;height:250px;object-fit:contain">` : `<div style="position:absolute;right:30px;top:20px;width:200px;height:260px;border:2px dashed #F2F0EC;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#bbb;font-size:11px">Image</div>`}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const tSize = t.length > 60 ? 28 : 36;
            return `${_ev2Dark()}
                <div data-field="title" style="position:absolute;left:80px;top:60px;width:840px;font-size:${tSize}px;font-weight:700;color:#fff;font-family:${EV2.TF};line-height:1.25">${t}</div>
                <div style="position:absolute;left:80px;top:${tSize > 28 ? 250 : 220}px;width:200px;height:3px;background:${EV2.GOLD}"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:80px;top:${tSize > 28 ? 270 : 240}px;width:800px;font-size:14px;color:${EV2.CREAM};font-family:${EV2.BF};line-height:1.4">${sub}</div>` : ''}
                ${auth || dt ? `<div style="position:absolute;left:80px;bottom:30px;font-size:10px;color:${EV2.LIGHT_GREEN};font-family:${EV2.BF}">${auth ? `<span data-field="author">${auth}</span>` : ''}${auth && dt ? '<br>' : ''}${dt ? `<span data-field="date">${dt}</span>` : ''}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bgColor = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const titleFont = theme.startsWith('editorial') ? 'Georgia, serif' : '"Rockwell", Georgia, serif';
        const ruleColor = theme.startsWith('editorial') ? '#044014' : accent;
        return `<div class="preview-slide theme-${theme}" style="background:${bgColor}">
            <div style="position:absolute;left:0;top:0;width:25px;height:100%;background:${accent}"></div>
            <div data-field="title" style="position:absolute;left:90px;top:100px;width:860px;font-size:36px;font-weight:700;color:#403F3E;font-family:${titleFont}">${t}</div>
            <div style="position:absolute;left:90px;top:265px;width:250px;height:4px;background:${ruleColor}"></div>
            ${sub ? `<div data-field="subtitle" style="position:absolute;left:90px;top:285px;font-size:15px;color:#555">${sub}</div>` : ''}
            ${auth ? `<div data-field="author" style="position:absolute;left:90px;top:410px;font-size:12px;color:#555;font-weight:600">${auth}</div>` : ''}
            ${dt ? `<div data-field="date" style="position:absolute;left:90px;top:445px;font-size:11px;color:#555">${dt}</div>` : ''}
            ${data.imagePath ? `<img src="${esc(data.imagePath)}" style="position:absolute;right:20px;top:15px;width:200px;height:250px;object-fit:contain">` : `<div style="position:absolute;right:30px;top:20px;width:200px;height:260px;border:2px dashed #ccc;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#bbb;font-size:11px">Image</div>`}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  2. CLOSER                                                            */
    /* --------------------------------------------------------------------- */
    closer(data, theme, sc) {
        const t = esc(data.title || 'Thank You');
        const sub = esc(data.subtitle || '');
        const contact = esc(data.contact || '');
        const accent = _accent(sc);

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful" style="background:${accent}">
                <div data-field="title" style="position:absolute;left:50px;top:140px;width:900px;font-size:42px;font-weight:700;color:#fff;text-align:center">${t}</div>
                <div style="position:absolute;left:400px;top:280px;width:200px;height:4px;background:rgba(255,255,255,0.5)"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:50px;top:310px;width:900px;font-size:16px;color:rgba(255,255,255,0.85);text-align:center">${sub}</div>` : ''}
                ${contact ? `<div data-field="contact" style="position:absolute;left:50px;top:420px;width:900px;font-size:13px;color:rgba(255,255,255,0.75);text-align:center">${contact}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#044014">
                <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:#00FFC8"></div>
                <div data-field="title" style="position:absolute;left:50px;top:140px;width:900px;font-size:42px;font-weight:700;color:#F0F0F0;text-align:center;font-family:'Trebuchet MS',sans-serif">${t}</div>
                <div style="position:absolute;left:400px;top:280px;width:200px;height:4px;background:#00FFC8"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:50px;top:310px;width:900px;font-size:16px;color:rgba(240,240,240,0.7);text-align:center">${sub}</div>` : ''}
                ${contact ? `<div data-field="contact" style="position:absolute;left:50px;top:420px;width:900px;font-size:13px;color:rgba(240,240,240,0.5);text-align:center">${contact}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#2A2A2A">
                <div style="position:absolute;left:0;top:0;width:30px;height:100%;background:${accent}"></div>
                <div style="position:absolute;right:0;top:0;width:12px;height:100%;background:${accent};opacity:0.3"></div>
                <div data-field="title" style="position:absolute;left:50px;top:140px;width:900px;font-size:42px;font-weight:800;color:#FFFFFF;text-align:center;font-family:'Trebuchet MS',sans-serif">${t}</div>
                <div style="position:absolute;left:400px;top:280px;width:200px;height:5px;background:${accent}"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:50px;top:310px;width:900px;font-size:16px;color:rgba(255,255,255,0.75);text-align:center">${sub}</div>` : ''}
                ${contact ? `<div data-field="contact" style="position:absolute;left:50px;top:420px;width:900px;font-size:13px;color:rgba(255,255,255,0.6);text-align:center">${contact}</div>` : ''}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            return `${_ev2Dark()}
                <div data-field="title" style="position:absolute;left:80px;top:100px;width:840px;font-size:36px;font-weight:700;color:#fff;font-family:${EV2.TF};line-height:1.2">${t}</div>
                <div style="position:absolute;left:80px;top:230px;width:200px;height:3px;background:${EV2.GOLD}"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:80px;top:260px;width:800px;font-size:18px;font-style:italic;color:${EV2.CREAM};font-family:${EV2.TF};line-height:1.4">${sub}</div>` : ''}
                ${contact ? `<div data-field="contact" style="position:absolute;left:80px;bottom:30px;font-size:10px;color:${EV2.LIGHT_GREEN};font-family:${EV2.BF}">${contact}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bgColor = theme.startsWith('editorial') ? '#044014' : accent;
        return `<div class="preview-slide theme-${theme}" style="background:${bgColor}">
            <div data-field="title" style="position:absolute;left:50px;top:140px;width:900px;font-size:42px;font-weight:700;color:#fff;text-align:center;font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${t}</div>
            <div style="position:absolute;left:400px;top:280px;width:200px;height:4px;background:rgba(255,255,255,0.4)"></div>
            ${sub ? `<div data-field="subtitle" style="position:absolute;left:50px;top:310px;width:900px;font-size:16px;color:rgba(255,255,255,0.85);text-align:center">${sub}</div>` : ''}
            ${contact ? `<div data-field="contact" style="position:absolute;left:50px;top:420px;width:900px;font-size:13px;color:rgba(255,255,255,0.75);text-align:center">${contact}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  3. AGENDA                                                            */
    /* --------------------------------------------------------------------- */
    agenda(data, theme, sc) {
        const t = esc(data.title || 'Agenda');
        const items = data.items && data.items.length ? data.items : [{ title: '', detail: '' }, { title: '', detail: '' }, { title: '', detail: '' }];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const rowH = 52;
            const startY = 120;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 10);
                const topic = esc(item.title) || _placeholder('Topic ' + (i + 1));
                const detail = esc(item.detail);
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${rowH}px;background:${bg};border-radius:6px;overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:940px;height:4px;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:10px;width:30px;height:30px;border-radius:50%;background:${c};color:#fff;font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:56px;top:10px;font-size:13px;font-weight:600;color:#333">${topic}</div>
                    ${detail ? `<div style="position:absolute;left:56px;top:30px;font-size:11px;color:#666">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const rowH = 52;
            const startY = 80;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const topic = esc(item.title) || _placeholder('Topic ' + (i + 1));
                const detail = esc(item.detail);
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px;height:${rowH}px;background:#141414;border-radius:6px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:18px;top:10px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:56px;top:10px;font-size:13px;font-weight:600;color:#F0F0F0">${topic}</div>
                    ${detail ? `<div style="position:absolute;left:56px;top:30px;font-size:11px;color:#999">${detail}</div>` : ''}
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const rowH = 52;
            const startY = 80;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const topic = esc(item.title) || _placeholder('Topic ' + (i + 1));
                const detail = esc(item.detail);
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:890px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:18px;top:10px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:56px;top:10px;font-size:13px;font-weight:700;color:#1A1A1A">${topic}</div>
                    ${detail ? `<div style="position:absolute;left:56px;top:30px;font-size:11px;color:#555">${detail}</div>` : ''}
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            let cards = '';
            const cols = items.length <= 4 ? 2 : 3;
            const colW = Math.floor((900 - (cols - 1) * 20) / cols);
            items.forEach((item, i) => {
                const color = EV2.ACCENTS[i % 7];
                const col = i % cols;
                const row = Math.floor(i / cols);
                const x = 50 + col * (colW + 20);
                const y = 160 + row * 110;
                const topic = esc(item.title || item.label) || _placeholder('Topic ' + (i + 1));
                const det = esc(item.detail || '');
                cards += `<div style="position:absolute;left:${x}px;top:${y}px;width:${colW}px;height:90px">
                    <div style="position:absolute;left:0;top:0;width:4px;height:90px;background:${color}"></div>
                    <div style="position:absolute;left:14px;top:6px;font-size:16px;font-weight:700;color:${color};font-family:${EV2.TF}">${i + 1}</div>
                    <div style="position:absolute;left:14px;top:30px;width:${colW - 24}px;font-size:12px;font-weight:700;color:${EV2.CHARCOAL};font-family:${EV2.TF};line-height:1.2">${topic}</div>
                    ${det ? `<div style="position:absolute;left:14px;top:58px;width:${colW - 24}px;font-size:9px;color:${EV2.QUIET};font-family:${EV2.BF}">${det}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-editorial_v2" style="background:#fff">
                <div style="position:absolute;left:50px;top:12px;width:900px;text-align:center;font-size:9px;font-weight:700;color:${EV2.QUIET};text-transform:uppercase;font-family:${EV2.BF};letter-spacing:0.5px">${esc(t).toUpperCase()}</div>
                <div style="position:absolute;left:30px;top:32px;width:940px;height:2px;background:${EV2.CHARCOAL}"></div>
                <div style="position:absolute;left:30px;top:35px;width:940px;height:1px;background:${EV2.CHARCOAL}"></div>
                <div style="position:absolute;left:50px;top:52px;font-size:36px;font-weight:700;color:${EV2.BLACK};font-family:${EV2.TF}">Agenda</div>
                <div style="position:absolute;left:30px;top:120px;width:940px;height:1px;background:${EV2.RULE_CLR}"></div>
                ${cards}
            </div>`;
        }

        /* slick / editorial */
        let rows = '';
        const rowH = 52;
        const startY = 80;
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        items.forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 10);
            const topic = esc(item.title) || _placeholder('Topic ' + (i + 1));
            const detail = esc(item.detail);
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:18px;top:10px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                <div style="position:absolute;left:56px;top:10px;font-size:13px;font-weight:600;color:#333">${topic}</div>
                ${detail ? `<div style="position:absolute;left:56px;top:30px;font-size:11px;color:#666">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  4. IN_BRIEF                                                          */
    /* --------------------------------------------------------------------- */
    in_brief(data, theme, sc) {
        const t = esc(data.title || 'In Brief');
        const bullets = data.bullets && data.bullets.length ? data.bullets : ['', '', ''];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const rowH = 44;
            const startY = 120;
            bullets.forEach((b, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 10);
                const text = esc(b) || _placeholder('Point ' + (i + 1));
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${rowH}px;background:${bg};border-radius:6px;overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:940px;height:4px;background:${c}"></div>
                    <div style="position:absolute;left:18px;top:13px;font-size:13px;color:#333">${text}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="bullets">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const rowH = 44;
            const startY = 80;
            bullets.forEach((b, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const text = esc(b) || _placeholder('Point ' + (i + 1));
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px;height:${rowH}px;background:#141414;border-radius:6px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:18px;top:13px;font-size:13px;color:#F0F0F0">${text}</div>
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="bullets">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const rowH = 44;
            const startY = 80;
            bullets.forEach((b, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const text = esc(b) || _placeholder('Point ' + (i + 1));
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:890px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:18px;top:13px;font-size:13px;font-weight:600;color:#1A1A1A">${text}</div>
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="bullets">${rows}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const n2 = bullets.length || 1;
            const firstBullet = esc(bullets[0]) || _placeholder('Key point 1');
            return `${_ev2White(t)}
                <div style="position:absolute;right:65px;top:10px;font-size:12px;font-weight:700;color:${EV2.GOLD};font-family:${EV2.BF}">1 of ${n2}</div>
                <div style="position:absolute;left:65px;top:35px;width:870px;height:3px;background:${EV2.RULE_CLR}"></div>
                <div style="position:absolute;left:65px;top:35px;width:${Math.floor(870 / n2)}px;height:3px;background:${EV2.DK_GREEN}"></div>
                <div data-field="bullets" style="position:absolute;left:65px;top:65px;width:870px;font-size:22px;font-weight:700;color:${EV2.BLACK};font-family:${EV2.TF};line-height:1.4">${firstBullet}</div>
                <div style="position:absolute;left:65px;top:200px;width:250px;height:3px;background:${EV2.GOLD}"></div>
                ${n2 > 1 ? `<div style="position:absolute;left:65px;top:225px;font-size:8px;font-weight:700;color:${EV2.QUIET};font-family:${EV2.BF};letter-spacing:0.5px">PREVIOUSLY</div>
                ${bullets.slice(1, 4).map((b, j) => `<div style="position:absolute;left:65px;top:${246 + j * 30}px;font-size:10px;color:${EV2.QUIET};font-family:${EV2.BF}"><span style="font-weight:700;color:${EV2.FAINT};font-family:${EV2.TF};margin-right:8px">${j + 2}</span>${esc(b)}</div>`).join('')}` : ''}
            </div>`;
        }

        /* slick / editorial */
        let rows = '';
        const rowH = 44;
        const startY = 80;
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        bullets.forEach((b, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 10);
            const text = esc(b) || _placeholder('Point ' + (i + 1));
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:18px;top:13px;font-size:13px;color:#333">${text}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="bullets">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  5. SECTION_DIVIDER                                                   */
    /* --------------------------------------------------------------------- */
    section_divider(data, theme, sc) {
        const t = esc(data.title || 'Section Title');
        const sub = esc(data.subtitle || '');
        const num = esc(data.sectionNumber || '');
        const accent = _accent(sc);

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${accent}"></div>
                ${num ? `<div data-field="sectionNumber" style="position:absolute;left:90px;top:100px;font-size:72px;font-weight:700;color:rgba(240,240,240,0.1);font-family:'Trebuchet MS',sans-serif">${num}</div>` : ''}
                <div data-field="title" style="position:absolute;left:90px;top:${num ? 200 : 170}px;width:820px;font-size:36px;font-weight:700;color:#F0F0F0;font-family:'Trebuchet MS',sans-serif">${t}</div>
                <div style="position:absolute;left:90px;top:${num ? 260 : 240}px;width:180px;height:4px;background:${accent}"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:90px;top:${num ? 280 : 260}px;width:820px;font-size:15px;color:#999">${sub}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:${accent}">
                <div style="position:absolute;left:0;top:0;width:30px;height:100%;background:rgba(0,0,0,0.2)"></div>
                <div style="position:absolute;right:0;top:0;width:12px;height:100%;background:rgba(255,255,255,0.15)"></div>
                ${num ? `<div data-field="sectionNumber" style="position:absolute;left:90px;top:100px;font-size:72px;font-weight:800;color:rgba(255,255,255,0.2);font-family:'Trebuchet MS',sans-serif">${num}</div>` : ''}
                <div data-field="title" style="position:absolute;left:90px;top:${num ? 200 : 170}px;width:820px;font-size:36px;font-weight:800;color:#fff;font-family:'Trebuchet MS',sans-serif">${t}</div>
                <div style="position:absolute;left:90px;top:${num ? 260 : 240}px;width:180px;height:5px;background:rgba(255,255,255,0.4)"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:90px;top:${num ? 280 : 260}px;width:820px;font-size:15px;color:rgba(255,255,255,0.8)">${sub}</div>` : ''}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            return `${_ev2Dark()}
                ${num ? `<div data-field="sectionNumber" style="position:absolute;left:80px;top:60px;font-size:48px;font-weight:700;color:${EV2.GOLD};font-family:${EV2.TF}">${num}</div>` : ''}
                <div data-field="title" style="position:absolute;left:80px;top:${num ? 130 : 100}px;width:840px;font-size:32px;font-weight:700;color:#fff;font-family:${EV2.TF};line-height:1.25">${t}</div>
                <div style="position:absolute;left:80px;top:${num ? 220 : 190}px;width:200px;height:3px;background:${EV2.GOLD}"></div>
                ${sub ? `<div data-field="subtitle" style="position:absolute;left:80px;top:${num ? 240 : 210}px;width:840px;font-size:13px;font-style:italic;color:${EV2.CREAM};font-family:${EV2.BF}">${sub}</div>` : ''}
            </div>`;
        }

        return `<div class="preview-slide theme-${theme}" style="background:${accent}">
            ${num ? `<div data-field="sectionNumber" style="position:absolute;left:90px;top:100px;font-size:72px;font-weight:700;color:rgba(255,255,255,0.2);font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${num}</div>` : ''}
            <div data-field="title" style="position:absolute;left:90px;top:${num ? 200 : 170}px;width:820px;font-size:36px;font-weight:700;color:#fff;font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${t}</div>
            <div style="position:absolute;left:90px;top:${num ? 260 : 240}px;width:180px;height:4px;background:rgba(255,255,255,0.4)"></div>
            ${sub ? `<div data-field="subtitle" style="position:absolute;left:90px;top:${num ? 280 : 260}px;width:820px;font-size:15px;color:rgba(255,255,255,0.8)">${sub}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  6. STAT_CALLOUT                                                      */
    /* --------------------------------------------------------------------- */
    stat_callout(data, theme, sc) {
        const t = esc(data.title || 'Key Metric');
        const stat = esc(data.stat || '67%');
        const headline = esc(data.headline || '');
        const detail = esc(data.detail || '');
        const source = esc(data.source || '');
        const accent = _accent(sc);

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:50%;top:140px;transform:translateX(-50%);width:140px;height:140px;border-radius:50%;background:${accent};opacity:0.12"></div>
                <div data-field="stat" style="position:absolute;left:50px;top:140px;width:900px;text-align:center;font-size:96px;font-weight:700;color:${accent}">${stat}</div>
                ${headline ? `<div data-field="headline" style="position:absolute;left:50px;top:310px;width:900px;text-align:center;font-size:18px;font-weight:600;color:#333">${headline}</div>` : ''}
                ${detail ? `<div data-field="detail" style="position:absolute;left:100px;top:345px;width:800px;text-align:center;font-size:13px;color:#555">${detail}</div>` : ''}
                ${source ? `<div data-field="source" style="position:absolute;left:50px;top:510px;width:900px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:50%;top:120px;transform:translateX(-50%);width:140px;height:140px;border-radius:50%;background:${accent};opacity:0.2"></div>
                <div data-field="stat" style="position:absolute;left:45px;top:120px;width:920px;text-align:center;font-size:96px;font-weight:700;color:${accent};font-family:'Trebuchet MS',sans-serif">${stat}</div>
                ${headline ? `<div data-field="headline" style="position:absolute;left:45px;top:300px;width:920px;text-align:center;font-size:18px;font-weight:600;color:#F0F0F0">${headline}</div>` : ''}
                ${detail ? `<div data-field="detail" style="position:absolute;left:100px;top:335px;width:800px;text-align:center;font-size:13px;color:#999">${detail}</div>` : ''}
                ${source ? `<div data-field="source" style="position:absolute;left:45px;top:510px;width:920px;text-align:center;font-size:10px;color:#666">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:50%;top:120px;transform:translateX(-50%);width:140px;height:140px;border-radius:50%;background:${accent};opacity:0.12"></div>
                <div data-field="stat" style="position:absolute;left:65px;top:120px;width:880px;text-align:center;font-size:96px;font-weight:800;color:${accent};font-family:'Trebuchet MS',sans-serif">${stat}</div>
                ${headline ? `<div data-field="headline" style="position:absolute;left:65px;top:300px;width:880px;text-align:center;font-size:18px;font-weight:700;color:#1A1A1A">${headline}</div>` : ''}
                ${detail ? `<div data-field="detail" style="position:absolute;left:100px;top:335px;width:800px;text-align:center;font-size:13px;color:#555">${detail}</div>` : ''}
                ${source ? `<div data-field="source" style="position:absolute;left:65px;top:510px;width:880px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            return `${_ev2White(t)}
                <div data-field="stat" style="position:absolute;left:65px;top:50px;width:500px;font-size:96px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.TF};line-height:1;display:flex;align-items:center;height:350px">${stat}</div>
                <div style="position:absolute;left:580px;top:100px;width:360px">
                    ${headline ? `<div data-field="headline" style="font-size:18px;font-weight:700;color:${EV2.CHARCOAL};font-family:${EV2.TF};margin-bottom:10px;line-height:1.3">${headline}</div>` : ''}
                    <div style="width:120px;height:2px;background:${EV2.GOLD};margin:10px 0"></div>
                    ${detail ? `<div data-field="detail" style="font-size:11px;color:${EV2.MID};font-family:${EV2.BF};line-height:1.5">${detail}</div>` : ''}
                    ${source ? `<div data-field="source" style="font-size:8px;font-style:italic;color:${EV2.QUIET};font-family:${EV2.BF};margin-top:15px">${source}</div>` : ''}
                </div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:50%;top:120px;transform:translateX(-50%);width:140px;height:140px;border-radius:50%;background:${accent};opacity:0.12"></div>
            <div data-field="stat" style="position:absolute;left:60px;top:120px;width:900px;text-align:center;font-size:96px;font-weight:700;color:${accent};font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${stat}</div>
            ${headline ? `<div data-field="headline" style="position:absolute;left:60px;top:300px;width:900px;text-align:center;font-size:18px;font-weight:600;color:#333">${headline}</div>` : ''}
            ${detail ? `<div data-field="detail" style="position:absolute;left:100px;top:335px;width:800px;text-align:center;font-size:13px;color:#555">${detail}</div>` : ''}
            ${source ? `<div data-field="source" style="position:absolute;left:60px;top:510px;width:900px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  7. QUOTE                                                             */
    /* --------------------------------------------------------------------- */
    quote(data, theme, sc) {
        const t = esc(data.title || 'In Their Words');
        const q = esc(data.quote || '');
        const attr = esc(data.attribution || '');
        const ctx = esc(data.context || '');
        const accent = _accent(sc);

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:60px;top:120px;font-size:72px;color:${accent};font-family:Georgia,serif;line-height:1">\u201C</div>
                <div data-field="quote" style="position:absolute;left:80px;top:170px;width:840px;font-size:18px;font-style:italic;color:#333;line-height:1.6">${q || _placeholder('Quote text...')}</div>
                <div style="position:absolute;left:80px;top:390px;width:200px;height:3px;background:${accent}"></div>
                ${attr ? `<div data-field="attribution" style="position:absolute;left:80px;top:405px;font-size:13px;font-weight:600;color:#333">\u2014 ${attr}</div>` : ''}
                ${ctx ? `<div data-field="context" style="position:absolute;left:80px;top:425px;font-size:11px;color:#777">${ctx}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:65px;top:80px;font-size:72px;color:${accent};font-family:Georgia,serif;line-height:1">\u201C</div>
                <div data-field="quote" style="position:absolute;left:85px;top:140px;width:840px;font-size:18px;font-style:italic;color:#F0F0F0;line-height:1.6">${q || _placeholder('Quote text...')}</div>
                <div style="position:absolute;left:85px;top:380px;width:200px;height:3px;background:${accent}"></div>
                ${attr ? `<div data-field="attribution" style="position:absolute;left:85px;top:395px;font-size:13px;font-weight:600;color:#F0F0F0">\u2014 ${attr}</div>` : ''}
                ${ctx ? `<div data-field="context" style="position:absolute;left:85px;top:415px;font-size:11px;color:#999">${ctx}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:65px;top:80px;width:4px;height:180px;background:${accent};border-radius:2px"></div>
                <div data-field="quote" style="position:absolute;left:85px;top:90px;width:840px;font-size:18px;font-style:italic;color:#1A1A1A;line-height:1.6">${q || _placeholder('Quote text...')}</div>
                <div style="position:absolute;left:85px;top:360px;width:150px;height:4px;background:${accent}"></div>
                ${attr ? `<div data-field="attribution" style="position:absolute;left:85px;top:375px;font-size:12px;font-weight:700;color:#1A1A1A">\u2014 ${attr}</div>` : ''}
                ${ctx ? `<div data-field="context" style="position:absolute;left:85px;top:393px;font-size:11px;color:#555">${ctx}</div>` : ''}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const q2 = esc(data.quote) || _placeholder('Quote text...');
            const attr2 = esc(data.attribution);
            return `${_ev2Dark()}
                <div style="position:absolute;left:80px;top:30px;font-size:120px;color:${EV2.GOLD};font-family:${EV2.TF};line-height:1">\u201C</div>
                <div data-field="quote" style="position:absolute;left:100px;top:150px;width:800px;font-size:22px;font-style:italic;color:#fff;line-height:1.6;font-family:${EV2.TF}">${q2}</div>
                ${attr2 ? `<div style="position:absolute;left:100px;top:380px;width:80px;height:2px;background:${EV2.GOLD}"></div><div data-field="attribution" style="position:absolute;left:195px;top:372px;font-size:12px;font-weight:700;color:${EV2.GOLD};font-family:${EV2.BF}">${attr2}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:50px;top:80px;width:900px;text-align:center;font-size:72px;color:${accent};font-family:Georgia,serif;line-height:1">\u201C</div>
            <div data-field="quote" style="position:absolute;left:100px;top:140px;width:820px;font-size:18px;font-style:italic;color:#333;line-height:1.6;text-align:center">${q || _placeholder('Quote text...')}</div>
            <div style="position:absolute;left:50%;top:380px;transform:translateX(-50%);width:200px;height:3px;background:${theme.startsWith('editorial') ? '#044014' : accent}"></div>
            ${attr ? `<div data-field="attribution" style="position:absolute;left:100px;top:395px;width:820px;text-align:center;font-size:13px;font-weight:600;color:#333">\u2014 ${attr}</div>` : ''}
            ${ctx ? `<div data-field="context" style="position:absolute;left:100px;top:415px;width:820px;text-align:center;font-size:11px;color:#777">${ctx}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  8. COMPARISON                                                        */
    /* --------------------------------------------------------------------- */
    comparison(data, theme, sc) {
        const t = esc(data.title || 'Comparison');
        const accent = _accent(sc);
        const leftLabel = esc(data.leftLabel || 'Before');
        const rightLabel = esc(data.rightLabel || 'After');
        const leftItems = data.leftItems || ['', '', ''];
        const rightItems = data.rightItems || ['', '', ''];

        const leftColor = accent === '#368727' ? '#04547C' : accent;
        const rightColor = '#368727';

        function renderCard(x, w, color, label, items) {
            let itemsHtml = '';
            items.forEach((item, i) => {
                const text = esc(typeof item === 'string' ? item : (item.text || '')) || _placeholder('Item ' + (i + 1));
                itemsHtml += `<div style="padding:6px 14px;font-size:12px;font-weight:600;color:#333">${text}</div>`;
            });
            return `<div style="position:absolute;left:${x}px;top:75px;width:${w}px;height:340px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.06)">
                <div style="height:40px;background:${color};display:flex;align-items:center;padding:0 14px">
                    <span style="font-size:13px;font-weight:700;color:#fff">${label}</span>
                </div>
                <div data-field="items" style="padding:8px 0">${itemsHtml}</div>
            </div>`;
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                ${renderCard(30, 460, leftColor, leftLabel, leftItems)}
                ${renderCard(510, 460, rightColor, rightLabel, rightItems)}
            </div>`;
        }
        if (theme === 'noir') {
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                ${renderCard(60, 430, leftColor, leftLabel, leftItems).replace(/#fff/g, '#141414').replace(/#333/g, '#F0F0F0')}
                ${renderCard(510, 430, rightColor, rightLabel, rightItems).replace(/#fff/g, '#141414').replace(/#333/g, '#F0F0F0')}
            </div>`;
        }
        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                ${renderCard(60, 430, leftColor, leftLabel, leftItems)}
                ${renderCard(510, 430, rightColor, rightLabel, rightItems)}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const ll = esc(data.leftLabel || 'Option A');
            const rl = esc(data.rightLabel || 'Option B');
            const n2 = Math.max(leftItems.length, rightItems.length);
            let rows = '';
            for (let i = 0; i < Math.min(n2, 5); i++) {
                const y = 115 + i * 68;
                const lText = esc(typeof leftItems[i] === 'string' ? leftItems[i] : (leftItems[i]?.text || '')) || '';
                const rText = esc(typeof rightItems[i] === 'string' ? rightItems[i] : (rightItems[i]?.text || '')) || '';
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:460px;height:58px;background:${EV2.WARM}"><div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${EV2.DK_GREEN}"></div><div style="padding:10px 12px 0 18px;font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.4">${lText}</div></div>
                <div style="position:absolute;left:505px;top:${y}px;width:460px;height:58px;background:${EV2.COOL}"><div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${EV2.COBALT}"></div><div style="padding:10px 12px 0 18px;font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.4">${rText}</div></div>`;
            }
            return `${_ev2White(t)}
                <div style="position:absolute;left:50px;top:42px;font-size:14px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.TF}">${ll}</div>
                <div style="position:absolute;left:50px;top:65px;width:140px;height:2px;background:${EV2.DK_GREEN}"></div>
                <div style="position:absolute;left:525px;top:42px;font-size:14px;font-weight:700;color:${EV2.COBALT};font-family:${EV2.TF}">${rl}</div>
                <div style="position:absolute;left:525px;top:65px;width:140px;height:2px;background:${EV2.COBALT}"></div>
                ${rows}
            </div>`;
        }

        // slick / editorial
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            ${renderCard(60, 430, leftColor, leftLabel, leftItems)}
            ${renderCard(510, 430, rightColor, rightLabel, rightItems)}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  9. TEXT_GRAPH                                                         */
    /* --------------------------------------------------------------------- */
    text_graph(data, theme, sc) {
        const t = esc(data.title || 'Title');
        const paragraphs = data.text && data.text.length ? data.text : ['', ''];
        const accent = _accent(sc);

        const pHtml = paragraphs.map((p, i) => {
            const text = esc(p) || _placeholder('Paragraph ' + (i + 1));
            return `<div style="margin-bottom:12px;font-size:12px;color:#444;line-height:1.5">${text}</div>`;
        }).join('');

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="text" style="position:absolute;left:30px;top:120px;width:440px">${pHtml}</div>
                <div style="position:absolute;left:490px;top:120px;width:1px;height:400px;background:#ddd"></div>
                <div style="position:absolute;left:510px;top:120px;width:460px;height:400px;border:2px dashed #ccc;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#aaa;font-size:14px">Chart area</div>
            </div>`;
        }

        if (theme === 'noir') {
            const noirPHtml = paragraphs.map((p, i) => {
                const text = esc(p) || _placeholder('Paragraph ' + (i + 1));
                return `<div style="margin-bottom:12px;font-size:12px;color:#F0F0F0;line-height:1.5">${text}</div>`;
            }).join('');
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="text" style="position:absolute;left:45px;top:80px;width:420px">${noirPHtml}</div>
                <div style="position:absolute;left:485px;top:80px;width:1px;height:440px;background:#2A2A2A"></div>
                <div style="position:absolute;left:505px;top:80px;width:440px;height:440px;border:2px dashed #2A2A2A;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#666;font-size:14px">Chart area</div>
            </div>`;
        }

        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="text" style="position:absolute;left:65px;top:80px;width:420px">${pHtml}</div>
                <div style="position:absolute;left:505px;top:80px;width:1px;height:440px;background:#ddd"></div>
                <div style="position:absolute;left:525px;top:80px;width:430px;height:440px;border:2px dashed #ccc;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#aaa;font-size:14px">Chart area</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            return `${_ev2White(t)}
                <div data-field="text" style="position:absolute;left:60px;top:80px;width:420px">${pHtml}</div>
                <div style="position:absolute;left:500px;top:80px;width:1px;height:440px;background:${EV2.RULE_CLR}"></div>
                <div style="position:absolute;left:520px;top:80px;width:440px;height:440px;border:2px dashed ${EV2.FAINT};border-radius:8px;display:flex;align-items:center;justify-content:center;color:${EV2.QUIET};font-size:14px;font-family:${EV2.BF}">Chart area</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="text" style="position:absolute;left:60px;top:80px;width:420px">${pHtml}</div>
            <div style="position:absolute;left:500px;top:80px;width:1px;height:440px;background:#ddd"></div>
            <div style="position:absolute;left:520px;top:80px;width:440px;height:440px;border:2px dashed #ccc;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#aaa;font-size:14px">Chart area</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  10. PROCESS_FLOW                                                     */
    /* --------------------------------------------------------------------- */
    process_flow(data, theme, sc) {
        const t = esc(data.title || 'Process');
        const steps = data.steps && data.steps.length ? data.steps : [{ title: '', detail: '' }, { title: '', detail: '' }, { title: '', detail: '' }];
        const accent = _accent(sc);
        const n = steps.length;
        const labels = steps.map((s, i) => s.title || 'Step ' + (i + 1));

        function buildCards(leftX, totalW, cardBg, titleColor, detailColor, borderStyle) {
            const cardW = Math.min(180, Math.floor((totalW - (n - 1) * 16) / n));
            let cards = '';
            steps.forEach((step, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const x = leftX + i * (cardW + 16);
                const title = esc(step.title) || _placeholder('Step ' + (i + 1));
                const detail = esc(step.detail);
                cards += `<div style="position:absolute;left:${x}px;top:140px;width:${cardW}px;height:330px;background:${cardBg};border-radius:8px;overflow:hidden;${borderStyle}">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:14px 12px 6px;font-size:13px;font-weight:700;color:${c}">${title}</div>
                    ${detail ? `<div style="padding:0 12px;font-size:11px;color:${detailColor};line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return cards;
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, n, accent, theme)}</div>
                <div data-field="steps">${buildCards(30, 940, PREVIEW_LIGHTS[0], '#333', '#666', '')}</div>
            </div>`;
        }

        if (theme === 'noir') {
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, n, accent, theme)}</div>
                <div data-field="steps">${buildCards(45, 920, '#141414', '#F0F0F0', '#999', 'border:1px solid #2A2A2A;')}</div>
            </div>`;
        }

        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, n, accent, theme)}</div>
                <div data-field="steps">${buildCards(65, 890, '#fff', '#1A1A1A', '#555', 'box-shadow:0 1px 4px rgba(0,0,0,0.1);')}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const cardW = Math.min(200, Math.floor((900 - (n - 1) * 35) / n));
            let cards = '';
            steps.forEach((step, i) => {
                const color = EV2.ACCENTS[i % 7];
                const x = 50 + i * (cardW + 35);
                const lbl = esc(step.label || step.title) || _placeholder('Step ' + (i + 1));
                const det = esc(step.detail || '');
                cards += `<div style="position:absolute;left:${x}px;top:90px;width:${cardW}px;height:350px">
                    <div style="width:100%;height:6px;background:${color}"></div>
                    <div style="position:absolute;left:${cardW/2 - 18}px;top:20px;width:36px;height:36px;border-radius:50%;background:${color};color:#fff;font-size:14px;font-weight:700;display:flex;align-items:center;justify-content:center;font-family:${EV2.BF}">${i + 1}</div>
                    <div style="position:absolute;left:15px;top:72px;width:${cardW - 30}px;font-size:13px;font-weight:700;color:${EV2.CHARCOAL};font-family:${EV2.TF};line-height:1.3">${lbl}</div>
                    <div style="position:absolute;left:15px;top:120px;width:${Math.floor(cardW * 0.4)}px;height:2px;background:${color}"></div>
                    ${det ? `<div style="position:absolute;left:15px;top:135px;width:${cardW - 30}px;font-size:10px;color:${EV2.MID};font-family:${EV2.BF};line-height:1.5">${det}</div>` : ''}
                </div>`;
                if (i < n - 1) cards += `<div style="position:absolute;left:${x + cardW + 8}px;top:110px;font-size:16px;font-weight:700;color:${EV2.GOLD};font-family:${EV2.TF}">\u2192</div>`;
            });
            return `<div class="preview-slide theme-editorial_v2" style="background:#fff">
                ${_ev2TopRule(false)}
                <div style="position:absolute;left:50px;top:18px;font-size:20px;font-weight:700;color:${EV2.CHARCOAL};font-family:${EV2.TF}">${t}</div>
                ${cards}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, n, accent, theme)}</div>
            <div data-field="steps">${buildCards(60, 900, '#fff', '#333', '#666', 'box-shadow:0 1px 4px rgba(0,0,0,0.08);')}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  10b. PROCESS_FLOW_REVEAL (step-1 state preview)                      */
    /* --------------------------------------------------------------------- */
    process_flow_reveal(data, theme, sc) {
        const t = esc(data.title || 'Process');
        const steps = data.steps && data.steps.length ? data.steps : [{title:'',detail:''},{title:'',detail:''},{title:'',detail:''}];
        const accent = _accent(sc);
        const labels = steps.map((s, i) => s.title || 'Step ' + (i + 1));

        const step = steps[0];
        const stepTitle = esc(step.title) || _placeholder('Step 1');
        const detail = esc(step.detail);
        const c = PREVIEW_COLORS[0];

        function featuredCard(leftX, w, cardBg, titleColor, detailColor, borderStyle) {
            return `<div style="position:absolute;left:${leftX}px;top:140px;width:${w}px;height:330px;background:${cardBg};border-radius:8px;overflow:hidden;${borderStyle}">
                <div style="height:4px;background:${c}"></div>
                <div style="padding:14px 18px 6px;font-size:15px;font-weight:700;color:${c}">${stepTitle}</div>
                ${detail ? `<div style="padding:0 18px;font-size:12px;color:${detailColor};line-height:1.5">${detail}</div>` : ''}
            </div>`;
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, PREVIEW_LIGHTS[0], '#333', '#666', '')}
            </div>`;
        }
        if (theme === 'noir') {
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, '#141414', '#F0F0F0', '#999', 'border:1px solid #2A2A2A;')}
            </div>`;
        }
        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, '#fff', '#1A1A1A', '#555', 'box-shadow:0 1px 4px rgba(0,0,0,0.1);')}
            </div>`;
        }
        if (theme === 'editorial_v2') {
            return `${_ev2White(t)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, EV2.LIGHT_BOX, EV2.CHARCOAL, EV2.MID, '')}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
            ${featuredCard(200, 600, '#fff', '#333', '#666', 'box-shadow:0 1px 4px rgba(0,0,0,0.08);')}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  11. MATRIX                                                           */
    /* --------------------------------------------------------------------- */
    matrix(data, theme, sc) {
        const t = esc(data.title || 'Framework');
        const xAxis = esc(data.xAxis || '');
        const yAxis = esc(data.yAxis || '');
        const quads = data.quadrants && data.quadrants.length === 4 ? data.quadrants : [
            { label: '', detail: '' }, { label: '', detail: '' },
            { label: '', detail: '' }, { label: '', detail: '' }
        ];
        const accent = _accent(sc);

        const positions = [
            { x: 0, y: 0 },  // TL
            { x: 1, y: 0 },  // TR
            { x: 0, y: 1 },  // BL
            { x: 1, y: 1 },  // BR
        ];
        const quadLabels = ['Top-Left', 'Top-Right', 'Bottom-Left', 'Bottom-Right'];

        if (theme === 'colorful') {
            const gridLeft = 30;
            const gridTop = 120;
            const cellW = 460;
            const cellH = 195;
            const gap = 12;

            let cells = '';
            quads.forEach((q, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const pos = positions[i];
                const cx = gridLeft + pos.x * (cellW + gap);
                const cy = gridTop + pos.y * (cellH + gap);
                const label = esc(q.label) || _placeholder(quadLabels[i]);
                const detail = esc(q.detail);
                cells += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:${bg};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:10px 14px 4px;font-size:13px;font-weight:700;color:${c}">${label}</div>
                    ${detail ? `<div style="padding:0 14px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                ${xAxis ? `<div data-field="xAxis" style="position:absolute;left:30px;top:530px;width:940px;text-align:center;font-size:11px;color:#888">${xAxis} \u2192</div>` : ''}
                ${yAxis ? `<div data-field="yAxis" style="position:absolute;left:6px;top:120px;width:18px;font-size:11px;color:#888;writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;height:400px">\u2190 ${yAxis}</div>` : ''}
                <div data-field="quadrants">${cells}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const gridLeft = 45;
            const gridTop = 80;
            const cellW = 450;
            const cellH = 210;
            const gap = 12;

            let cells = '';
            quads.forEach((q, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const pos = positions[i];
                const cx = gridLeft + pos.x * (cellW + gap);
                const cy = gridTop + pos.y * (cellH + gap);
                const label = esc(q.label) || _placeholder(quadLabels[i]);
                const detail = esc(q.detail);
                cells += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:#141414;border-radius:8px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="padding:10px 14px 4px 18px;font-size:13px;font-weight:700;color:${c}">${label}</div>
                    ${detail ? `<div style="padding:0 14px 0 18px;font-size:11px;color:#999;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                ${xAxis ? `<div data-field="xAxis" style="position:absolute;left:45px;top:530px;width:920px;text-align:center;font-size:11px;color:#666">${xAxis} \u2192</div>` : ''}
                ${yAxis ? `<div data-field="yAxis" style="position:absolute;left:16px;top:80px;width:18px;font-size:11px;color:#666;writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;height:430px">\u2190 ${yAxis}</div>` : ''}
                <div data-field="quadrants">${cells}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const gridLeft = 65;
            const gridTop = 80;
            const cellW = 430;
            const cellH = 210;
            const gap = 12;

            let cells = '';
            quads.forEach((q, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const pos = positions[i];
                const cx = gridLeft + pos.x * (cellW + gap);
                const cy = gridTop + pos.y * (cellH + gap);
                const label = esc(q.label) || _placeholder(quadLabels[i]);
                const detail = esc(q.detail);
                cells += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.1)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="padding:10px 14px 4px 18px;font-size:13px;font-weight:800;color:${c}">${label}</div>
                    ${detail ? `<div style="padding:0 14px 0 18px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                ${xAxis ? `<div data-field="xAxis" style="position:absolute;left:65px;top:530px;width:880px;text-align:center;font-size:11px;color:#888">${xAxis} \u2192</div>` : ''}
                ${yAxis ? `<div data-field="yAxis" style="position:absolute;left:36px;top:80px;width:18px;font-size:11px;color:#888;writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;height:430px">\u2190 ${yAxis}</div>` : ''}
                <div data-field="quadrants">${cells}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const gridLeft2 = 60;
            const gridTop2 = 80;
            const cellW2 = 440;
            const cellH2 = 210;
            const gap2 = 12;
            let cells2 = '';
            quads.forEach((q, i) => {
                const color = EV2.ACCENTS[i % 7];
                const pos = positions[i];
                const cx = gridLeft2 + pos.x * (cellW2 + gap2);
                const cy = gridTop2 + pos.y * (cellH2 + gap2);
                const label = esc(q.label) || _placeholder(quadLabels[i]);
                const detail = esc(q.detail);
                cells2 += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW2}px;height:${cellH2}px;background:${EV2.LIGHT_BOX};overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${color}"></div>
                    <div style="padding:10px 14px 4px 18px;font-size:13px;font-weight:700;color:${color};font-family:${EV2.TF}">${label}</div>
                    ${detail ? `<div style="padding:0 14px 0 18px;font-size:11px;color:${EV2.MID};font-family:${EV2.BF};line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `${_ev2White(t)}
                ${xAxis ? `<div data-field="xAxis" style="position:absolute;left:60px;top:530px;width:900px;text-align:center;font-size:11px;color:${EV2.QUIET};font-family:${EV2.BF}">${xAxis} \u2192</div>` : ''}
                ${yAxis ? `<div data-field="yAxis" style="position:absolute;left:30px;top:80px;width:18px;font-size:11px;color:${EV2.QUIET};font-family:${EV2.BF};writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;height:430px">\u2190 ${yAxis}</div>` : ''}
                <div data-field="quadrants">${cells2}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const gridLeft = 60;
        const gridTop = 80;
        const cellW = 440;
        const cellH = 210;
        const gap = 12;

        let cells = '';
        quads.forEach((q, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const pos = positions[i];
            const cx = gridLeft + pos.x * (cellW + gap);
            const cy = gridTop + pos.y * (cellH + gap);
            const label = esc(q.label) || _placeholder(quadLabels[i]);
            const detail = esc(q.detail);
            cells += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="padding:10px 14px 4px 18px;font-size:13px;font-weight:700;color:${c}">${label}</div>
                ${detail ? `<div style="padding:0 14px 0 18px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            ${xAxis ? `<div data-field="xAxis" style="position:absolute;left:60px;top:530px;width:900px;text-align:center;font-size:11px;color:#888">${xAxis} \u2192</div>` : ''}
            ${yAxis ? `<div data-field="yAxis" style="position:absolute;left:30px;top:80px;width:18px;font-size:11px;color:#888;writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;height:430px">\u2190 ${yAxis}</div>` : ''}
            <div data-field="quadrants">${cells}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  12. METHODS                                                          */
    /* --------------------------------------------------------------------- */
    methods(data, theme, sc) {
        const t = esc(data.title || 'Approach');
        const fields = data.fields && data.fields.length ? data.fields : [
            { label: '', value: '' }, { label: '', value: '' }, { label: '', value: '' }
        ];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const rowH = 48;
            const startY = 120;
            fields.forEach((f, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(f.label) || _placeholder('Label');
                const value = esc(f.value) || _placeholder('Value');
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${rowH}px;background:${bg};border-radius:6px;overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:940px;height:4px;background:${c}"></div>
                    <div style="position:absolute;left:16px;top:14px;font-size:13px;font-weight:700;color:${c};width:180px">${label}</div>
                    <div style="position:absolute;left:210px;top:14px;font-size:12px;color:#444;width:720px">${value}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="fields">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const rowH = 48;
            const startY = 80;
            fields.forEach((f, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(f.label) || _placeholder('Label');
                const value = esc(f.value) || _placeholder('Value');
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px;height:${rowH}px;background:#141414;border-radius:6px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:16px;top:14px;font-size:13px;font-weight:700;color:${c};width:180px">${label}</div>
                    <div style="position:absolute;left:210px;top:14px;font-size:12px;color:#F0F0F0;width:700px">${value}</div>
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="fields">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const rowH = 48;
            const startY = 80;
            fields.forEach((f, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(f.label) || _placeholder('Label');
                const value = esc(f.value) || _placeholder('Value');
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:890px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:16px;top:14px;font-size:13px;font-weight:800;color:${c};width:180px">${label}</div>
                    <div style="position:absolute;left:210px;top:14px;font-size:12px;color:#1A1A1A;width:690px">${value}</div>
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="fields">${rows}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            let rows2 = '';
            const rowH2 = 48;
            const startY2 = 80;
            fields.forEach((f, i) => {
                const color = EV2.ACCENTS[i % 7];
                const y = startY2 + i * (rowH2 + 10);
                const label = esc(f.label) || _placeholder('Label');
                const value = esc(f.value) || _placeholder('Value');
                rows2 += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH2}px;background:${EV2.LIGHT_BOX};overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${color}"></div>
                    <div style="position:absolute;left:16px;top:14px;font-size:13px;font-weight:700;color:${color};font-family:${EV2.TF};width:180px">${label}</div>
                    <div style="position:absolute;left:210px;top:14px;font-size:12px;color:${EV2.CHARCOAL};font-family:${EV2.BF};width:700px">${value}</div>
                </div>`;
            });
            return `${_ev2White(t)}
                <div data-field="fields">${rows2}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        const rowH = 48;
        const startY = 80;
        fields.forEach((f, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 10);
            const label = esc(f.label) || _placeholder('Label');
            const value = esc(f.value) || _placeholder('Value');
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:16px;top:14px;font-size:13px;font-weight:700;color:${c};width:180px">${label}</div>
                <div style="position:absolute;left:210px;top:14px;font-size:12px;color:#444;width:700px">${value}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="fields">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  13. HYPOTHESES                                                       */
    /* --------------------------------------------------------------------- */
    hypotheses(data, theme, sc) {
        const t = esc(data.title || 'Hypotheses');
        const hyps = data.hypotheses && data.hypotheses.length ? data.hypotheses : [
            { text: '', status: '' }, { text: '', status: '' }, { text: '', status: '' }
        ];
        const accent = _accent(sc);

        function badgeHtml(status) {
            if (!status) return '';
            const colors = {
                'Confirmed': { bg: '#E2F0D9', fg: '#368727' },
                'Rejected': { bg: '#FCDEDE', fg: '#C23B22' },
                'Partial': { bg: '#F5ECD4', fg: '#CC7A2E' }
            };
            const style = colors[status] || { bg: '#eee', fg: '#666' };
            return `<span style="display:inline-block;padding:2px 10px;border-radius:10px;font-size:10px;font-weight:600;background:${style.bg};color:${style.fg}">${esc(status)}</span>`;
        }

        if (theme === 'colorful') {
            let rows = '';
            const rowH = 58;
            const startY = 120;
            hyps.forEach((h, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 10);
                const text = esc(h.text) || _placeholder('Hypothesis ' + (i + 1));
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${rowH}px;background:${bg};border-radius:6px;overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:940px;height:4px;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:14px;width:30px;height:30px;border-radius:50%;background:${c};color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center">H${i + 1}</div>
                    <div style="position:absolute;left:56px;top:18px;font-size:12px;color:#333;width:680px">${text}</div>
                    <div style="position:absolute;right:16px;top:16px">${badgeHtml(h.status)}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="hypotheses">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const rowH = 58;
            const startY = 80;
            hyps.forEach((h, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const text = esc(h.text) || _placeholder('Hypothesis ' + (i + 1));
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px;height:${rowH}px;background:#141414;border-radius:6px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:14px;width:30px;height:30px;border-radius:50%;background:${c};color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center">H${i + 1}</div>
                    <div style="position:absolute;left:56px;top:18px;font-size:12px;color:#F0F0F0;width:660px">${text}</div>
                    <div style="position:absolute;right:16px;top:16px">${badgeHtml(h.status)}</div>
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="hypotheses">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const rowH = 58;
            const startY = 80;
            hyps.forEach((h, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const text = esc(h.text) || _placeholder('Hypothesis ' + (i + 1));
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:890px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:14px;width:30px;height:30px;border-radius:50%;background:${c};color:#fff;font-size:11px;font-weight:800;display:flex;align-items:center;justify-content:center">H${i + 1}</div>
                    <div style="position:absolute;left:56px;top:18px;font-size:12px;font-weight:600;color:#1A1A1A;width:650px">${text}</div>
                    <div style="position:absolute;right:16px;top:16px">${badgeHtml(h.status)}</div>
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="hypotheses">${rows}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            let rows2 = '';
            const rowH2 = 58;
            const startY2 = 80;
            hyps.forEach((h, i) => {
                const color = EV2.ACCENTS[i % 7];
                const y = startY2 + i * (rowH2 + 10);
                const text = esc(h.text) || _placeholder('Hypothesis ' + (i + 1));
                rows2 += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH2}px;background:${EV2.LIGHT_BOX};overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${color}"></div>
                    <div style="position:absolute;left:14px;top:14px;width:30px;height:30px;border-radius:50%;background:${color};color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;font-family:${EV2.BF}">H${i + 1}</div>
                    <div style="position:absolute;left:56px;top:18px;font-size:12px;color:${EV2.CHARCOAL};font-family:${EV2.BF};width:660px">${text}</div>
                    <div style="position:absolute;right:16px;top:16px">${badgeHtml(h.status)}</div>
                </div>`;
            });
            return `${_ev2White(t)}
                <div data-field="hypotheses">${rows2}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        const rowH = 58;
        const startY = 80;
        hyps.forEach((h, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 10);
            const text = esc(h.text) || _placeholder('Hypothesis ' + (i + 1));
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:14px;top:14px;width:30px;height:30px;border-radius:50%;background:${c};color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center">H${i + 1}</div>
                <div style="position:absolute;left:56px;top:18px;font-size:12px;color:#333;width:660px">${text}</div>
                <div style="position:absolute;right:16px;top:16px">${badgeHtml(h.status)}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="hypotheses">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  14. WSN_DENSE                                                        */
    /* --------------------------------------------------------------------- */
    wsn_dense(data, theme, sc) {
        const t = esc(data.title || 'Key Finding');
        const what = data.what || {};
        const soWhat = data.soWhat || {};
        const nowWhat = data.nowWhat || {};
        const accent = _accent(sc);
        const labels = data.labels && data.labels.some(l => l) ? data.labels : ['What', 'So What', 'Now What'];

        const cols = [
            { label: labels[0], key: 'what', data: what, color: '#368727' },
            { label: labels[1], key: 'soWhat', data: soWhat, color: '#3880F3' },
            { label: labels[2], key: 'nowWhat', data: nowWhat, color: '#5B2C8F' },
        ];

        if (theme === 'colorful') {
            const cardW = 300;
            const gap = 16;
            let cards = '';
            cols.forEach((col, i) => {
                const x = 30 + i * (cardW + gap);
                const headline = esc(col.data.headline) || _placeholder(col.label + ' headline');
                const detail = esc(col.data.detail);
                const bg = PREVIEW_LIGHTS[i % 5];
                cards += `<div data-field="${col.key}" style="position:absolute;left:${x}px;top:120px;width:${cardW}px;height:400px;background:${bg};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${col.color}"></div>
                    <div style="padding:12px 14px 4px;font-size:14px;font-weight:700;color:${col.color}">${col.label}</div>
                    <div style="padding:4px 14px;font-size:12px;font-weight:600;color:#333">${headline}</div>
                    ${detail ? `<div style="padding:4px 14px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                ${cards}
            </div>`;
        }

        if (theme === 'noir') {
            const cardW = 290;
            const gap = 16;
            let cards = '';
            cols.forEach((col, i) => {
                const x = 45 + i * (cardW + gap);
                const headline = esc(col.data.headline) || _placeholder(col.label + ' headline');
                const detail = esc(col.data.detail);
                cards += `<div data-field="${col.key}" style="position:absolute;left:${x}px;top:80px;width:${cardW}px;height:440px;background:#141414;border-radius:8px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${col.color}"></div>
                    <div style="padding:14px 14px 4px 18px;font-size:14px;font-weight:700;color:${col.color}">${col.label}</div>
                    <div style="padding:4px 14px 0 18px;font-size:12px;font-weight:600;color:#F0F0F0">${headline}</div>
                    ${detail ? `<div style="padding:4px 14px 0 18px;font-size:11px;color:#999;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                ${cards}
            </div>`;
        }

        if (theme === 'bold') {
            const cardW = 275;
            const gap = 16;
            let cards = '';
            cols.forEach((col, i) => {
                const x = 65 + i * (cardW + gap);
                const headline = esc(col.data.headline) || _placeholder(col.label + ' headline');
                const detail = esc(col.data.detail);
                cards += `<div data-field="${col.key}" style="position:absolute;left:${x}px;top:80px;width:${cardW}px;height:440px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.1)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${col.color}"></div>
                    <div style="padding:14px 14px 4px 18px;font-size:14px;font-weight:800;color:${col.color}">${col.label}</div>
                    <div style="padding:4px 14px 0 18px;font-size:12px;font-weight:700;color:#1A1A1A">${headline}</div>
                    ${detail ? `<div style="padding:4px 14px 0 18px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                ${cards}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const ev2CardW = 280;
            const ev2Gap = 16;
            const ev2Colors = [EV2.DK_GREEN, EV2.COBALT, EV2.PURPLE];
            let ev2Cards = '';
            cols.forEach((col, i) => {
                const x = 60 + i * (ev2CardW + ev2Gap);
                const cc = ev2Colors[i % 3];
                const headline = esc(col.data.headline) || _placeholder(col.label + ' headline');
                const detail = esc(col.data.detail);
                ev2Cards += `<div data-field="${col.key}" style="position:absolute;left:${x}px;top:80px;width:${ev2CardW}px;height:440px;background:${EV2.LIGHT_BOX}">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${cc}"></div>
                    <div style="padding:14px 14px 4px 18px;font-size:14px;font-weight:700;color:${cc};font-family:${EV2.TF}">${col.label}</div>
                    <div style="padding:4px 14px 0 18px;font-size:12px;font-weight:600;color:${EV2.CHARCOAL};font-family:${EV2.BF}">${headline}</div>
                    ${detail ? `<div style="padding:4px 14px 0 18px;font-size:11px;color:${EV2.MID};font-family:${EV2.BF};line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `${_ev2White(t)}
                ${ev2Cards}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const cardW = 280;
        const gap = 16;
        let cards = '';
        cols.forEach((col, i) => {
            const x = 60 + i * (cardW + gap);
            const headline = esc(col.data.headline) || _placeholder(col.label + ' headline');
            const detail = esc(col.data.detail);
            cards += `<div data-field="${col.key}" style="position:absolute;left:${x}px;top:80px;width:${cardW}px;height:440px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${col.color}"></div>
                <div style="padding:14px 14px 4px 18px;font-size:14px;font-weight:700;color:${col.color}">${col.label}</div>
                <div style="padding:4px 14px 0 18px;font-size:12px;font-weight:600;color:#333">${headline}</div>
                ${detail ? `<div style="padding:4px 14px 0 18px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            ${cards}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  15. WSN_REVEAL  (preview shows final "all revealed" state)           */
    /* --------------------------------------------------------------------- */
    wsn_reveal(data, theme, sc) {
        const t = esc(data.title || 'Key Finding');
        const accent = _accent(sc);
        const labels = data.labels && data.labels.some(l => l) ? data.labels : ['What', 'So What', 'Now What'];
        const whatData = data.what || {};
        const headline = esc(whatData.headline) || _placeholder('What we found');
        const detail = esc(whatData.detail);
        const c = '#368727';

        function featuredCard(leftX, w, cardBg, titleColor, detailColor, borderStyle) {
            return `<div data-field="what" style="position:absolute;left:${leftX}px;top:140px;width:${w}px;height:330px;background:${cardBg};border-radius:8px;overflow:hidden;${borderStyle}">
                <div style="height:4px;background:${c}"></div>
                <div style="padding:14px 18px 6px;font-size:15px;font-weight:700;color:${c}">${labels[0]}</div>
                <div style="padding:4px 18px;font-size:13px;font-weight:600;color:${titleColor}">${headline}</div>
                ${detail ? `<div style="padding:4px 18px;font-size:12px;color:${detailColor};line-height:1.5">${detail}</div>` : ''}
            </div>`;
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, PREVIEW_LIGHTS[0], '#333', '#666', '')}
            </div>`;
        }
        if (theme === 'noir') {
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, '#141414', '#F0F0F0', '#999', 'border:1px solid #2A2A2A;')}
            </div>`;
        }
        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, '#fff', '#1A1A1A', '#555', 'box-shadow:0 1px 4px rgba(0,0,0,0.1);')}
            </div>`;
        }
        if (theme === 'editorial_v2') {
            return `${_ev2White(t)}
                <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
                ${featuredCard(200, 600, EV2.LIGHT_BOX, EV2.CHARCOAL, EV2.MID, '')}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:0;top:75px;width:1000px;height:55px">${_stepperBar(labels, 1, accent, theme)}</div>
            ${featuredCard(200, 600, '#fff', '#333', '#666', 'box-shadow:0 1px 4px rgba(0,0,0,0.08);')}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  16. FINDINGS_RECS                                                    */
    /* --------------------------------------------------------------------- */
    findings_recs(data, theme, sc) {
        const t = esc(data.title || 'Findings & Recommendations');
        const items = data.items && data.items.length ? data.items : [
            { finding: '', recommendation: '' },
            { finding: '', recommendation: '' },
            { finding: '', recommendation: '' },
        ];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const rowH = 60;
            const startY = 120;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 14);
                const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
                const rec = esc(item.recommendation) || _placeholder('Recommendation ' + (i + 1));
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:420px;height:${rowH}px;background:${bg};border-radius:6px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:8px 12px;font-size:11px;color:#333">${finding}</div>
                </div>
                <div style="position:absolute;left:462px;top:${y + 14}px;width:36px;height:32px;border-radius:50%;background:${c};color:#fff;font-size:16px;display:flex;align-items:center;justify-content:center">\u2192</div>
                <div style="position:absolute;left:510px;top:${y}px;width:460px;height:${rowH}px;background:${bg};border-radius:6px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:8px 12px;font-size:11px;color:#333">${rec}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const rowH = 60;
            const startY = 80;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 14);
                const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
                const rec = esc(item.recommendation) || _placeholder('Recommendation ' + (i + 1));
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:390px;height:${rowH}px;background:#141414;border-radius:6px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="padding:8px 12px 0 14px;font-size:11px;color:#F0F0F0">${finding}</div>
                </div>
                <div style="position:absolute;left:447px;top:${y + 14}px;width:32px;height:32px;border-radius:50%;background:${c};color:#fff;font-size:15px;display:flex;align-items:center;justify-content:center">\u2192</div>
                <div style="position:absolute;left:491px;top:${y}px;width:454px;height:${rowH}px;background:#141414;border-radius:6px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="padding:8px 12px 0 14px;font-size:11px;color:#F0F0F0">${rec}</div>
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const rowH = 60;
            const startY = 80;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 14);
                const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
                const rec = esc(item.recommendation) || _placeholder('Recommendation ' + (i + 1));
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:380px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="padding:8px 12px 0 14px;font-size:11px;color:#1A1A1A">${finding}</div>
                </div>
                <div style="position:absolute;left:457px;top:${y + 14}px;width:32px;height:32px;border-radius:50%;background:${c};color:#fff;font-size:15px;display:flex;align-items:center;justify-content:center">\u2192</div>
                <div style="position:absolute;left:501px;top:${y}px;width:444px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="padding:8px 12px 0 14px;font-size:11px;color:#1A1A1A">${rec}</div>
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            let rows = '';
            items.forEach((item, i) => {
                const y = 95 + i * 70;
                const finding = esc(item.finding) || '';
                const rec = esc(item.recommendation || item.rec) || '';
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:420px;height:58px;background:${EV2.WARM}"><div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${EV2.DK_GREEN}"></div><div style="padding:10px 12px 0 16px;font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.4">${finding}</div></div>
                <div style="position:absolute;left:462px;top:${y + 14}px;font-size:18px;font-weight:700;color:${EV2.GOLD};font-family:${EV2.TF}">\u2192</div>
                <div style="position:absolute;left:500px;top:${y}px;width:470px;height:58px;background:${EV2.COOL}"><div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${EV2.COBALT}"></div><div style="padding:10px 12px 0 16px;font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.4">${rec}</div></div>`;
            });
            return `${_ev2White(t)}
                <div style="position:absolute;left:48px;top:42px;font-size:9px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.BF};letter-spacing:0.5px">FINDING</div>
                <div style="position:absolute;left:518px;top:42px;font-size:9px;font-weight:700;color:${EV2.COBALT};font-family:${EV2.BF};letter-spacing:0.5px">RECOMMENDATION</div>
                <div style="position:absolute;left:30px;top:62px;width:940px;height:1px;background:${EV2.CHARCOAL}"></div>
                ${rows}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        const rowH = 60;
        const startY = 80;
        items.forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 14);
            const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
            const rec = esc(item.recommendation) || _placeholder('Recommendation ' + (i + 1));
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:390px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="padding:8px 12px 0 14px;font-size:11px;color:#333">${finding}</div>
            </div>
            <div style="position:absolute;left:462px;top:${y + 14}px;width:32px;height:32px;border-radius:50%;background:${c};color:#fff;font-size:15px;display:flex;align-items:center;justify-content:center">\u2192</div>
            <div style="position:absolute;left:506px;top:${y}px;width:454px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="padding:8px 12px 0 14px;font-size:11px;color:#333">${rec}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  17. FINDINGS_RECS_DENSE                                              */
    /* --------------------------------------------------------------------- */
    findings_recs_dense(data, theme, sc) {
        const t = esc(data.title || 'Complete Findings');
        const items = data.items && data.items.length ? data.items : [
            { finding: '', recommendation: '' },
            { finding: '', recommendation: '' },
            { finding: '', recommendation: '' },
            { finding: '', recommendation: '' },
        ];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const rowH = 40;
            const startY = 116;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 8);
                const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
                const rec = esc(item.recommendation) || _placeholder('Rec ' + (i + 1));
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:410px;height:${rowH}px;background:${bg};border-radius:4px;overflow:hidden">
                    <div style="height:3px;background:${c}"></div>
                    <div style="padding:6px 10px;font-size:10px;color:#333">${finding}</div>
                </div>
                <div style="position:absolute;left:448px;top:${y + 8}px;font-size:14px;color:${c}">\u2192</div>
                <div style="position:absolute;left:470px;top:${y}px;width:500px;height:${rowH}px;background:${bg};border-radius:4px;overflow:hidden">
                    <div style="height:3px;background:${c}"></div>
                    <div style="padding:6px 10px;font-size:10px;color:#333">${rec}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const rowH = 40;
            const startY = 76;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 8);
                const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
                const rec = esc(item.recommendation) || _placeholder('Rec ' + (i + 1));
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:390px;height:${rowH}px;background:#141414;border-radius:4px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${c}"></div>
                    <div style="padding:6px 10px 0 12px;font-size:10px;color:#F0F0F0">${finding}</div>
                </div>
                <div style="position:absolute;left:443px;top:${y + 8}px;font-size:14px;color:${c}">\u2192</div>
                <div style="position:absolute;left:465px;top:${y}px;width:480px;height:${rowH}px;background:#141414;border-radius:4px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${c}"></div>
                    <div style="padding:6px 10px 0 12px;font-size:10px;color:#F0F0F0">${rec}</div>
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const rowH = 40;
            const startY = 76;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 8);
                const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
                const rec = esc(item.recommendation) || _placeholder('Rec ' + (i + 1));
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:380px;height:${rowH}px;background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 1px 2px rgba(0,0,0,0.07)">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="padding:6px 10px 0 12px;font-size:10px;color:#1A1A1A">${finding}</div>
                </div>
                <div style="position:absolute;left:453px;top:${y + 8}px;font-size:14px;color:${c}">\u2192</div>
                <div style="position:absolute;left:475px;top:${y}px;width:470px;height:${rowH}px;background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 1px 2px rgba(0,0,0,0.07)">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="padding:6px 10px 0 12px;font-size:10px;color:#1A1A1A">${rec}</div>
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            let ev2Rows = '';
            items.forEach((item, i) => {
                const y = 95 + i * 54;
                const finding = esc(item.finding) || '';
                const rec = esc(item.recommendation) || '';
                ev2Rows += `<div style="position:absolute;left:30px;top:${y}px;width:420px;height:42px;background:${EV2.WARM}"><div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${EV2.DK_GREEN}"></div><div style="padding:8px 10px 0 14px;font-size:10px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.3">${finding}</div></div>
                <div style="position:absolute;left:462px;top:${y + 8}px;font-size:16px;font-weight:700;color:${EV2.GOLD};font-family:${EV2.TF}">\u2192</div>
                <div style="position:absolute;left:500px;top:${y}px;width:470px;height:42px;background:${EV2.COOL}"><div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${EV2.COBALT}"></div><div style="padding:8px 10px 0 14px;font-size:10px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.3">${rec}</div></div>`;
            });
            return `${_ev2White(t)}
                <div style="position:absolute;left:48px;top:42px;font-size:9px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.BF};letter-spacing:0.5px">FINDING</div>
                <div style="position:absolute;left:518px;top:42px;font-size:9px;font-weight:700;color:${EV2.COBALT};font-family:${EV2.BF};letter-spacing:0.5px">RECOMMENDATION</div>
                <div style="position:absolute;left:30px;top:62px;width:940px;height:1px;background:${EV2.CHARCOAL}"></div>
                ${ev2Rows}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        const rowH = 40;
        const startY = 76;
        items.forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 8);
            const finding = esc(item.finding) || _placeholder('Finding ' + (i + 1));
            const rec = esc(item.recommendation) || _placeholder('Rec ' + (i + 1));
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:390px;height:${rowH}px;background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 1px 2px rgba(0,0,0,0.05)">
                <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${c}"></div>
                <div style="padding:6px 10px 0 12px;font-size:10px;color:#333">${finding}</div>
            </div>
            <div style="position:absolute;left:458px;top:${y + 8}px;font-size:14px;color:${c}">\u2192</div>
            <div style="position:absolute;left:480px;top:${y}px;width:480px;height:${rowH}px;background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 1px 2px rgba(0,0,0,0.05)">
                <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${c}"></div>
                <div style="padding:6px 10px 0 12px;font-size:10px;color:#333">${rec}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  18. OPEN_QUESTIONS                                                   */
    /* --------------------------------------------------------------------- */
    open_questions(data, theme, sc) {
        const t = esc(data.title || 'Open Questions');
        const questions = data.questions && data.questions.length ? data.questions : ['', '', ''];
        const accent = _accent(sc);

        /* 2x2 grid (max 4 items) */
        const positions = [
            { x: 0, y: 0 }, { x: 1, y: 0 },
            { x: 0, y: 1 }, { x: 1, y: 1 },
        ];

        if (theme === 'colorful') {
            const gridLeft = 30;
            const gridTop = 120;
            const cellW = 460;
            const cellH = 195;
            const gap = 12;

            let cards = '';
            questions.forEach((q, i) => {
                if (i > 3) return;
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const pos = positions[i];
                const cx = gridLeft + pos.x * (cellW + gap);
                const cy = gridTop + pos.y * (cellH + gap);
                const text = esc(q) || _placeholder('Question ' + (i + 1));
                cards += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:${bg};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:14px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:52px;top:14px;width:${cellW - 72}px;font-size:13px;color:#333;line-height:1.5;padding-top:4px">${text}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="questions">${cards}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const gridLeft = 45;
            const gridTop = 80;
            const cellW = 450;
            const cellH = 210;
            const gap = 12;

            let cards = '';
            questions.forEach((q, i) => {
                if (i > 3) return;
                const c = PREVIEW_COLORS[i % 5];
                const pos = positions[i];
                const cx = gridLeft + pos.x * (cellW + gap);
                const cy = gridTop + pos.y * (cellH + gap);
                const text = esc(q) || _placeholder('Question ' + (i + 1));
                cards += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:#141414;border-radius:8px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:14px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:52px;top:14px;width:${cellW - 72}px;font-size:13px;color:#F0F0F0;line-height:1.5;padding-top:4px">${text}</div>
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="questions">${cards}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const gridLeft = 65;
            const gridTop = 80;
            const cellW = 430;
            const cellH = 210;
            const gap = 12;

            let cards = '';
            questions.forEach((q, i) => {
                if (i > 3) return;
                const c = PREVIEW_COLORS[i % 5];
                const pos = positions[i];
                const cx = gridLeft + pos.x * (cellW + gap);
                const cy = gridTop + pos.y * (cellH + gap);
                const text = esc(q) || _placeholder('Question ' + (i + 1));
                cards += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.1)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:14px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:800;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:52px;top:14px;width:${cellW - 72}px;font-size:13px;font-weight:600;color:#1A1A1A;line-height:1.5;padding-top:4px">${text}</div>
                </div>`;
            });
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="questions">${cards}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const ev2Left = 50;
            const ev2Top = 80;
            const ev2CellW = 430;
            const ev2CellH = 210;
            const ev2Gap = 14;
            let ev2Cards = '';
            questions.forEach((q, i) => {
                if (i > 3) return;
                const cc = EV2.ACCENTS[i % 7];
                const pos = positions[i];
                const cx = ev2Left + pos.x * (ev2CellW + ev2Gap);
                const cy = ev2Top + pos.y * (ev2CellH + ev2Gap);
                const text = esc(q) || _placeholder('Question ' + (i + 1));
                ev2Cards += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${ev2CellW}px;height:${ev2CellH}px;background:${EV2.LIGHT_BOX};overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:${ev2CellW}px;height:4px;background:${cc}"></div>
                    <div style="position:absolute;right:10px;top:20px;font-size:80px;font-weight:700;color:${EV2.FAINT};font-family:${EV2.TF};opacity:0.4">?</div>
                    <div style="position:absolute;left:16px;top:18px;width:${ev2CellW - 60}px;font-size:13px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.5">${text}</div>
                </div>`;
            });
            return `${_ev2White(t)}
                <div data-field="questions">${ev2Cards}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const gridLeft = 60;
        const gridTop = 80;
        const cellW = 440;
        const cellH = 210;
        const gap = 12;

        let cards = '';
        questions.forEach((q, i) => {
            if (i > 3) return;
            const c = PREVIEW_COLORS[i % 5];
            const pos = positions[i];
            const cx = gridLeft + pos.x * (cellW + gap);
            const cy = gridTop + pos.y * (cellH + gap);
            const text = esc(q) || _placeholder('Question ' + (i + 1));
            cards += `<div style="position:absolute;left:${cx}px;top:${cy}px;width:${cellW}px;height:${cellH}px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:14px;top:14px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                <div style="position:absolute;left:52px;top:14px;width:${cellW - 72}px;font-size:13px;color:#333;line-height:1.5;padding-top:4px">${text}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="questions">${cards}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  19. PROGRESSIVE_REVEAL                                               */
    /* --------------------------------------------------------------------- */
    progressive_reveal(data, theme, sc) {
        const t = esc(data.title || 'Building the Picture');
        const takeaways = data.takeaways && data.takeaways.length ? data.takeaways : [
            { headline: '', detail: '', summary: '' },
            { headline: '', detail: '', summary: '' },
        ];
        const accent = _accent(sc);

        /* Show the LAST takeaway as the "current" content (final slide state) */
        const lastIdx = takeaways.length - 1;
        const current = takeaways[lastIdx];
        const currentHeadline = esc(current.headline) || _placeholder('Takeaway ' + (lastIdx + 1));
        const currentDetail = esc(current.detail);

        /* Build running takeaway list (all items -- gray, compact) */
        let listHtml = '';
        takeaways.forEach((ta, i) => {
            const summary = esc(ta.summary) || esc(ta.headline) || _placeholder('Takeaway ' + (i + 1));
            listHtml += `<div style="padding:2px 0;font-size:9px;color:#999;font-weight:400">
                <span style="color:${PREVIEW_COLORS[i % 5]};margin-right:4px">\u2022</span>${summary}
            </div>`;
        });

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                <div style="position:absolute;left:0;top:0;width:1000px;height:6px;background:${accent}"></div>
                <div data-field="title" style="position:absolute;left:30px;top:20px;width:940px;font-size:20px;font-weight:700;color:#333">${t}</div>
                <div data-field="takeaways" style="position:absolute;left:30px;top:60px;width:940px;height:220px;background:${PREVIEW_LIGHTS[lastIdx % 5]};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${PREVIEW_COLORS[lastIdx % 5]}"></div>
                    <div style="padding:14px 18px 6px;font-size:15px;font-weight:600;color:#333">${currentHeadline}</div>
                    ${currentDetail ? `<div style="padding:0 18px;font-size:12px;color:#555;line-height:1.5">${currentDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:30px;top:300px;width:940px;height:1px;background:#ddd"></div>
                <div style="position:absolute;left:30px;top:312px;width:940px">${listHtml}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let noirListHtml = '';
            takeaways.forEach((ta, i) => {
                const summary = esc(ta.summary) || esc(ta.headline) || _placeholder('Takeaway ' + (i + 1));
                noirListHtml += `<div style="padding:2px 0;font-size:9px;color:#666;font-weight:400">
                    <span style="color:${PREVIEW_COLORS[i % 5]};margin-right:4px">\u2022</span>${summary}
                </div>`;
            });
            return `${_slideOpen('noir')}
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="takeaways" style="position:absolute;left:45px;top:74px;width:920px;height:220px;background:#141414;border-radius:8px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${PREVIEW_COLORS[lastIdx % 5]}"></div>
                    <div style="padding:14px 18px 6px;font-size:15px;font-weight:600;color:#F0F0F0">${currentHeadline}</div>
                    ${currentDetail ? `<div style="padding:0 18px;font-size:12px;color:#999;line-height:1.5">${currentDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:45px;top:310px;width:920px;height:1px;background:#2A2A2A"></div>
                <div style="position:absolute;left:45px;top:322px;width:920px">${noirListHtml}</div>
            </div>`;
        }

        if (theme === 'bold') {
            return `${_slideOpen('bold')}
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="takeaways" style="position:absolute;left:65px;top:74px;width:880px;height:220px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.1)">
                    <div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${PREVIEW_COLORS[lastIdx % 5]}"></div>
                    <div style="padding:14px 18px 6px;font-size:15px;font-weight:700;color:#1A1A1A">${currentHeadline}</div>
                    ${currentDetail ? `<div style="padding:0 18px;font-size:12px;color:#555;line-height:1.5">${currentDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:65px;top:310px;width:880px;height:1px;background:#ddd"></div>
                <div style="position:absolute;left:65px;top:322px;width:880px">${listHtml}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            return `${_ev2White(t)}
                <div data-field="takeaways" style="position:absolute;left:60px;top:74px;width:880px;height:220px;background:${EV2.LIGHT_BOX}">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${EV2.ACCENTS[lastIdx % 7]}"></div>
                    <div style="padding:14px 18px 6px;font-size:15px;font-weight:600;color:${EV2.CHARCOAL};font-family:${EV2.TF}">${currentHeadline}</div>
                    ${currentDetail ? `<div style="padding:0 18px;font-size:12px;color:${EV2.MID};font-family:${EV2.BF};line-height:1.5">${currentDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:60px;top:310px;width:880px;height:1px;background:${EV2.RULE_CLR}"></div>
                <div style="position:absolute;left:60px;top:322px;width:880px">${listHtml}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="takeaways" style="position:absolute;left:60px;top:74px;width:900px;height:220px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${PREVIEW_COLORS[lastIdx % 5]}"></div>
                <div style="padding:14px 18px 6px;font-size:15px;font-weight:600;color:#333">${currentHeadline}</div>
                ${currentDetail ? `<div style="padding:0 18px;font-size:12px;color:#555;line-height:1.5">${currentDetail}</div>` : ''}
            </div>
            <div style="position:absolute;left:60px;top:310px;width:900px;height:1px;background:#ddd"></div>
            <div style="position:absolute;left:60px;top:322px;width:900px">${listHtml}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  20. TIMELINE                                                         */
    /* --------------------------------------------------------------------- */
    timeline(data, theme, sc) {
        const t = esc(data.title || 'Timeline');
        const milestones = data.milestones && data.milestones.length ? data.milestones : [
            { date: '', title: '', detail: '', status: 'upcoming' },
            { date: '', title: '', detail: '', status: 'current' },
            { date: '', title: '', detail: '', status: 'upcoming' },
        ];
        const accent = _accent(sc);
        const n = milestones.length;

        function statusColor(status, idx) {
            if (status === 'complete') return '#368727';
            if (status === 'current') return PREVIEW_COLORS[idx % 5];
            return '#bbb';
        }

        if (theme === 'colorful') {
            const lineY = 280;
            const startX = 60;
            const endX = 940;
            const span = endX - startX;
            let dots = '';
            milestones.forEach((m, i) => {
                const x = startX + (n === 1 ? span / 2 : i * span / (n - 1));
                const c = PREVIEW_COLORS[i % 5];
                const sc2 = statusColor(m.status, i);
                const filled = m.status === 'complete' || m.status === 'current';
                const dateText = esc(m.date) || '';
                const mTitle = esc(m.title) || _placeholder('Milestone ' + (i + 1));
                const detail = esc(m.detail);
                dots += `<div style="position:absolute;left:${x - 10}px;top:${lineY - 10}px;width:20px;height:20px;border-radius:50%;background:${filled ? sc2 : '#fff'};border:3px solid ${sc2}"></div>`;
                dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 60}px;width:120px;text-align:center;font-size:10px;font-weight:600;color:${c}">${dateText}</div>`;
                dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 42}px;width:120px;text-align:center;font-size:11px;font-weight:600;color:#333">${mTitle}</div>`;
                if (detail) dots += `<div style="position:absolute;left:${x - 70}px;top:${lineY + 18}px;width:140px;text-align:center;font-size:9px;color:#666">${detail}</div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:${startX}px;top:${lineY}px;width:${span}px;height:3px;background:#ddd"></div>
                <div data-field="milestones">${dots}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const lineY = 280;
            const startX = 60;
            const endX = 940;
            const span = endX - startX;
            let dots = '';
            milestones.forEach((m, i) => {
                const x = startX + (n === 1 ? span / 2 : i * span / (n - 1));
                const sc2 = statusColor(m.status, i);
                const filled = m.status === 'complete' || m.status === 'current';
                const dateText = esc(m.date) || '';
                const mTitle = esc(m.title) || _placeholder('Milestone ' + (i + 1));
                const detail = esc(m.detail);
                dots += `<div style="position:absolute;left:${x - 10}px;top:${lineY - 10}px;width:20px;height:20px;border-radius:50%;background:${filled ? sc2 : '#141414'};border:3px solid ${sc2}"></div>`;
                dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 60}px;width:120px;text-align:center;font-size:10px;font-weight:600;color:${accent}">${dateText}</div>`;
                dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 42}px;width:120px;text-align:center;font-size:11px;font-weight:600;color:#F0F0F0">${mTitle}</div>`;
                if (detail) dots += `<div style="position:absolute;left:${x - 70}px;top:${lineY + 18}px;width:140px;text-align:center;font-size:9px;color:#999">${detail}</div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:${startX}px;top:${lineY}px;width:${span}px;height:3px;background:#2A2A2A"></div>
                <div data-field="milestones">${dots}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const lineY = 280;
            const startX = 80;
            const endX = 940;
            const span = endX - startX;
            let dots = '';
            milestones.forEach((m, i) => {
                const x = startX + (n === 1 ? span / 2 : i * span / (n - 1));
                const sc2 = statusColor(m.status, i);
                const filled = m.status === 'complete' || m.status === 'current';
                const dateText = esc(m.date) || '';
                const mTitle = esc(m.title) || _placeholder('Milestone ' + (i + 1));
                const detail = esc(m.detail);
                dots += `<div style="position:absolute;left:${x - 12}px;top:${lineY - 12}px;width:24px;height:24px;border-radius:50%;background:${filled ? sc2 : '#F2F0EC'};border:3px solid ${sc2}"></div>`;
                dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 60}px;width:120px;text-align:center;font-size:10px;font-weight:700;color:${accent}">${dateText}</div>`;
                dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 42}px;width:120px;text-align:center;font-size:11px;font-weight:700;color:#1A1A1A">${mTitle}</div>`;
                if (detail) dots += `<div style="position:absolute;left:${x - 70}px;top:${lineY + 20}px;width:140px;text-align:center;font-size:9px;color:#666">${detail}</div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:${startX}px;top:${lineY}px;width:${span}px;height:4px;background:#ccc"></div>
                <div data-field="milestones">${dots}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const ev2LineY = 280;
            const ev2StartX = 80;
            const ev2EndX = 940;
            const ev2Span = ev2EndX - ev2StartX;
            function ev2StatusColor(status) {
                if (status === 'complete') return EV2.DK_GREEN;
                if (status === 'current') return EV2.GOLD;
                return EV2.FAINT;
            }
            let ev2Dots = '';
            milestones.forEach((m, i) => {
                const x = ev2StartX + (n === 1 ? ev2Span / 2 : i * ev2Span / (n - 1));
                const sc2 = ev2StatusColor(m.status);
                const filled = m.status === 'complete' || m.status === 'current';
                const above = i % 2 === 0;
                const dateText = esc(m.date) || '';
                const mTitle = esc(m.title) || _placeholder('Milestone ' + (i + 1));
                const detail = esc(m.detail);
                ev2Dots += `<div style="position:absolute;left:${x - 8}px;top:${ev2LineY - 8}px;width:16px;height:16px;border-radius:50%;background:${filled ? sc2 : '#fff'};border:3px solid ${sc2}"></div>`;
                if (above) {
                    ev2Dots += `<div style="position:absolute;left:${x - 60}px;top:${ev2LineY - 65}px;width:120px;text-align:center;font-size:9px;font-weight:700;color:${EV2.QUIET};font-family:${EV2.BF}">${dateText}</div>`;
                    ev2Dots += `<div style="position:absolute;left:${x - 60}px;top:${ev2LineY - 48}px;width:120px;text-align:center;font-size:11px;font-weight:600;color:${EV2.CHARCOAL};font-family:${EV2.TF}">${mTitle}</div>`;
                } else {
                    ev2Dots += `<div style="position:absolute;left:${x - 60}px;top:${ev2LineY + 20}px;width:120px;text-align:center;font-size:11px;font-weight:600;color:${EV2.CHARCOAL};font-family:${EV2.TF}">${mTitle}</div>`;
                    ev2Dots += `<div style="position:absolute;left:${x - 60}px;top:${ev2LineY + 40}px;width:120px;text-align:center;font-size:9px;font-weight:700;color:${EV2.QUIET};font-family:${EV2.BF}">${dateText}</div>`;
                }
                if (detail) {
                    const detY = above ? ev2LineY - 82 : ev2LineY + 55;
                    ev2Dots += `<div style="position:absolute;left:${x - 70}px;top:${detY}px;width:140px;text-align:center;font-size:9px;color:${EV2.MID};font-family:${EV2.BF}">${detail}</div>`;
                }
            });
            return `${_ev2White(t)}
                <div style="position:absolute;left:${ev2StartX}px;top:${ev2LineY}px;width:${ev2Span}px;height:3px;background:${EV2.DK_GREEN}"></div>
                <div data-field="milestones">${ev2Dots}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const lineY = 280;
        const startX = 80;
        const endX = 940;
        const span = endX - startX;
        let dots = '';
        milestones.forEach((m, i) => {
            const x = startX + (n === 1 ? span / 2 : i * span / (n - 1));
            const sc2 = statusColor(m.status, i);
            const filled = m.status === 'complete' || m.status === 'current';
            const dateText = esc(m.date) || '';
            const mTitle = esc(m.title) || _placeholder('Milestone ' + (i + 1));
            const detail = esc(m.detail);
            dots += `<div style="position:absolute;left:${x - 10}px;top:${lineY - 10}px;width:20px;height:20px;border-radius:50%;background:${filled ? sc2 : bg};border:3px solid ${sc2}"></div>`;
            dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 60}px;width:120px;text-align:center;font-size:10px;font-weight:600;color:${accent}">${dateText}</div>`;
            dots += `<div style="position:absolute;left:${x - 60}px;top:${lineY - 42}px;width:120px;text-align:center;font-size:11px;font-weight:600;color:#333">${mTitle}</div>`;
            if (detail) dots += `<div style="position:absolute;left:${x - 70}px;top:${lineY + 18}px;width:140px;text-align:center;font-size:9px;color:#666">${detail}</div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:${startX}px;top:${lineY}px;width:${span}px;height:3px;background:#ddd"></div>
            <div data-field="milestones">${dots}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  21. DATA_TABLE                                                       */
    /* --------------------------------------------------------------------- */
    data_table(data, theme, sc) {
        const t = esc(data.title || 'Data Table');
        const headers = data.headers && data.headers.length ? data.headers : ['Column A', 'Column B', 'Column C'];
        const rows = data.rows && data.rows.length ? data.rows : [];
        const highlightCol = data.highlightCol != null ? data.highlightCol : -1;
        const note = esc(data.note || '');
        const accent = _accent(sc);
        const nCols = headers.length;

        function buildTable(headerBg, headerFg, cellBg, cellFg, borderColor, hlBg) {
            const colW = Math.floor(880 / nCols);
            let html = '<div style="position:absolute;left:60px;top:80px;width:880px">';
            html += '<div style="display:flex">';
            headers.forEach((h, ci) => {
                const bg2 = ci === highlightCol ? hlBg : headerBg;
                html += `<div style="width:${colW}px;padding:8px 10px;font-size:11px;font-weight:700;color:${headerFg};background:${bg2};border-bottom:2px solid ${borderColor}">${esc(h)}</div>`;
            });
            html += '</div>';
            rows.forEach((row, ri) => {
                html += '<div style="display:flex">';
                headers.forEach((_, ci) => {
                    const val = esc(row[ci] || row[String(ci)] || '');
                    const bg2 = ci === highlightCol ? hlBg : (ri % 2 === 0 ? cellBg : 'transparent');
                    html += `<div style="width:${colW}px;padding:6px 10px;font-size:11px;color:${cellFg};background:${bg2};border-bottom:1px solid ${borderColor}">${val}</div>`;
                });
                html += '</div>';
            });
            html += '</div>';
            return html;
        }

        if (theme === 'colorful') {
            const table = buildTable(accent, '#fff', PREVIEW_LIGHTS[0], '#333', '#eee', PREVIEW_LIGHTS[1]);
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="rows">${table}</div>
                ${note ? `<div data-field="note" style="position:absolute;left:60px;top:520px;font-size:10px;color:#888">${note}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            const table = buildTable(accent, '#fff', '#141414', '#F0F0F0', '#2A2A2A', '#1A1A1A');
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="rows">${table}</div>
                ${note ? `<div data-field="note" style="position:absolute;left:60px;top:520px;font-size:10px;color:#666">${note}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            const table = buildTable(accent, '#fff', '#fff', '#1A1A1A', '#E0DDD8', '#F5F3EF');
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="rows">${table}</div>
                ${note ? `<div data-field="note" style="position:absolute;left:60px;top:520px;font-size:10px;color:#888">${note}</div>` : ''}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const ev2Table = buildTable(EV2.DK_GREEN, '#fff', EV2.LIGHT_BOX, EV2.CHARCOAL, EV2.RULE_CLR, EV2.WARM);
            return `${_ev2White(t)}
                <div data-field="rows">${ev2Table}</div>
                ${note ? `<div data-field="note" style="position:absolute;left:60px;top:520px;font-size:10px;color:${EV2.QUIET};font-family:${EV2.BF}">${note}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const table = buildTable(accent, '#fff', '#fff', '#444', '#eee', '#F5F5F5');
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="rows">${table}</div>
            ${note ? `<div data-field="note" style="position:absolute;left:60px;top:520px;font-size:10px;color:#999">${note}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  22. MULTI_STAT                                                       */
    /* --------------------------------------------------------------------- */
    multi_stat(data, theme, sc) {
        const t = esc(data.title || 'Key Metrics');
        const stats = data.stats && data.stats.length ? data.stats : [
            { value: '', label: '', detail: '' },
            { value: '', label: '', detail: '' },
            { value: '', label: '', detail: '' },
        ];
        const source = esc(data.source || '');
        const accent = _accent(sc);
        const n = stats.length;

        if (theme === 'colorful') {
            const cardW = Math.floor((940 - (n - 1) * 16) / n);
            let cards = '';
            stats.forEach((s, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const x = 30 + i * (cardW + 16);
                const val = esc(s.value) || _placeholder('--');
                const label = esc(s.label) || _placeholder('Metric');
                const detail = esc(s.detail);
                cards += `<div style="position:absolute;left:${x}px;top:120px;width:${cardW}px;height:340px;background:${bg2};border-radius:8px;overflow:hidden;text-align:center">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:50px 12px 8px;font-size:42px;font-weight:700;color:${c}">${val}</div>
                    <div style="padding:0 12px 4px;font-size:13px;font-weight:600;color:#333">${label}</div>
                    ${detail ? `<div style="padding:0 12px;font-size:10px;color:#666">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="stats">${cards}</div>
                ${source ? `<div data-field="source" style="position:absolute;left:30px;top:520px;width:940px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            const cardW = Math.floor((900 - (n - 1) * 16) / n);
            let cards = '';
            stats.forEach((s, i) => {
                const x = 50 + i * (cardW + 16);
                const val = esc(s.value) || _placeholder('--');
                const label = esc(s.label) || _placeholder('Metric');
                const detail = esc(s.detail);
                cards += `<div style="position:absolute;left:${x}px;top:90px;width:${cardW}px;height:360px;background:#141414;border-radius:8px;text-align:center;border:1px solid #2A2A2A">
                    <div style="padding:60px 12px 8px;font-size:42px;font-weight:700;color:${accent}">${val}</div>
                    <div style="padding:0 12px 4px;font-size:13px;font-weight:600;color:#F0F0F0">${label}</div>
                    ${detail ? `<div style="padding:0 12px;font-size:10px;color:#999">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="stats">${cards}</div>
                ${source ? `<div data-field="source" style="position:absolute;left:50px;top:520px;width:900px;text-align:center;font-size:10px;color:#666">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            const cardW = Math.floor((880 - (n - 1) * 16) / n);
            let cards = '';
            stats.forEach((s, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const x = 65 + i * (cardW + 16);
                const val = esc(s.value) || _placeholder('--');
                const label = esc(s.label) || _placeholder('Metric');
                const detail = esc(s.detail);
                cards += `<div style="position:absolute;left:${x}px;top:90px;width:${cardW}px;height:360px;background:#fff;border-radius:8px;text-align:center;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:${c};border-radius:8px 8px 0 0"></div>
                    <div style="padding:55px 12px 8px;font-size:42px;font-weight:800;color:${c}">${val}</div>
                    <div style="padding:0 12px 4px;font-size:13px;font-weight:700;color:#1A1A1A">${label}</div>
                    ${detail ? `<div style="padding:0 12px;font-size:10px;color:#666">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="stats">${cards}</div>
                ${source ? `<div data-field="source" style="position:absolute;left:65px;top:520px;width:880px;text-align:center;font-size:10px;color:#888">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'editorial_v2') {
            const ev2CardW = Math.floor((900 - (n - 1) * 16) / n);
            let ev2Cards = '';
            stats.forEach((s, i) => {
                const cc = EV2.ACCENTS[i % 7];
                const x = 60 + i * (ev2CardW + 16);
                const val = esc(s.value) || _placeholder('--');
                const label = esc(s.label) || _placeholder('Metric');
                const detail = esc(s.detail);
                ev2Cards += `<div style="position:absolute;left:${x}px;top:90px;width:${ev2CardW}px;height:360px;background:${EV2.LIGHT_BOX};text-align:center">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${cc}"></div>
                    <div style="padding:55px 12px 8px;font-size:42px;font-weight:700;color:${cc};font-family:${EV2.TF}">${val}</div>
                    <div style="padding:0 12px 4px;font-size:13px;font-weight:600;color:${EV2.CHARCOAL};font-family:${EV2.BF}">${label}</div>
                    ${detail ? `<div style="padding:0 12px;font-size:10px;color:${EV2.MID};font-family:${EV2.BF}">${detail}</div>` : ''}
                </div>`;
            });
            return `${_ev2White(t)}
                <div data-field="stats">${ev2Cards}</div>
                ${source ? `<div data-field="source" style="position:absolute;left:60px;top:520px;width:900px;text-align:center;font-size:10px;color:${EV2.QUIET};font-family:${EV2.BF}">Source: ${source}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const cardW = Math.floor((900 - (n - 1) * 16) / n);
        let cards = '';
        stats.forEach((s, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const x = 60 + i * (cardW + 16);
            const val = esc(s.value) || _placeholder('--');
            const label = esc(s.label) || _placeholder('Metric');
            const detail = esc(s.detail);
            cards += `<div style="position:absolute;left:${x}px;top:90px;width:${cardW}px;height:360px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08);text-align:center">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="padding:55px 12px 8px;font-size:42px;font-weight:700;color:${c}">${val}</div>
                <div style="padding:0 12px 4px;font-size:13px;font-weight:600;color:#333">${label}</div>
                ${detail ? `<div style="padding:0 12px;font-size:10px;color:#666">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="stats">${cards}</div>
            ${source ? `<div data-field="source" style="position:absolute;left:60px;top:520px;width:900px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  23. PERSONA                                                          */
    /* --------------------------------------------------------------------- */
    persona(data, theme, sc) {
        const t = esc(data.title || 'Persona');
        const name = esc(data.name || 'User Name');
        const archetype = esc(data.archetype || '');
        const traits = data.traits && data.traits.length ? data.traits : [];
        const strategy = esc(data.strategy || '');
        const detail = esc(data.detail || '');
        const accent = _accent(sc);

        function traitPills(pillBg, pillFg) {
            return traits.map(tr => `<span style="display:inline-block;padding:3px 10px;margin:2px 4px 2px 0;border-radius:12px;font-size:10px;background:${pillBg};color:${pillFg}">${esc(tr)}</span>`).join('');
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:60px;top:120px;width:880px;background:${PREVIEW_LIGHTS[0]};border-radius:10px;overflow:hidden;padding-bottom:20px">
                    <div style="height:4px;background:${accent}"></div>
                    <div style="padding:18px 24px 4px">
                        <span data-field="name" style="font-size:22px;font-weight:700;color:#333">${name}</span>
                        ${archetype ? `<span data-field="archetype" style="display:inline-block;margin-left:12px;padding:3px 12px;border-radius:12px;font-size:11px;font-weight:600;background:${accent};color:#fff">${archetype}</span>` : ''}
                    </div>
                    ${traits.length ? `<div data-field="traits" style="padding:8px 24px">${traitPills(PREVIEW_LIGHTS[1], '#3880F3')}</div>` : ''}
                    ${strategy ? `<div style="padding:4px 24px"><span style="font-size:11px;font-weight:600;color:${accent}">Strategy:</span> <span data-field="strategy" style="font-size:11px;color:#444">${strategy}</span></div>` : ''}
                    ${detail ? `<div data-field="detail" style="padding:4px 24px;font-size:11px;color:#666;line-height:1.5">${detail}</div>` : ''}
                </div>
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:50px;top:90px;width:900px;background:#141414;border-radius:10px;overflow:hidden;padding-bottom:20px;border:1px solid #2A2A2A">
                    <div style="padding:18px 24px 4px">
                        <span data-field="name" style="font-size:22px;font-weight:700;color:#F0F0F0">${name}</span>
                        ${archetype ? `<span data-field="archetype" style="display:inline-block;margin-left:12px;padding:3px 12px;border-radius:12px;font-size:11px;font-weight:600;background:${accent};color:#fff">${archetype}</span>` : ''}
                    </div>
                    ${traits.length ? `<div data-field="traits" style="padding:8px 24px">${traitPills('#2A2A2A', '#F0F0F0')}</div>` : ''}
                    ${strategy ? `<div style="padding:4px 24px"><span style="font-size:11px;font-weight:600;color:${accent}">Strategy:</span> <span data-field="strategy" style="font-size:11px;color:#ccc">${strategy}</span></div>` : ''}
                    ${detail ? `<div data-field="detail" style="padding:4px 24px;font-size:11px;color:#999;line-height:1.5">${detail}</div>` : ''}
                </div>
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:65px;top:90px;width:880px;background:#fff;border-radius:10px;overflow:hidden;padding-bottom:20px;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:${accent}"></div>
                    <div style="padding:18px 24px 4px">
                        <span data-field="name" style="font-size:22px;font-weight:800;color:#1A1A1A">${name}</span>
                        ${archetype ? `<span data-field="archetype" style="display:inline-block;margin-left:12px;padding:3px 12px;border-radius:12px;font-size:11px;font-weight:700;background:${accent};color:#fff">${archetype}</span>` : ''}
                    </div>
                    ${traits.length ? `<div data-field="traits" style="padding:8px 24px">${traitPills('#F2F0EC', '#1A1A1A')}</div>` : ''}
                    ${strategy ? `<div style="padding:4px 24px"><span style="font-size:11px;font-weight:700;color:${accent}">Strategy:</span> <span data-field="strategy" style="font-size:11px;color:#333">${strategy}</span></div>` : ''}
                    ${detail ? `<div data-field="detail" style="padding:4px 24px;font-size:11px;color:#666;line-height:1.5">${detail}</div>` : ''}
                </div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            return `${_ev2White(t)}
                <div style="position:absolute;left:30px;top:50px;width:300px;height:460px;background:${EV2.DK_GREEN}">
                    <div style="padding:30px 24px 8px">
                        <div data-field="name" style="font-size:22px;font-weight:700;color:#fff;font-family:${EV2.TF}">${name}</div>
                        ${archetype ? `<div data-field="archetype" style="margin-top:8px;font-size:11px;font-weight:600;color:${EV2.GOLD};font-family:${EV2.BF};text-transform:uppercase;letter-spacing:0.5px">${archetype}</div>` : ''}
                    </div>
                    ${traits.length ? `<div data-field="traits" style="padding:12px 24px 0">${traitPills('rgba(255,255,255,0.15)', EV2.CREAM)}</div>` : ''}
                </div>
                <div style="position:absolute;left:350px;top:50px;width:620px;height:260px;background:${EV2.LIGHT_BOX}">
                    <div style="padding:18px 20px 4px">
                        <div style="font-size:9px;font-weight:700;color:${EV2.QUIET};font-family:${EV2.BF};text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px">DETAIL</div>
                        ${detail ? `<div data-field="detail" style="font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.5">${detail}</div>` : ''}
                    </div>
                </div>
                ${strategy ? `<div style="position:absolute;left:350px;top:324px;width:620px;height:120px;background:${EV2.WARM}">
                    <div style="padding:14px 20px">
                        <div style="font-size:9px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.BF};text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px">STRATEGY</div>
                        <div data-field="strategy" style="font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.5">${strategy}</div>
                    </div>
                </div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:60px;top:90px;width:900px;background:#fff;border-radius:10px;overflow:hidden;padding-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${accent}"></div>
                <div style="padding:18px 24px 4px 20px">
                    <span data-field="name" style="font-size:22px;font-weight:700;color:#333">${name}</span>
                    ${archetype ? `<span data-field="archetype" style="display:inline-block;margin-left:12px;padding:3px 12px;border-radius:12px;font-size:11px;font-weight:600;background:${accent};color:#fff">${archetype}</span>` : ''}
                </div>
                ${traits.length ? `<div data-field="traits" style="padding:8px 24px 0 20px">${traitPills('#F0F0F0', '#555')}</div>` : ''}
                ${strategy ? `<div style="padding:4px 24px 0 20px"><span style="font-size:11px;font-weight:600;color:${accent}">Strategy:</span> <span data-field="strategy" style="font-size:11px;color:#444">${strategy}</span></div>` : ''}
                ${detail ? `<div data-field="detail" style="padding:4px 24px 0 20px;font-size:11px;color:#666;line-height:1.5">${detail}</div>` : ''}
            </div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  24. RISK_TRADEOFF                                                    */
    /* --------------------------------------------------------------------- */
    risk_tradeoff(data, theme, sc) {
        const t = esc(data.title || 'Risks & Rewards');
        const risks = data.risks && data.risks.length ? data.risks : [{ label: '', detail: '', severity: '' }];
        const rewards = data.rewards && data.rewards.length ? data.rewards : [{ label: '', detail: '' }];
        const accent = _accent(sc);

        function severityBadge(sev, bgMap) {
            if (!sev) return '';
            const colors = bgMap || { high: { bg: '#FCDEDE', fg: '#C23B22' }, medium: { bg: '#F5ECD4', fg: '#CC7A2E' }, low: { bg: '#E2F0D9', fg: '#368727' } };
            const s = colors[sev.toLowerCase()] || { bg: '#eee', fg: '#666' };
            return `<span style="display:inline-block;padding:2px 8px;border-radius:8px;font-size:9px;font-weight:600;background:${s.bg};color:${s.fg}">${esc(sev)}</span>`;
        }

        if (theme === 'colorful') {
            let riskHtml = '';
            risks.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Risk ' + (i + 1));
                const detail = esc(r.detail);
                riskHtml += `<div style="padding:8px 14px;border-bottom:1px solid rgba(194,59,34,0.1)">
                    <div style="font-size:12px;font-weight:600;color:#C23B22">${label} ${severityBadge(r.severity)}</div>
                    ${detail ? `<div style="font-size:10px;color:#666;margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            let rewardHtml = '';
            rewards.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Reward ' + (i + 1));
                const detail = esc(r.detail);
                rewardHtml += `<div style="padding:8px 14px;border-bottom:1px solid rgba(54,135,39,0.1)">
                    <div style="font-size:12px;font-weight:600;color:#368727">${label}</div>
                    ${detail ? `<div style="font-size:10px;color:#666;margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="risks" style="position:absolute;left:30px;top:120px;width:450px;height:400px;background:#FFF0EE;border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:#C23B22"></div>
                    <div style="padding:10px 14px;font-size:13px;font-weight:700;color:#C23B22;border-bottom:2px solid rgba(194,59,34,0.2)">Risks</div>
                    ${riskHtml}
                </div>
                <div data-field="rewards" style="position:absolute;left:500px;top:120px;width:470px;height:400px;background:#E2F0D9;border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:#368727"></div>
                    <div style="padding:10px 14px;font-size:13px;font-weight:700;color:#368727;border-bottom:2px solid rgba(54,135,39,0.2)">Rewards</div>
                    ${rewardHtml}
                </div>
            </div>`;
        }

        if (theme === 'noir') {
            let riskHtml = '';
            risks.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Risk ' + (i + 1));
                const detail = esc(r.detail);
                const noirSevMap = { high: { bg: '#3A1515', fg: '#FF6B6B' }, medium: { bg: '#3A2A15', fg: '#FFB347' }, low: { bg: '#153A15', fg: '#6BCB77' } };
                riskHtml += `<div style="padding:8px 14px;border-bottom:1px solid #2A2A2A">
                    <div style="font-size:12px;font-weight:600;color:#FF6B6B">${label} ${severityBadge(r.severity, noirSevMap)}</div>
                    ${detail ? `<div style="font-size:10px;color:#999;margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            let rewardHtml = '';
            rewards.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Reward ' + (i + 1));
                const detail = esc(r.detail);
                rewardHtml += `<div style="padding:8px 14px;border-bottom:1px solid #2A2A2A">
                    <div style="font-size:12px;font-weight:600;color:#6BCB77">${label}</div>
                    ${detail ? `<div style="font-size:10px;color:#999;margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="risks" style="position:absolute;left:30px;top:90px;width:450px;height:430px;background:#141414;border-radius:8px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="padding:10px 14px;font-size:13px;font-weight:700;color:#FF6B6B;border-bottom:1px solid #2A2A2A">Risks</div>
                    ${riskHtml}
                </div>
                <div data-field="rewards" style="position:absolute;left:500px;top:90px;width:470px;height:430px;background:#141414;border-radius:8px;overflow:hidden;border:1px solid #2A2A2A">
                    <div style="padding:10px 14px;font-size:13px;font-weight:700;color:#6BCB77;border-bottom:1px solid #2A2A2A">Rewards</div>
                    ${rewardHtml}
                </div>
            </div>`;
        }

        if (theme === 'bold') {
            let riskHtml = '';
            risks.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Risk ' + (i + 1));
                const detail = esc(r.detail);
                riskHtml += `<div style="padding:8px 14px;border-bottom:1px solid #E0DDD8">
                    <div style="font-size:12px;font-weight:700;color:#C23B22">${label} ${severityBadge(r.severity)}</div>
                    ${detail ? `<div style="font-size:10px;color:#666;margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            let rewardHtml = '';
            rewards.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Reward ' + (i + 1));
                const detail = esc(r.detail);
                rewardHtml += `<div style="padding:8px 14px;border-bottom:1px solid #E0DDD8">
                    <div style="font-size:12px;font-weight:700;color:#368727">${label}</div>
                    ${detail ? `<div style="font-size:10px;color:#666;margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="risks" style="position:absolute;left:45px;top:90px;width:440px;height:430px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:#C23B22;border-radius:8px 8px 0 0"></div>
                    <div style="padding:10px 14px;font-size:13px;font-weight:800;color:#C23B22;border-bottom:2px solid #E0DDD8">Risks</div>
                    ${riskHtml}
                </div>
                <div data-field="rewards" style="position:absolute;left:505px;top:90px;width:450px;height:430px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:#368727;border-radius:8px 8px 0 0"></div>
                    <div style="padding:10px 14px;font-size:13px;font-weight:800;color:#368727;border-bottom:2px solid #E0DDD8">Rewards</div>
                    ${rewardHtml}
                </div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            let ev2RiskHtml = '';
            risks.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Risk ' + (i + 1));
                const detail = esc(r.detail);
                ev2RiskHtml += `<div style="padding:8px 14px;border-bottom:1px solid ${EV2.RULE_CLR}">
                    <div style="font-size:12px;font-weight:600;color:#C23B22;font-family:${EV2.BF}">${label} ${severityBadge(r.severity)}</div>
                    ${detail ? `<div style="font-size:10px;color:${EV2.MID};font-family:${EV2.BF};margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            let ev2RewardHtml = '';
            rewards.forEach((r, i) => {
                const label = esc(r.label) || _placeholder('Reward ' + (i + 1));
                const detail = esc(r.detail);
                ev2RewardHtml += `<div style="padding:8px 14px;border-bottom:1px solid ${EV2.RULE_CLR}">
                    <div style="font-size:12px;font-weight:600;color:${EV2.DK_GREEN};font-family:${EV2.BF}">${label}</div>
                    ${detail ? `<div style="font-size:10px;color:${EV2.MID};font-family:${EV2.BF};margin-top:2px">${detail}</div>` : ''}
                </div>`;
            });
            return `${_ev2White(t)}
                <div data-field="risks" style="position:absolute;left:30px;top:80px;width:450px;height:430px;background:${EV2.LIGHT_BOX}">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:#C23B22"></div>
                    <div style="padding:10px 14px 0 18px;font-size:13px;font-weight:700;color:#C23B22;font-family:${EV2.TF};border-bottom:2px solid ${EV2.RULE_CLR}">Risks</div>
                    ${ev2RiskHtml}
                </div>
                <div data-field="rewards" style="position:absolute;left:500px;top:80px;width:470px;height:430px;background:${EV2.LIGHT_BOX}">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${EV2.DK_GREEN}"></div>
                    <div style="padding:10px 14px 0 18px;font-size:13px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.TF};border-bottom:2px solid ${EV2.RULE_CLR}">Rewards</div>
                    ${ev2RewardHtml}
                </div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let riskHtml = '';
        risks.forEach((r, i) => {
            const label = esc(r.label) || _placeholder('Risk ' + (i + 1));
            const detail = esc(r.detail);
            riskHtml += `<div style="padding:8px 14px;border-bottom:1px solid rgba(0,0,0,0.06)">
                <div style="font-size:12px;font-weight:600;color:#C23B22">${label} ${severityBadge(r.severity)}</div>
                ${detail ? `<div style="font-size:10px;color:#666;margin-top:2px">${detail}</div>` : ''}
            </div>`;
        });
        let rewardHtml = '';
        rewards.forEach((r, i) => {
            const label = esc(r.label) || _placeholder('Reward ' + (i + 1));
            const detail = esc(r.detail);
            rewardHtml += `<div style="padding:8px 14px;border-bottom:1px solid rgba(0,0,0,0.06)">
                <div style="font-size:12px;font-weight:600;color:#368727">${label}</div>
                ${detail ? `<div style="font-size:10px;color:#666;margin-top:2px">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="risks" style="position:absolute;left:60px;top:90px;width:420px;height:430px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:#C23B22"></div>
                <div style="padding:10px 14px 0 18px;font-size:13px;font-weight:700;color:#C23B22;border-bottom:2px solid rgba(0,0,0,0.06)">Risks</div>
                ${riskHtml}
            </div>
            <div data-field="rewards" style="position:absolute;left:500px;top:90px;width:460px;height:430px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:#368727"></div>
                <div style="padding:10px 14px 0 18px;font-size:13px;font-weight:700;color:#368727;border-bottom:2px solid rgba(0,0,0,0.06)">Rewards</div>
                ${rewardHtml}
            </div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  25. APPENDIX                                                         */
    /* --------------------------------------------------------------------- */
    appendix(data, theme, sc) {
        const t = esc(data.title || 'Appendix');
        const sections = data.sections && data.sections.length ? data.sections : [{ label: '', content: '' }, { label: '', content: '' }];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const startY = 120;
            const rowH = 56;
            sections.forEach((s, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(s.label) || _placeholder('Section ' + (i + 1));
                const content = esc(s.content);
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;min-height:${rowH}px;background:${bg2};border-radius:6px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:8px 14px 2px;font-size:12px;font-weight:700;color:${c}">${label}</div>
                    ${content ? `<div style="padding:2px 14px 8px;font-size:11px;color:#444;line-height:1.4">${content}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="sections">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const startY = 90;
            const rowH = 56;
            sections.forEach((s, i) => {
                const y = startY + i * (rowH + 10);
                const label = esc(s.label) || _placeholder('Section ' + (i + 1));
                const content = esc(s.content);
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px;min-height:${rowH}px;background:#141414;border-radius:6px;border:1px solid #2A2A2A">
                    <div style="padding:8px 14px 2px;font-size:12px;font-weight:700;color:${accent}">${label}</div>
                    ${content ? `<div style="padding:2px 14px 8px;font-size:11px;color:#999;line-height:1.4">${content}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="sections">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const startY = 90;
            const rowH = 56;
            sections.forEach((s, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(s.label) || _placeholder('Section ' + (i + 1));
                const content = esc(s.content);
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:880px;min-height:${rowH}px;background:#fff;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:${c};border-radius:6px 6px 0 0"></div>
                    <div style="padding:8px 14px 2px;font-size:12px;font-weight:800;color:${c}">${label}</div>
                    ${content ? `<div style="padding:2px 14px 8px;font-size:11px;color:#444;line-height:1.4">${content}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="sections">${rows}</div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            let ev2Rows = '';
            const ev2StartY = 80;
            const ev2RowH = 56;
            sections.forEach((s, i) => {
                const y = ev2StartY + i * (ev2RowH + 10);
                const label = esc(s.label) || _placeholder('Section ' + (i + 1));
                const content = esc(s.content);
                ev2Rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;min-height:${ev2RowH}px;display:flex">
                    <div style="width:160px;padding:8px 14px 2px 0;font-size:11px;font-weight:700;color:${EV2.QUIET};font-family:${EV2.BF};text-transform:uppercase;letter-spacing:0.3px">${label}</div>
                    <div style="flex:1;padding:8px 14px 2px;border-left:1px solid ${EV2.RULE_CLR}">
                        ${content ? `<div style="font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.4">${content}</div>` : ''}
                    </div>
                </div>`;
            });
            return `${_ev2White(t)}
                <div data-field="sections">${ev2Rows}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        const startY = 80;
        const rowH = 56;
        sections.forEach((s, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 10);
            const label = esc(s.label) || _placeholder('Section ' + (i + 1));
            const content = esc(s.content);
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;min-height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="padding:8px 14px 2px 18px;font-size:12px;font-weight:700;color:${c}">${label}</div>
                ${content ? `<div style="padding:2px 14px 8px 18px;font-size:11px;color:#444;line-height:1.4">${content}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="sections">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  26. BEFORE_AFTER                                                     */
    /* --------------------------------------------------------------------- */
    before_after(data, theme, sc) {
        const t = esc(data.title || 'Before & After');
        const before = data.before || {};
        const after = data.after || {};
        const intervention = esc(data.intervention || '');
        const bLabel = esc(before.label || 'Before');
        const bDetail = esc(before.detail || '');
        const aLabel = esc(after.label || 'After');
        const aDetail = esc(after.detail || '');
        const accent = _accent(sc);

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="before" style="position:absolute;left:30px;top:120px;width:300px;height:390px;background:#FFF0EE;border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:#C23B22"></div>
                    <div style="padding:14px 16px 6px;font-size:14px;font-weight:700;color:#C23B22">${bLabel}</div>
                    ${bDetail ? `<div style="padding:0 16px;font-size:11px;color:#555;line-height:1.5">${bDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:345px;top:120px;width:310px;height:390px;background:${PREVIEW_LIGHTS[1]};border-radius:8px;overflow:hidden;text-align:center">
                    <div style="height:4px;background:${accent}"></div>
                    <div style="padding:140px 16px 8px;font-size:36px;color:${accent}">\u2192</div>
                    ${intervention ? `<div data-field="intervention" style="padding:0 16px;font-size:13px;font-weight:600;color:#333">${intervention}</div>` : ''}
                </div>
                <div data-field="after" style="position:absolute;left:670px;top:120px;width:300px;height:390px;background:#E2F0D9;border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:#368727"></div>
                    <div style="padding:14px 16px 6px;font-size:14px;font-weight:700;color:#368727">${aLabel}</div>
                    ${aDetail ? `<div style="padding:0 16px;font-size:11px;color:#555;line-height:1.5">${aDetail}</div>` : ''}
                </div>
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="before" style="position:absolute;left:30px;top:90px;width:300px;height:420px;background:#141414;border-radius:8px;border:1px solid #2A2A2A">
                    <div style="padding:14px 16px 6px;font-size:14px;font-weight:700;color:#FF6B6B">${bLabel}</div>
                    ${bDetail ? `<div style="padding:0 16px;font-size:11px;color:#999;line-height:1.5">${bDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:345px;top:90px;width:310px;height:420px;background:#141414;border-radius:8px;border:1px solid #2A2A2A;text-align:center">
                    <div style="padding:150px 16px 8px;font-size:36px;color:${accent}">\u2192</div>
                    ${intervention ? `<div data-field="intervention" style="padding:0 16px;font-size:13px;font-weight:600;color:#F0F0F0">${intervention}</div>` : ''}
                </div>
                <div data-field="after" style="position:absolute;left:670px;top:90px;width:300px;height:420px;background:#141414;border-radius:8px;border:1px solid #2A2A2A">
                    <div style="padding:14px 16px 6px;font-size:14px;font-weight:700;color:#6BCB77">${aLabel}</div>
                    ${aDetail ? `<div style="padding:0 16px;font-size:11px;color:#999;line-height:1.5">${aDetail}</div>` : ''}
                </div>
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="before" style="position:absolute;left:45px;top:90px;width:290px;height:420px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:#C23B22;border-radius:8px 8px 0 0"></div>
                    <div style="padding:14px 16px 6px;font-size:14px;font-weight:800;color:#C23B22">${bLabel}</div>
                    ${bDetail ? `<div style="padding:0 16px;font-size:11px;color:#444;line-height:1.5">${bDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:350px;top:90px;width:300px;height:420px;background:#fff;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.08);text-align:center;overflow:hidden">
                    <div style="height:5px;background:${accent};border-radius:8px 8px 0 0"></div>
                    <div style="padding:150px 16px 8px;font-size:36px;color:${accent}">\u2192</div>
                    ${intervention ? `<div data-field="intervention" style="padding:0 16px;font-size:13px;font-weight:700;color:#1A1A1A">${intervention}</div>` : ''}
                </div>
                <div data-field="after" style="position:absolute;left:665px;top:90px;width:290px;height:420px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:#368727;border-radius:8px 8px 0 0"></div>
                    <div style="padding:14px 16px 6px;font-size:14px;font-weight:800;color:#368727">${aLabel}</div>
                    ${aDetail ? `<div style="padding:0 16px;font-size:11px;color:#444;line-height:1.5">${aDetail}</div>` : ''}
                </div>
            </div>`;
        }

        if (theme === 'editorial_v2') {
            return `${_ev2White(t)}
                <div data-field="before" style="position:absolute;left:30px;top:80px;width:290px;height:420px;background:${EV2.WARM}">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:#C23B22"></div>
                    <div style="padding:14px 16px 6px 18px;font-size:14px;font-weight:700;color:#C23B22;font-family:${EV2.TF}">${bLabel}</div>
                    ${bDetail ? `<div style="padding:0 16px 0 18px;font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.5">${bDetail}</div>` : ''}
                </div>
                <div style="position:absolute;left:335px;top:80px;width:280px;height:420px;background:${EV2.LIGHT_BOX};text-align:center">
                    <div style="padding:150px 16px 8px;font-size:36px;font-weight:700;color:${EV2.GOLD};font-family:${EV2.TF}">\u2192</div>
                    ${intervention ? `<div data-field="intervention" style="padding:0 16px;font-size:13px;font-weight:600;color:${EV2.CHARCOAL};font-family:${EV2.BF}">${intervention}</div>` : ''}
                </div>
                <div data-field="after" style="position:absolute;left:630px;top:80px;width:340px;height:420px;background:${EV2.COOL}">
                    <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${EV2.DK_GREEN}"></div>
                    <div style="padding:14px 16px 6px 18px;font-size:14px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.TF}">${aLabel}</div>
                    ${aDetail ? `<div style="padding:0 16px 0 18px;font-size:11px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.5">${aDetail}</div>` : ''}
                </div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="before" style="position:absolute;left:60px;top:90px;width:280px;height:420px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:#C23B22"></div>
                <div style="padding:14px 16px 6px 18px;font-size:14px;font-weight:700;color:#C23B22">${bLabel}</div>
                ${bDetail ? `<div style="padding:0 16px 0 18px;font-size:11px;color:#555;line-height:1.5">${bDetail}</div>` : ''}
            </div>
            <div style="position:absolute;left:355px;top:90px;width:290px;height:420px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08);text-align:center">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${accent}"></div>
                <div style="padding:150px 16px 8px;font-size:36px;color:${accent}">\u2192</div>
                ${intervention ? `<div data-field="intervention" style="padding:0 16px;font-size:13px;font-weight:600;color:#333">${intervention}</div>` : ''}
            </div>
            <div data-field="after" style="position:absolute;left:660px;top:90px;width:300px;height:420px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:#368727"></div>
                <div style="padding:14px 16px 6px 18px;font-size:14px;font-weight:700;color:#368727">${aLabel}</div>
                ${aDetail ? `<div style="padding:0 16px 0 18px;font-size:11px;color:#555;line-height:1.5">${aDetail}</div>` : ''}
            </div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  27. SUMMARY                                                          */
    /* --------------------------------------------------------------------- */
    summary(data, theme, sc) {
        const t = esc(data.title || 'Summary');
        const sections = data.sections && data.sections.length ? data.sections : [
            { heading: '', points: ['', ''] },
            { heading: '', points: ['', ''] },
        ];
        const accent = _accent(sc);
        const n = sections.length;

        function buildCols(headingColor, textColor, cardBg, borderColor, leftStripe) {
            const colW = Math.floor((880 - (n - 1) * 16) / n);
            let cols = '';
            sections.forEach((sec, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const x = 60 + i * (colW + 16);
                const heading = esc(sec.heading) || _placeholder('Section ' + (i + 1));
                let pointsHtml = '';
                const pts = sec.points && sec.points.length ? sec.points : [''];
                pts.forEach((p, pi) => {
                    const ptText = esc(p) || _placeholder('Point ' + (pi + 1));
                    pointsHtml += `<div style="padding:4px 0;font-size:11px;color:${textColor};border-bottom:1px solid ${borderColor}">\u2022 ${ptText}</div>`;
                });
                cols += `<div style="position:absolute;left:${x}px;top:90px;width:${colW}px;height:420px;background:${cardBg};border-radius:8px;overflow:hidden;${leftStripe ? `box-shadow:0 1px 4px rgba(0,0,0,0.08)` : ''}">
                    ${leftStripe ? `<div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>` : `<div style="height:4px;background:${c}"></div>`}
                    <div style="padding:12px 14px 6px ${leftStripe ? '18px' : '14px'};font-size:13px;font-weight:700;color:${headingColor === 'auto' ? c : headingColor}">${heading}</div>
                    <div style="padding:0 14px 0 ${leftStripe ? '18px' : '14px'}">${pointsHtml}</div>
                </div>`;
            });
            return cols;
        }

        if (theme === 'colorful') {
            const colW = Math.floor((940 - (n - 1) * 16) / n);
            let cols = '';
            sections.forEach((sec, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const x = 30 + i * (colW + 16);
                const heading = esc(sec.heading) || _placeholder('Section ' + (i + 1));
                let pointsHtml = '';
                const pts = sec.points && sec.points.length ? sec.points : [''];
                pts.forEach((p, pi) => {
                    const ptText = esc(p) || _placeholder('Point ' + (pi + 1));
                    pointsHtml += `<div style="padding:4px 0;font-size:11px;color:#444;border-bottom:1px solid rgba(0,0,0,0.06)">\u2022 ${ptText}</div>`;
                });
                cols += `<div style="position:absolute;left:${x}px;top:120px;width:${colW}px;height:400px;background:${bg2};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:12px 14px 6px;font-size:13px;font-weight:700;color:${c}">${heading}</div>
                    <div style="padding:0 14px">${pointsHtml}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="sections">${cols}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const colW = Math.floor((900 - (n - 1) * 16) / n);
            let cols = '';
            sections.forEach((sec, i) => {
                const x = 50 + i * (colW + 16);
                const heading = esc(sec.heading) || _placeholder('Section ' + (i + 1));
                let pointsHtml = '';
                const pts = sec.points && sec.points.length ? sec.points : [''];
                pts.forEach((p, pi) => {
                    const ptText = esc(p) || _placeholder('Point ' + (pi + 1));
                    pointsHtml += `<div style="padding:4px 0;font-size:11px;color:#999;border-bottom:1px solid #2A2A2A">\u2022 ${ptText}</div>`;
                });
                cols += `<div style="position:absolute;left:${x}px;top:90px;width:${colW}px;height:420px;background:#141414;border-radius:8px;border:1px solid #2A2A2A">
                    <div style="padding:12px 14px 6px;font-size:13px;font-weight:700;color:${accent}">${heading}</div>
                    <div style="padding:0 14px">${pointsHtml}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="sections">${cols}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const colW = Math.floor((880 - (n - 1) * 16) / n);
            let cols = '';
            sections.forEach((sec, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const x = 65 + i * (colW + 16);
                const heading = esc(sec.heading) || _placeholder('Section ' + (i + 1));
                let pointsHtml = '';
                const pts = sec.points && sec.points.length ? sec.points : [''];
                pts.forEach((p, pi) => {
                    const ptText = esc(p) || _placeholder('Point ' + (pi + 1));
                    pointsHtml += `<div style="padding:4px 0;font-size:11px;color:#333;border-bottom:1px solid #E0DDD8">\u2022 ${ptText}</div>`;
                });
                cols += `<div style="position:absolute;left:${x}px;top:90px;width:${colW}px;height:420px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:${c};border-radius:8px 8px 0 0"></div>
                    <div style="padding:12px 14px 6px;font-size:13px;font-weight:800;color:${c}">${heading}</div>
                    <div style="padding:0 14px">${pointsHtml}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="sections">${cols}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const cols = buildCols('auto', '#444', '#fff', 'rgba(0,0,0,0.06)', true);
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="sections">${cols}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  28. QUOTE_FULL                                                       */
    /* --------------------------------------------------------------------- */
    quote_full(data, theme, sc) {
        const q = esc(data.quote || '');
        const attr = esc(data.attribution || '');
        const ctx = esc(data.context || '');
        const accent = _accent(sc);

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful" style="background:${accent}">
                <div style="position:absolute;left:80px;top:80px;font-size:96px;color:rgba(255,255,255,0.3);font-family:Georgia,serif;line-height:1">\u201C</div>
                <div data-field="quote" style="position:absolute;left:100px;top:160px;width:800px;font-size:22px;font-style:italic;color:#fff;line-height:1.6;text-align:center">${q || _placeholder('Quote text...')}</div>
                <div style="position:absolute;left:400px;top:380px;width:200px;height:3px;background:rgba(255,255,255,0.4)"></div>
                ${attr ? `<div data-field="attribution" style="position:absolute;left:100px;top:400px;width:800px;text-align:center;font-size:14px;font-weight:600;color:rgba(255,255,255,0.9)">\u2014 ${attr}</div>` : ''}
                ${ctx ? `<div data-field="context" style="position:absolute;left:100px;top:425px;width:800px;text-align:center;font-size:11px;color:rgba(255,255,255,0.7)">${ctx}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                <div style="position:absolute;left:80px;top:80px;font-size:96px;color:${accent};opacity:0.4;font-family:Georgia,serif;line-height:1">\u201C</div>
                <div data-field="quote" style="position:absolute;left:100px;top:160px;width:800px;font-size:22px;font-style:italic;color:#F0F0F0;line-height:1.6;text-align:center">${q || _placeholder('Quote text...')}</div>
                <div style="position:absolute;left:400px;top:380px;width:200px;height:3px;background:${accent}"></div>
                ${attr ? `<div data-field="attribution" style="position:absolute;left:100px;top:400px;width:800px;text-align:center;font-size:14px;font-weight:600;color:#F0F0F0">\u2014 ${attr}</div>` : ''}
                ${ctx ? `<div data-field="context" style="position:absolute;left:100px;top:425px;width:800px;text-align:center;font-size:11px;color:#999">${ctx}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:${accent}">
                <div style="position:absolute;left:0;top:0;width:30px;height:100%;background:rgba(0,0,0,0.15)"></div>
                <div style="position:absolute;right:0;top:0;width:12px;height:100%;background:rgba(0,0,0,0.08)"></div>
                <div style="position:absolute;left:80px;top:80px;font-size:96px;color:rgba(255,255,255,0.25);font-family:Georgia,serif;line-height:1">\u201C</div>
                <div data-field="quote" style="position:absolute;left:100px;top:160px;width:800px;font-size:22px;font-style:italic;color:#fff;line-height:1.6;text-align:center;font-weight:700">${q || _placeholder('Quote text...')}</div>
                <div style="position:absolute;left:400px;top:380px;width:200px;height:4px;background:rgba(255,255,255,0.3)"></div>
                ${attr ? `<div data-field="attribution" style="position:absolute;left:100px;top:400px;width:800px;text-align:center;font-size:14px;font-weight:800;color:rgba(255,255,255,0.9)">\u2014 ${attr}</div>` : ''}
                ${ctx ? `<div data-field="context" style="position:absolute;left:100px;top:425px;width:800px;text-align:center;font-size:11px;color:rgba(255,255,255,0.6)">${ctx}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bgColor = theme.startsWith('editorial') ? '#044014' : accent;
        return `<div class="preview-slide theme-${theme}" style="background:${bgColor}">
            <div style="position:absolute;left:80px;top:80px;font-size:96px;color:rgba(255,255,255,0.2);font-family:Georgia,serif;line-height:1">\u201C</div>
            <div data-field="quote" style="position:absolute;left:100px;top:160px;width:800px;font-size:22px;font-style:italic;color:#fff;line-height:1.6;text-align:center;font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${q || _placeholder('Quote text...')}</div>
            <div style="position:absolute;left:400px;top:380px;width:200px;height:3px;background:rgba(255,255,255,0.4)"></div>
            ${attr ? `<div data-field="attribution" style="position:absolute;left:100px;top:400px;width:800px;text-align:center;font-size:14px;font-weight:600;color:rgba(255,255,255,0.9)">\u2014 ${attr}</div>` : ''}
            ${ctx ? `<div data-field="context" style="position:absolute;left:100px;top:425px;width:800px;text-align:center;font-size:11px;color:rgba(255,255,255,0.7)">${ctx}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  29. STAT_HERO                                                        */
    /* --------------------------------------------------------------------- */
    stat_hero(data, theme, sc) {
        const t = esc(data.title || 'Key Metric');
        const hero = data.hero || {};
        const heroVal = esc(hero.value || '');
        const heroLabel = esc(hero.label || '');
        const supporting = data.supporting && data.supporting.length ? data.supporting : [];
        const source = esc(data.source || '');
        const accent = _accent(sc);

        function supportingHtml(valColor, labelColor) {
            const n = supporting.length;
            if (!n) return '';
            const w = Math.floor(600 / n);
            let html = '<div style="position:absolute;left:200px;top:370px;width:600px;display:flex;justify-content:center;gap:30px">';
            supporting.forEach((s) => {
                const v = esc(s.value) || '--';
                const l = esc(s.label) || '';
                html += `<div style="text-align:center;width:${w}px">
                    <div style="font-size:22px;font-weight:700;color:${valColor}">${v}</div>
                    <div style="font-size:10px;color:${labelColor}">${l}</div>
                </div>`;
            });
            html += '</div>';
            return html;
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="hero" style="position:absolute;left:100px;top:140px;width:800px;text-align:center">
                    <div style="font-size:110px;font-weight:700;color:${accent}">${heroVal || _placeholder('--')}</div>
                    <div style="font-size:16px;font-weight:600;color:#333;margin-top:-10px">${heroLabel}</div>
                </div>
                <div style="position:absolute;left:100px;top:350px;width:800px;height:1px;background:#ddd"></div>
                ${supportingHtml(accent, '#666')}
                ${source ? `<div data-field="source" style="position:absolute;left:100px;top:510px;width:800px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="hero" style="position:absolute;left:100px;top:110px;width:800px;text-align:center">
                    <div style="font-size:110px;font-weight:700;color:${accent}">${heroVal || _placeholder('--')}</div>
                    <div style="font-size:16px;font-weight:600;color:#F0F0F0;margin-top:-10px">${heroLabel}</div>
                </div>
                <div style="position:absolute;left:100px;top:350px;width:800px;height:1px;background:#2A2A2A"></div>
                ${supportingHtml(accent, '#999')}
                ${source ? `<div data-field="source" style="position:absolute;left:100px;top:510px;width:800px;text-align:center;font-size:10px;color:#666">Source: ${source}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="hero" style="position:absolute;left:100px;top:110px;width:800px;text-align:center">
                    <div style="font-size:110px;font-weight:800;color:${accent}">${heroVal || _placeholder('--')}</div>
                    <div style="font-size:16px;font-weight:700;color:#1A1A1A;margin-top:-10px">${heroLabel}</div>
                </div>
                <div style="position:absolute;left:100px;top:350px;width:800px;height:2px;background:#ccc"></div>
                ${supportingHtml(accent, '#666')}
                ${source ? `<div data-field="source" style="position:absolute;left:100px;top:510px;width:800px;text-align:center;font-size:10px;color:#888">Source: ${source}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="hero" style="position:absolute;left:100px;top:110px;width:800px;text-align:center">
                <div style="font-size:110px;font-weight:700;color:${accent};font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${heroVal || _placeholder('--')}</div>
                <div style="font-size:16px;font-weight:600;color:#333;margin-top:-10px">${heroLabel}</div>
            </div>
            <div style="position:absolute;left:100px;top:350px;width:800px;height:1px;background:#ddd"></div>
            ${supportingHtml(accent, '#666')}
            ${source ? `<div data-field="source" style="position:absolute;left:100px;top:510px;width:800px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  30. IN_BRIEF_FEATURED                                                */
    /* --------------------------------------------------------------------- */
    in_brief_featured(data, theme, sc) {
        const t = esc(data.title || 'Key Points');
        const featured = esc(data.featured || '');
        const supporting = data.supporting && data.supporting.length ? data.supporting : ['', ''];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let bullets = '';
            supporting.forEach((s, i) => {
                const text = esc(s) || _placeholder('Point ' + (i + 1));
                bullets += `<div style="padding:6px 0;font-size:12px;color:#444;border-bottom:1px solid rgba(0,0,0,0.06)"><span style="color:${PREVIEW_COLORS[i % 5]};margin-right:8px">\u2022</span>${text}</div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="featured" style="position:absolute;left:30px;top:120px;width:940px;height:140px;background:${PREVIEW_LIGHTS[0]};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${accent}"></div>
                    <div style="padding:20px 24px;font-size:16px;font-weight:600;color:#333;line-height:1.5">${featured || _placeholder('Featured point...')}</div>
                </div>
                <div data-field="supporting" style="position:absolute;left:30px;top:280px;width:940px;padding:0 16px">${bullets}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let bullets = '';
            supporting.forEach((s, i) => {
                const text = esc(s) || _placeholder('Point ' + (i + 1));
                bullets += `<div style="padding:6px 0;font-size:12px;color:#999;border-bottom:1px solid #2A2A2A"><span style="color:${accent};margin-right:8px">\u2022</span>${text}</div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="featured" style="position:absolute;left:45px;top:90px;width:920px;height:140px;background:#141414;border-radius:8px;border:1px solid #2A2A2A">
                    <div style="padding:20px 24px;font-size:16px;font-weight:600;color:#F0F0F0;line-height:1.5">${featured || _placeholder('Featured point...')}</div>
                </div>
                <div data-field="supporting" style="position:absolute;left:45px;top:250px;width:920px;padding:0 16px">${bullets}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let bullets = '';
            supporting.forEach((s, i) => {
                const text = esc(s) || _placeholder('Point ' + (i + 1));
                bullets += `<div style="padding:6px 0;font-size:12px;color:#333;border-bottom:1px solid #E0DDD8"><span style="color:${PREVIEW_COLORS[i % 5]};margin-right:8px">\u2022</span>${text}</div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="featured" style="position:absolute;left:65px;top:90px;width:880px;height:140px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:${accent};border-radius:8px 8px 0 0"></div>
                    <div style="padding:20px 24px;font-size:16px;font-weight:700;color:#1A1A1A;line-height:1.5">${featured || _placeholder('Featured point...')}</div>
                </div>
                <div data-field="supporting" style="position:absolute;left:65px;top:250px;width:880px;padding:0 16px">${bullets}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let bullets = '';
        supporting.forEach((s, i) => {
            const text = esc(s) || _placeholder('Point ' + (i + 1));
            bullets += `<div style="padding:6px 0;font-size:12px;color:#444;border-bottom:1px solid rgba(0,0,0,0.06)"><span style="color:${PREVIEW_COLORS[i % 5]};margin-right:8px">\u2022</span>${text}</div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="featured" style="position:absolute;left:60px;top:90px;width:900px;height:140px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${accent}"></div>
                <div style="padding:20px 24px 20px 20px;font-size:16px;font-weight:600;color:#333;line-height:1.5">${featured || _placeholder('Featured point...')}</div>
            </div>
            <div data-field="supporting" style="position:absolute;left:60px;top:250px;width:900px;padding:0 16px">${bullets}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  IN_BRIEF_REVEAL — Spotlight reveal                                   */
    /* --------------------------------------------------------------------- */
    in_brief_reveal(data, theme, sc) {
        const t = esc(data.title || 'Key Points');
        const items = data.items && data.items.length ? data.items : ['', '', '', ''];
        const accent = _accent(sc);
        const n = items.length;

        /* Show the FIRST item as featured (preview = slide 1 state) */
        const featIdx = 0;

        function renderStack(slideOpen, titleHtml, cardBg, featuredBg, textColor, mutedColor, borderColor) {
            let html = '';
            let y = 75;
            for (let i = 0; i < n; i++) {
                const col = PREVIEW_COLORS[i % 5];
                const text = esc(items[i]) || _placeholder('Point ' + (i + 1));
                if (i === featIdx) {
                    /* Featured card */
                    html += `<div style="position:absolute;left:${slideOpen.lx}px;top:${y}px;width:${slideOpen.w}px;height:95px;background:${featuredBg};border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.06)">
                        <div style="position:absolute;left:0;top:0;width:8px;height:100%;background:${col}"></div>
                        <div style="padding:14px 16px 14px 20px;font-size:14px;font-weight:700;color:${textColor};line-height:1.5">${text}</div>
                    </div>`;
                    y += 103;
                } else {
                    /* Small row */
                    const c2 = i < featIdx ? mutedColor : textColor;
                    html += `<div style="position:absolute;left:${slideOpen.lx}px;top:${y}px;width:${slideOpen.w}px;height:38px;background:${cardBg};border-radius:4px;overflow:hidden">
                        <div style="position:absolute;left:0;top:0;width:4px;height:100%;background:${col}"></div>
                        <div style="padding:8px 12px 8px 14px;font-size:11px;color:${c2}">${text}</div>
                    </div>`;
                    y += 44;
                }
            }
            return html;
        }

        if (theme === 'colorful') {
            const stack = renderStack({lx: 30, w: 940}, '', '#fff', PREVIEW_LIGHTS[0], '#333', '#999', '#eee');
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${stack}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const stack = renderStack({lx: 45, w: 920}, '', '#141414', '#1a1a1a', '#F0F0F0', '#666', '#2A2A2A');
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${stack}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const stack = renderStack({lx: 65, w: 880}, '', '#fff', '#fff', '#1A1A1A', '#888', '#E0DDD8');
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${stack}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const stack = renderStack({lx: 60, w: 900}, '', '#fff', '#fff', '#333', '#999', '#eee');
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${stack}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  31. PERSONA_DUO                                                      */
    /* --------------------------------------------------------------------- */
    persona_duo(data, theme, sc) {
        const t = esc(data.title || 'Personas');
        const personas = data.personas && data.personas.length >= 2 ? data.personas : [
            { name: '', archetype: '', traits: [], strategy: '' },
            { name: '', archetype: '', traits: [], strategy: '' },
        ];
        const accent = _accent(sc);

        function personaCard(p, idx, cardBg, nameFg, traitBg, traitFg, stratFg, border, leftStripe) {
            const c = PREVIEW_COLORS[idx % 5];
            const name = esc(p.name) || _placeholder('Persona ' + (idx + 1));
            const archetype = esc(p.archetype);
            const traits = p.traits && p.traits.length ? p.traits : [];
            const strategy = esc(p.strategy);
            const pills = traits.map(tr => `<span style="display:inline-block;padding:2px 8px;margin:2px 3px 2px 0;border-radius:10px;font-size:9px;background:${traitBg};color:${traitFg}">${esc(tr)}</span>`).join('');
            return `<div style="width:100%;height:100%;background:${cardBg};border-radius:8px;overflow:hidden;${border ? `border:1px solid ${border}` : `box-shadow:0 1px 4px rgba(0,0,0,0.08)`}">
                ${leftStripe ? `<div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>` : `<div style="height:4px;background:${c}"></div>`}
                <div style="padding:14px 16px 4px ${leftStripe ? '18px' : '16px'}">
                    <span style="font-size:18px;font-weight:700;color:${nameFg}">${name}</span>
                    ${archetype ? `<span style="display:inline-block;margin-left:8px;padding:2px 10px;border-radius:10px;font-size:10px;font-weight:600;background:${c};color:#fff">${archetype}</span>` : ''}
                </div>
                ${traits.length ? `<div style="padding:6px 16px 0 ${leftStripe ? '18px' : '16px'}">${pills}</div>` : ''}
                ${strategy ? `<div style="padding:8px 16px 0 ${leftStripe ? '18px' : '16px'}"><span style="font-size:10px;font-weight:600;color:${c}">Strategy:</span> <span style="font-size:10px;color:${stratFg}">${strategy}</span></div>` : ''}
            </div>`;
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:30px;top:120px;width:455px;height:390px">${personaCard(personas[0], 0, PREVIEW_LIGHTS[0], '#333', PREVIEW_LIGHTS[1], '#3880F3', '#444', null, false)}</div>
                <div style="position:absolute;left:505px;top:120px;width:465px;height:390px">${personaCard(personas[1], 1, PREVIEW_LIGHTS[2], '#333', PREVIEW_LIGHTS[3], '#04547C', '#444', null, false)}</div>
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:30px;top:90px;width:460px;height:420px">${personaCard(personas[0], 0, '#141414', '#F0F0F0', '#2A2A2A', '#F0F0F0', '#999', '#2A2A2A', false)}</div>
                <div style="position:absolute;left:510px;top:90px;width:460px;height:420px">${personaCard(personas[1], 1, '#141414', '#F0F0F0', '#2A2A2A', '#F0F0F0', '#999', '#2A2A2A', false)}</div>
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:45px;top:90px;width:445px;height:420px">${personaCard(personas[0], 0, '#fff', '#1A1A1A', '#F2F0EC', '#1A1A1A', '#444', null, false)}</div>
                <div style="position:absolute;left:510px;top:90px;width:445px;height:420px">${personaCard(personas[1], 1, '#fff', '#1A1A1A', '#F2F0EC', '#1A1A1A', '#444', null, false)}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:60px;top:90px;width:430px;height:420px">${personaCard(personas[0], 0, '#fff', '#333', '#F0F0F0', '#555', '#444', null, true)}</div>
            <div style="position:absolute;left:510px;top:90px;width:450px;height:420px">${personaCard(personas[1], 1, '#fff', '#333', '#F0F0F0', '#555', '#444', null, true)}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  32. PROCESS_FLOW_VERTICAL                                            */
    /* --------------------------------------------------------------------- */
    process_flow_vertical(data, theme, sc) {
        const t = esc(data.title || 'Process');
        const steps = data.steps && data.steps.length ? data.steps : [
            { title: '', detail: '' }, { title: '', detail: '' }, { title: '', detail: '' }
        ];
        const accent = _accent(sc);

        function buildSteps(cardBg, titleColor, detailColor, borderOrShadow, arrowColor, leftStripe) {
            const stepH = 58;
            const arrowH = 20;
            const startY = 80;
            let html = '';
            steps.forEach((step, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (stepH + arrowH);
                const sTitle = esc(step.title) || _placeholder('Step ' + (i + 1));
                const detail = esc(step.detail);
                html += `<div style="position:absolute;left:60px;top:${y}px;width:880px;height:${stepH}px;background:${cardBg};border-radius:6px;overflow:hidden;${borderOrShadow}">
                    ${leftStripe ? `<div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>` : `<div style="height:4px;background:${c}"></div>`}
                    <div style="position:absolute;left:${leftStripe ? '18px' : '14px'};top:${leftStripe ? '10px' : '12px'};width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:56px;top:${leftStripe ? '10px' : '12px'};font-size:13px;font-weight:600;color:${titleColor}">${sTitle}</div>
                    ${detail ? `<div style="position:absolute;left:56px;top:${leftStripe ? '30px' : '32px'};font-size:10px;color:${detailColor};width:800px">${detail}</div>` : ''}
                </div>`;
                if (i < steps.length - 1) {
                    html += `<div style="position:absolute;left:490px;top:${y + stepH + 2}px;font-size:16px;color:${arrowColor};text-align:center;width:20px">\u2193</div>`;
                }
            });
            return html;
        }

        if (theme === 'colorful') {
            const stepH = 58;
            const arrowH = 20;
            const startY = 120;
            let html = '';
            steps.forEach((step, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (stepH + arrowH);
                const sTitle = esc(step.title) || _placeholder('Step ' + (i + 1));
                const detail = esc(step.detail);
                html += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${stepH}px;background:${bg2};border-radius:6px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:12px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:56px;top:12px;font-size:13px;font-weight:600;color:#333">${sTitle}</div>
                    ${detail ? `<div style="position:absolute;left:56px;top:32px;font-size:10px;color:#666;width:860px">${detail}</div>` : ''}
                </div>`;
                if (i < steps.length - 1) {
                    html += `<div style="position:absolute;left:490px;top:${y + stepH + 2}px;font-size:16px;color:#bbb;text-align:center;width:20px">\u2193</div>`;
                }
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="steps">${html}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const stepH = 58;
            const arrowH = 20;
            const startY = 90;
            let html = '';
            steps.forEach((step, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (stepH + arrowH);
                const sTitle = esc(step.title) || _placeholder('Step ' + (i + 1));
                const detail = esc(step.detail);
                html += `<div style="position:absolute;left:45px;top:${y}px;width:920px;height:${stepH}px;background:#141414;border-radius:6px;border:1px solid #2A2A2A">
                    <div style="position:absolute;left:14px;top:10px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:56px;top:10px;font-size:13px;font-weight:600;color:#F0F0F0">${sTitle}</div>
                    ${detail ? `<div style="position:absolute;left:56px;top:30px;font-size:10px;color:#999;width:840px">${detail}</div>` : ''}
                </div>`;
                if (i < steps.length - 1) {
                    html += `<div style="position:absolute;left:495px;top:${y + stepH + 2}px;font-size:16px;color:#555;text-align:center;width:20px">\u2193</div>`;
                }
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="steps">${html}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const stepH = 58;
            const arrowH = 20;
            const startY = 90;
            let html = '';
            steps.forEach((step, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (stepH + arrowH);
                const sTitle = esc(step.title) || _placeholder('Step ' + (i + 1));
                const detail = esc(step.detail);
                html += `<div style="position:absolute;left:65px;top:${y}px;width:880px;height:${stepH}px;background:#fff;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.08);overflow:hidden">
                    <div style="height:5px;background:${c};border-radius:6px 6px 0 0"></div>
                    <div style="position:absolute;left:14px;top:12px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="position:absolute;left:56px;top:12px;font-size:13px;font-weight:700;color:#1A1A1A">${sTitle}</div>
                    ${detail ? `<div style="position:absolute;left:56px;top:32px;font-size:10px;color:#666;width:800px">${detail}</div>` : ''}
                </div>`;
                if (i < steps.length - 1) {
                    html += `<div style="position:absolute;left:495px;top:${y + stepH + 2}px;font-size:16px;color:#bbb;text-align:center;width:20px">\u2193</div>`;
                }
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="steps">${html}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const stepH = 58;
        const arrowH = 20;
        const startY = 80;
        let html = '';
        steps.forEach((step, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (stepH + arrowH);
            const sTitle = esc(step.title) || _placeholder('Step ' + (i + 1));
            const detail = esc(step.detail);
            html += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${stepH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:18px;top:10px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                <div style="position:absolute;left:56px;top:10px;font-size:13px;font-weight:600;color:#333">${sTitle}</div>
                ${detail ? `<div style="position:absolute;left:56px;top:30px;font-size:10px;color:#666;width:820px">${detail}</div>` : ''}
            </div>`;
            if (i < steps.length - 1) {
                html += `<div style="position:absolute;left:500px;top:${y + stepH + 2}px;font-size:16px;color:#bbb;text-align:center;width:20px">\u2193</div>`;
            }
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="steps">${html}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  33. TEXT_CARDS                                                        */
    /* --------------------------------------------------------------------- */
    text_cards(data, theme, sc) {
        const t = esc(data.title || 'Overview');
        const items = data.items && data.items.length ? data.items : [
            { title: '', detail: '' }, { title: '', detail: '' },
            { title: '', detail: '' }, { title: '', detail: '' },
        ];
        const accent = _accent(sc);
        const n = items.length;
        const cols = 2;
        const rows = Math.ceil(n / cols);

        if (theme === 'colorful') {
            const cardW = 455;
            const cardH = Math.min(180, Math.floor((400 - (rows - 1) * 12) / rows));
            let cards = '';
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const col = i % cols;
                const row = Math.floor(i / cols);
                const x = 30 + col * (cardW + 20);
                const y = 120 + row * (cardH + 12);
                const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
                const detail = esc(item.detail);
                cards += `<div style="position:absolute;left:${x}px;top:${y}px;width:${cardW}px;height:${cardH}px;background:${bg2};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:10px 14px 4px;font-size:13px;font-weight:700;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="padding:0 14px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${cards}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const cardW = 445;
            const cardH = Math.min(190, Math.floor((430 - (rows - 1) * 12) / rows));
            let cards = '';
            items.forEach((item, i) => {
                const col = i % cols;
                const row = Math.floor(i / cols);
                const x = 30 + col * (cardW + 20);
                const y = 90 + row * (cardH + 12);
                const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
                const detail = esc(item.detail);
                cards += `<div style="position:absolute;left:${x}px;top:${y}px;width:${cardW}px;height:${cardH}px;background:#141414;border-radius:8px;border:1px solid #2A2A2A">
                    <div style="padding:10px 14px 4px;font-size:13px;font-weight:700;color:${accent}">${iTitle}</div>
                    ${detail ? `<div style="padding:0 14px;font-size:11px;color:#999;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${cards}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const cardW = 430;
            const cardH = Math.min(190, Math.floor((430 - (rows - 1) * 12) / rows));
            let cards = '';
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const col = i % cols;
                const row = Math.floor(i / cols);
                const x = 50 + col * (cardW + 20);
                const y = 90 + row * (cardH + 12);
                const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
                const detail = esc(item.detail);
                cards += `<div style="position:absolute;left:${x}px;top:${y}px;width:${cardW}px;height:${cardH}px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:${c};border-radius:8px 8px 0 0"></div>
                    <div style="padding:10px 14px 4px;font-size:13px;font-weight:800;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="padding:0 14px;font-size:11px;color:#444;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${cards}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const cardW = 435;
        const cardH = Math.min(200, Math.floor((440 - (rows - 1) * 12) / rows));
        let cards = '';
        items.forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const col = i % cols;
            const row = Math.floor(i / cols);
            const x = 60 + col * (cardW + 16);
            const y = 80 + row * (cardH + 12);
            const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
            const detail = esc(item.detail);
            cards += `<div style="position:absolute;left:${x}px;top:${y}px;width:${cardW}px;height:${cardH}px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="padding:10px 14px 4px 18px;font-size:13px;font-weight:700;color:${c}">${iTitle}</div>
                ${detail ? `<div style="padding:0 14px 0 18px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${cards}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  34. TEXT_COLUMNS                                                      */
    /* --------------------------------------------------------------------- */
    text_columns(data, theme, sc) {
        const t = esc(data.title || 'Overview');
        const columns = data.columns && data.columns.length ? data.columns : [
            { heading: '', body: '' }, { heading: '', body: '' },
        ];
        const accent = _accent(sc);
        const n = columns.length;

        if (theme === 'colorful') {
            const colW = Math.floor((940 - (n - 1) * 16) / n);
            let cols = '';
            columns.forEach((col, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const x = 30 + i * (colW + 16);
                const heading = esc(col.heading) || _placeholder('Column ' + (i + 1));
                const body = esc(col.body);
                cols += `<div style="position:absolute;left:${x}px;top:120px;width:${colW}px;height:400px;background:${bg2};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="padding:12px 14px 6px;font-size:13px;font-weight:700;color:${c}">${heading}</div>
                    ${body ? `<div style="padding:0 14px;font-size:11px;color:#444;line-height:1.5">${body}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="columns">${cols}</div>
            </div>`;
        }

        if (theme === 'noir') {
            const colW = Math.floor((900 - (n - 1) * 16) / n);
            let cols = '';
            columns.forEach((col, i) => {
                const x = 50 + i * (colW + 16);
                const heading = esc(col.heading) || _placeholder('Column ' + (i + 1));
                const body = esc(col.body);
                cols += `<div style="position:absolute;left:${x}px;top:90px;width:${colW}px;height:420px;background:#141414;border-radius:8px;border:1px solid #2A2A2A">
                    <div style="padding:12px 14px 6px;font-size:13px;font-weight:700;color:${accent}">${heading}</div>
                    ${body ? `<div style="padding:0 14px;font-size:11px;color:#999;line-height:1.5">${body}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="columns">${cols}</div>
            </div>`;
        }

        if (theme === 'bold') {
            const colW = Math.floor((880 - (n - 1) * 16) / n);
            let cols = '';
            columns.forEach((col, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const x = 65 + i * (colW + 16);
                const heading = esc(col.heading) || _placeholder('Column ' + (i + 1));
                const body = esc(col.body);
                cols += `<div style="position:absolute;left:${x}px;top:90px;width:${colW}px;height:420px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.08)">
                    <div style="height:5px;background:${c};border-radius:8px 8px 0 0"></div>
                    <div style="padding:12px 14px 6px;font-size:13px;font-weight:800;color:${c}">${heading}</div>
                    ${body ? `<div style="padding:0 14px;font-size:11px;color:#444;line-height:1.5">${body}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="columns">${cols}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        const colW = Math.floor((900 - (n - 1) * 16) / n);
        let cols = '';
        columns.forEach((col, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const x = 60 + i * (colW + 16);
            const heading = esc(col.heading) || _placeholder('Column ' + (i + 1));
            const body = esc(col.body);
            cols += `<div style="position:absolute;left:${x}px;top:80px;width:${colW}px;height:440px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="padding:12px 14px 6px 18px;font-size:13px;font-weight:700;color:${c}">${heading}</div>
                ${body ? `<div style="padding:0 14px 0 18px;font-size:11px;color:#444;line-height:1.5">${body}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="columns">${cols}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  35. TEXT_NARRATIVE                                                    */
    /* --------------------------------------------------------------------- */
    text_narrative(data, theme, sc) {
        const t = esc(data.title || 'Narrative');
        const lede = esc(data.lede || '');
        const body = esc(data.body || '');
        const accent = _accent(sc);

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="lede" style="position:absolute;left:60px;top:120px;width:880px;font-size:18px;font-weight:600;color:#333;line-height:1.6">${lede || _placeholder('Lead paragraph...')}</div>
                <div style="position:absolute;left:60px;top:220px;width:200px;height:3px;background:${accent}"></div>
                ${body ? `<div data-field="body" style="position:absolute;left:60px;top:240px;width:880px;font-size:12px;color:#555;line-height:1.6">${body}</div>` : ''}
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="lede" style="position:absolute;left:50px;top:90px;width:900px;font-size:18px;font-weight:600;color:#F0F0F0;line-height:1.6">${lede || _placeholder('Lead paragraph...')}</div>
                <div style="position:absolute;left:50px;top:190px;width:200px;height:3px;background:${accent}"></div>
                ${body ? `<div data-field="body" style="position:absolute;left:50px;top:210px;width:900px;font-size:12px;color:#999;line-height:1.6">${body}</div>` : ''}
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="lede" style="position:absolute;left:65px;top:90px;width:880px;font-size:18px;font-weight:700;color:#1A1A1A;line-height:1.6">${lede || _placeholder('Lead paragraph...')}</div>
                <div style="position:absolute;left:65px;top:190px;width:200px;height:4px;background:${accent}"></div>
                ${body ? `<div data-field="body" style="position:absolute;left:65px;top:210px;width:880px;font-size:12px;color:#444;line-height:1.6">${body}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="lede" style="position:absolute;left:60px;top:90px;width:900px;font-size:18px;font-weight:600;color:#333;line-height:1.6;font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : 'inherit'}">${lede || _placeholder('Lead paragraph...')}</div>
            <div style="position:absolute;left:60px;top:190px;width:200px;height:3px;background:${theme.startsWith('editorial') ? '#044014' : accent}"></div>
            ${body ? `<div data-field="body" style="position:absolute;left:60px;top:210px;width:900px;font-size:12px;color:#555;line-height:1.6">${body}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  36. TEXT_NESTED                                                       */
    /* --------------------------------------------------------------------- */
    text_nested(data, theme, sc) {
        const t = esc(data.title || 'Overview');
        const items = data.items && data.items.length ? data.items : [
            { text: '', children: [''] },
            { text: '', children: [''] },
        ];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const startY = 120;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * 90;
                const text = esc(item.text) || _placeholder('Section ' + (i + 1));
                const children = item.children && item.children.length ? item.children : [];
                let childHtml = children.map(ch => `<div style="padding:2px 0;font-size:10px;color:#555">\u2022 ${esc(ch) || _placeholder('Sub-item')}</div>`).join('');
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px">
                    <div style="display:inline-block;padding:4px 12px;border-radius:6px;font-size:12px;font-weight:700;background:${c};color:#fff;margin-bottom:4px">${text}</div>
                    <div style="padding-left:20px">${childHtml}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const startY = 90;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * 90;
                const text = esc(item.text) || _placeholder('Section ' + (i + 1));
                const children = item.children && item.children.length ? item.children : [];
                let childHtml = children.map(ch => `<div style="padding:2px 0;font-size:10px;color:#999">\u2022 ${esc(ch) || _placeholder('Sub-item')}</div>`).join('');
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px">
                    <div style="display:inline-block;padding:4px 12px;border-radius:6px;font-size:12px;font-weight:700;background:${c};color:#fff;margin-bottom:4px">${text}</div>
                    <div style="padding-left:20px">${childHtml}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const startY = 90;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * 90;
                const text = esc(item.text) || _placeholder('Section ' + (i + 1));
                const children = item.children && item.children.length ? item.children : [];
                let childHtml = children.map(ch => `<div style="padding:2px 0;font-size:10px;color:#444">\u2022 ${esc(ch) || _placeholder('Sub-item')}</div>`).join('');
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:880px">
                    <div style="display:inline-block;padding:4px 14px;border-radius:6px;font-size:12px;font-weight:800;background:${c};color:#fff;margin-bottom:4px">${text}</div>
                    <div style="padding-left:20px">${childHtml}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        const startY = 80;
        items.forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * 90;
            const text = esc(item.text) || _placeholder('Section ' + (i + 1));
            const children = item.children && item.children.length ? item.children : [];
            let childHtml = children.map(ch => `<div style="padding:2px 0;font-size:10px;color:#555">\u2022 ${esc(ch) || _placeholder('Sub-item')}</div>`).join('');
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px">
                <div style="display:inline-block;padding:4px 12px;border-radius:6px;font-size:12px;font-weight:700;background:${c};color:#fff;margin-bottom:4px">${text}</div>
                <div style="padding-left:20px">${childHtml}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  37. TEXT_SPLIT                                                        */
    /* --------------------------------------------------------------------- */
    text_split(data, theme, sc) {
        const t = esc(data.title || 'Overview');
        const headline = esc(data.headline || '');
        const detail = esc(data.detail || '');
        const points = data.points && data.points.length ? data.points : ['', ''];
        const accent = _accent(sc);

        function pointsList(bulletColor, textColor) {
            return points.map((p, i) => {
                const text = esc(p) || _placeholder('Point ' + (i + 1));
                return `<div style="padding:5px 0;font-size:12px;color:${textColor};border-bottom:1px solid rgba(0,0,0,0.06)"><span style="color:${bulletColor};margin-right:8px">\u2022</span>${text}</div>`;
            }).join('');
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div style="position:absolute;left:30px;top:120px;width:440px">
                    <div data-field="headline" style="font-size:20px;font-weight:700;color:${accent};margin-bottom:12px">${headline || _placeholder('Headline')}</div>
                    ${detail ? `<div data-field="detail" style="font-size:12px;color:#555;line-height:1.5">${detail}</div>` : ''}
                </div>
                <div style="position:absolute;left:490px;top:120px;width:1px;height:380px;background:#ddd"></div>
                <div data-field="points" style="position:absolute;left:510px;top:120px;width:460px">${pointsList(accent, '#444')}</div>
            </div>`;
        }

        if (theme === 'noir') {
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div style="position:absolute;left:45px;top:90px;width:440px">
                    <div data-field="headline" style="font-size:20px;font-weight:700;color:#F0F0F0;margin-bottom:12px">${headline || _placeholder('Headline')}</div>
                    ${detail ? `<div data-field="detail" style="font-size:12px;color:#999;line-height:1.5">${detail}</div>` : ''}
                </div>
                <div style="position:absolute;left:500px;top:90px;width:1px;height:400px;background:#2A2A2A"></div>
                <div data-field="points" style="position:absolute;left:520px;top:90px;width:440px">${pointsList(accent, '#999')}</div>
            </div>`;
        }

        if (theme === 'bold') {
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div style="position:absolute;left:65px;top:90px;width:420px">
                    <div data-field="headline" style="font-size:20px;font-weight:800;color:#1A1A1A;margin-bottom:12px">${headline || _placeholder('Headline')}</div>
                    ${detail ? `<div data-field="detail" style="font-size:12px;color:#444;line-height:1.5">${detail}</div>` : ''}
                </div>
                <div style="position:absolute;left:505px;top:90px;width:2px;height:400px;background:#ccc"></div>
                <div data-field="points" style="position:absolute;left:525px;top:90px;width:430px">${pointsList(accent, '#333')}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:60px;top:90px;width:420px">
                <div data-field="headline" style="font-size:20px;font-weight:700;color:#333;margin-bottom:12px;font-family:${theme.startsWith('editorial') ? 'Georgia, serif' : 'inherit'}">${headline || _placeholder('Headline')}</div>
                ${detail ? `<div data-field="detail" style="font-size:12px;color:#555;line-height:1.5">${detail}</div>` : ''}
            </div>
            <div style="position:absolute;left:500px;top:90px;width:1px;height:420px;background:#ddd"></div>
            <div data-field="points" style="position:absolute;left:520px;top:90px;width:440px">${pointsList(accent, '#444')}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  38. TEXT_ANNOTATED                                                    */
    /* --------------------------------------------------------------------- */
    text_annotated(data, theme, sc) {
        const t = esc(data.title || 'Details');
        const items = data.items && data.items.length ? data.items : [
            { label: '', text: '' }, { label: '', text: '' }, { label: '', text: '' },
        ];
        const accent = _accent(sc);

        if (theme === 'colorful') {
            let rows = '';
            const startY = 120;
            const rowH = 48;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(item.label) || _placeholder('Label');
                const text = esc(item.text) || _placeholder('Description');
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${rowH}px;display:flex;align-items:center">
                    <div style="display:inline-block;padding:5px 14px;border-radius:6px;font-size:11px;font-weight:700;background:${c};color:#fff;min-width:100px;text-align:center">${label}</div>
                    <div style="margin-left:14px;font-size:12px;color:#444;flex:1">${text}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            const startY = 90;
            const rowH = 48;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(item.label) || _placeholder('Label');
                const text = esc(item.text) || _placeholder('Description');
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px;height:${rowH}px;display:flex;align-items:center">
                    <div style="display:inline-block;padding:5px 14px;border-radius:6px;font-size:11px;font-weight:700;background:${c};color:#fff;min-width:100px;text-align:center">${label}</div>
                    <div style="margin-left:14px;font-size:12px;color:#999;flex:1">${text}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            const startY = 90;
            const rowH = 48;
            items.forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = startY + i * (rowH + 10);
                const label = esc(item.label) || _placeholder('Label');
                const text = esc(item.text) || _placeholder('Description');
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:880px;height:${rowH}px;display:flex;align-items:center">
                    <div style="display:inline-block;padding:5px 14px;border-radius:6px;font-size:11px;font-weight:800;background:${c};color:#fff;min-width:100px;text-align:center">${label}</div>
                    <div style="margin-left:14px;font-size:12px;color:#333;flex:1">${text}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        const startY = 80;
        const rowH = 48;
        items.forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 10);
            const label = esc(item.label) || _placeholder('Label');
            const text = esc(item.text) || _placeholder('Description');
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH}px;display:flex;align-items:center">
                <div style="display:inline-block;padding:5px 14px;border-radius:6px;font-size:11px;font-weight:700;background:${c};color:#fff;min-width:100px;text-align:center">${label}</div>
                <div style="margin-left:14px;font-size:12px;color:#444;flex:1">${text}</div>
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${rows}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  ICON_CARDS                                                            */
    /* --------------------------------------------------------------------- */
    icon_cards(data, theme, sc) {
        const t = esc(data.title || 'Key Points');
        const items = data.items && data.items.length ? data.items : [
            { title: '', detail: '' }, { title: '', detail: '' }, { title: '', detail: '' },
        ];
        const accent = _accent(sc);
        const n = Math.min(items.length, 3);
        const iconSize = 50;
        const gap = 16;
        const totalW = 880;
        const cardW = Math.floor((totalW - (n - 1) * gap) / n);

        if (theme === 'colorful') {
            let cards = '';
            items.slice(0, n).forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const x = 30 + i * (cardW + gap);
                const iconCx = x + cardW / 2;
                const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
                const detail = esc(item.detail);
                cards += `<svg style="position:absolute;left:${iconCx - iconSize/2}px;top:115px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>`;
                cards += `<div style="position:absolute;left:${x}px;top:${115 + iconSize + 10}px;width:${cardW}px;height:200px;background:${bg2};border-radius:8px;overflow:hidden">
                    <div style="height:5px;background:${c}"></div>
                    <div style="padding:10px 14px 4px;font-size:13px;font-weight:700;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="padding:0 14px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${cards}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let cards = '';
            items.slice(0, n).forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const x = 45 + i * (cardW + gap);
                const iconCx = x + cardW / 2;
                const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
                const detail = esc(item.detail);
                cards += `<svg style="position:absolute;left:${iconCx - iconSize/2}px;top:90px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>`;
                cards += `<div style="position:absolute;left:${x}px;top:${90 + iconSize + 10}px;width:${cardW}px;height:200px;background:#141414;border:1px solid #2A2A2A;border-radius:8px;overflow:hidden">
                    <div style="padding:10px 14px 4px;font-size:13px;font-weight:700;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="padding:0 14px;font-size:11px;color:#999;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${cards}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let cards = '';
            items.slice(0, n).forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const x = 65 + i * (cardW + gap);
                const iconCx = x + cardW / 2;
                const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
                const detail = esc(item.detail);
                cards += `<svg style="position:absolute;left:${iconCx - iconSize/2}px;top:90px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>`;
                cards += `<div style="position:absolute;left:${x}px;top:${90 + iconSize + 10}px;width:${cardW}px;height:200px;background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,0.08);overflow:hidden">
                    <div style="padding:10px 14px 4px;font-size:13px;font-weight:800;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="padding:0 14px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${cards}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let cards = '';
        items.slice(0, n).forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const x = 60 + i * (cardW + gap);
            const iconCx = x + cardW / 2;
            const iTitle = esc(item.title) || _placeholder('Card ' + (i + 1));
            const detail = esc(item.detail);
            cards += `<svg style="position:absolute;left:${iconCx - iconSize/2}px;top:80px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>`;
            cards += `<div style="position:absolute;left:${x}px;top:${80 + iconSize + 10}px;width:${cardW}px;height:200px;background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,0.08);overflow:hidden">
                <div style="padding:10px 14px 4px;font-size:13px;font-weight:700;color:${c}">${iTitle}</div>
                ${detail ? `<div style="padding:0 14px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${cards}</div>
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  FEATURE_CARDS                                                         */
    /* --------------------------------------------------------------------- */
    feature_cards(data, theme, sc) {
        const t = esc(data.title || 'Features');
        const items = data.items && data.items.length ? data.items : [
            { title: '', detail: '' }, { title: '', detail: '' },
        ];
        const accent = _accent(sc);
        const n = Math.min(items.length, 2);
        const iconSize = 60;
        const gap = 14;
        const rowH = Math.floor((300 - (n - 1) * gap) / n);

        if (theme === 'colorful') {
            let rows = '';
            items.slice(0, n).forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg2 = PREVIEW_LIGHTS[i % 5];
                const y = 120 + i * (rowH + gap);
                const iTitle = esc(item.title) || _placeholder('Feature ' + (i + 1));
                const detail = esc(item.detail);
                const iconCy = y + rowH / 2;
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${rowH}px;background:${bg2};border-radius:8px;overflow:hidden">
                    <div style="height:5px;background:${c}"></div>
                    <svg style="position:absolute;left:20px;top:${(rowH - iconSize) / 2}px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>
                    <div style="position:absolute;left:100px;top:14px;width:820px;font-size:14px;font-weight:700;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="position:absolute;left:100px;top:38px;width:820px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'noir') {
            let rows = '';
            items.slice(0, n).forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = 90 + i * (rowH + gap);
                const iTitle = esc(item.title) || _placeholder('Feature ' + (i + 1));
                const detail = esc(item.detail);
                rows += `<div style="position:absolute;left:45px;top:${y}px;width:920px;height:${rowH}px;background:#141414;border:1px solid #2A2A2A;border-radius:8px;overflow:hidden">
                    <svg style="position:absolute;left:20px;top:${(rowH - iconSize) / 2}px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>
                    <div style="position:absolute;left:100px;top:14px;width:820px;font-size:14px;font-weight:700;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="position:absolute;left:100px;top:38px;width:820px;font-size:11px;color:#999;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-noir" style="background:#0D0D0D">
                ${_noirBar(accent)}
                ${_noirTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        if (theme === 'bold') {
            let rows = '';
            items.slice(0, n).forEach((item, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const y = 90 + i * (rowH + gap);
                const iTitle = esc(item.title) || _placeholder('Feature ' + (i + 1));
                const detail = esc(item.detail);
                rows += `<div style="position:absolute;left:65px;top:${y}px;width:880px;height:${rowH}px;background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,0.08);overflow:hidden">
                    <svg style="position:absolute;left:20px;top:${(rowH - iconSize) / 2}px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>
                    <div style="position:absolute;left:100px;top:14px;width:760px;font-size:14px;font-weight:800;color:${c}">${iTitle}</div>
                    ${detail ? `<div style="position:absolute;left:100px;top:38px;width:760px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
                </div>`;
            });
            return `<div class="preview-slide theme-bold" style="background:#F2F0EC">
                ${_boldBar(accent)}
                ${_boldTitle(t, accent)}
                <div data-field="items">${rows}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme.startsWith('editorial') ? '#FAF6F0' : '#fff';
        let rows = '';
        items.slice(0, n).forEach((item, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = 80 + i * (rowH + gap);
            const iTitle = esc(item.title) || _placeholder('Feature ' + (i + 1));
            const detail = esc(item.detail);
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH}px;background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,0.08);overflow:hidden">
                <svg style="position:absolute;left:20px;top:${(rowH - iconSize) / 2}px;width:${iconSize}px;height:${iconSize}px" viewBox="0 0 ${iconSize} ${iconSize}">${_iconShape(i, iconSize/2, iconSize/2, iconSize*0.8, c)}</svg>
                <div style="position:absolute;left:100px;top:14px;width:780px;font-size:14px;font-weight:700;color:${c}">${iTitle}</div>
                ${detail ? `<div style="position:absolute;left:100px;top:38px;width:780px;font-size:11px;color:#555;line-height:1.4">${detail}</div>` : ''}
            </div>`;
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="items">${rows}</div>
        </div>`;
    },

    comparison_reveal(d, theme, sc) {
        if (theme !== 'editorial_v2') return PREVIEW_RENDERERS.comparison(d, theme, sc);
        const t = esc(d.title || 'Comparison');
        const ll = esc(d.leftLabel || 'Option A');
        const rl = esc(d.rightLabel || 'Option B');
        const leftItems = d.leftItems || [];
        const rightItems = d.rightItems || [];
        let leftCards = '', rightCards = '';
        leftItems.slice(0, 4).forEach((item, i) => {
            const y = 110 + i * 75;
            const text = esc(typeof item === 'string' ? item : (item.text || ''));
            leftCards += `<div style="position:absolute;left:30px;top:${y}px;width:540px;height:65px;background:${EV2.WARM}"><div style="position:absolute;left:0;top:0;width:6px;height:100%;background:${EV2.GOLD}"></div><div style="padding:12px 12px 0 20px;font-size:12px;color:${EV2.CHARCOAL};font-family:${EV2.BF};line-height:1.4">${text}</div></div>`;
        });
        rightItems.slice(0, 4).forEach((item, i) => {
            const y = 110 + i * 75;
            const text = esc(typeof item === 'string' ? item : (item.text || ''));
            rightCards += `<div style="position:absolute;left:600px;top:${y}px;width:365px;height:30px;background:${EV2.LIGHT_BOX}"><div style="position:absolute;left:0;top:0;width:3px;height:100%;background:${EV2.FAINT}"></div><div style="padding:6px 10px 0 14px;font-size:9px;color:${EV2.QUIET};font-family:${EV2.BF}">${text}</div></div>`;
        });
        return `${_ev2White(t)}
            <div style="position:absolute;left:50px;top:42px;font-size:18px;font-weight:700;color:${EV2.DK_GREEN};font-family:${EV2.TF}">${ll}</div>
            <div style="position:absolute;left:50px;top:72px;width:180px;height:2px;background:${EV2.GOLD}"></div>
            <div style="position:absolute;left:600px;top:46px;font-size:11px;font-weight:700;color:${EV2.QUIET};font-family:${EV2.BF}">${rl}</div>
            <div style="position:absolute;left:600px;top:66px;width:100px;height:1px;background:${EV2.RULE_CLR}"></div>
            <div style="position:absolute;left:588px;top:42px;width:1px;height:400px;background:${EV2.RULE_CLR}"></div>
            ${leftCards}${rightCards}
        </div>`;
    },

    process_flow_accordion(d, theme, sc) {
        if (theme !== 'editorial_v2') return PREVIEW_RENDERERS.process_flow(d, theme, sc);
        const t = esc(d.title || 'Process');
        const steps = d.steps || [];
        const n = Math.min(steps.length || 1, 6);
        const cardH = Math.min(65, Math.floor(340 / n));
        let cards = '';
        steps.slice(0, n).forEach((step, i) => {
            const color = EV2.ACCENTS[i % 7];
            const isLeft = i % 2 === 0;
            const y = 85 + i * (cardH + 8);
            const x = isLeft ? 80 : 530;
            const bg = isLeft ? EV2.WARM : EV2.COOL;
            const lbl = esc(step.label || step.title) || '';
            const det = esc(step.detail || '');
            cards += `<div style="position:absolute;left:${x}px;top:${y}px;width:380px;height:${cardH}px;background:${bg}">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${color}"></div>
                <div style="padding:8px 10px 0 16px;font-size:11px;font-weight:700;color:${EV2.CHARCOAL};font-family:${EV2.TF}">${lbl}</div>
                ${det ? `<div style="padding:2px 10px 0 16px;font-size:9px;color:${EV2.MID};font-family:${EV2.BF}">${det}</div>` : ''}
            </div>`;
            cards += `<div style="position:absolute;left:492px;top:${y + cardH / 2 - 12}px;width:24px;height:24px;border-radius:50%;background:${color};color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;font-family:${EV2.BF}">${i + 1}</div>`;
        });
        return `<div class="preview-slide theme-editorial_v2" style="background:#fff">
            ${_ev2TopRule(false)}
            <div style="position:absolute;left:50px;top:18px;font-size:20px;font-weight:700;color:${EV2.CHARCOAL};font-family:${EV2.TF}">${t}</div>
            <div style="position:absolute;left:504px;top:80px;width:2px;height:${n * (cardH + 8)}px;background:${EV2.FAINT}"></div>
            ${cards}
        </div>`;
    },
};
