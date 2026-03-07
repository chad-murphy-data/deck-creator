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

/* ---------- tiny helpers ------------------------------------------------- */

function _accent(sc) {
    return SECTION_COLOR_HEX[sc] || '#368727';
}

function _slickBar(accent) {
    return `<div style="position:absolute;left:0;top:0;width:25px;height:100%;background:${accent}"></div>`;
}

function _slickTitle(title, accent, theme) {
    const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
    const font = theme === 'editorial' ? 'Georgia, serif' : '"Rockwell", Georgia, serif';
    const ruleColor = theme === 'editorial' ? '#044014' : accent;
    return `<div data-field="title" style="position:absolute;left:60px;top:22px;width:900px;font-size:20px;font-weight:700;color:#403F3E;font-family:${font}">${title}</div>
        <div style="position:absolute;left:60px;top:56px;width:160px;height:4px;background:${ruleColor}"></div>`;
}

function _colorfulBar(accent) {
    return `<div class="pv-header-bar" style="background:${accent}"></div>`;
}

function _colorfulTitle(title) {
    return `<div data-field="title" style="position:absolute;left:30px;top:24px;width:940px;font-size:20px;font-weight:700;color:#fff">${title}</div>`;
}

function _slideOpen(theme, bg) {
    const bgStyle = bg ? `background:${bg}` : (theme === 'editorial' ? 'background:#FAF6F0' : 'background:#fff');
    return `<div class="preview-slide theme-${theme}" style="${bgStyle}">`;
}

function _placeholder(text) {
    return `<span style="color:#bbb;font-style:italic">${esc(text)}</span>`;
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
            </div>`;
        }

        /* slick / editorial */
        const bgColor = theme === 'editorial' ? '#FAF6F0' : '#fff';
        const titleFont = theme === 'editorial' ? 'Georgia, serif' : '"Rockwell", Georgia, serif';
        const ruleColor = theme === 'editorial' ? '#044014' : accent;
        return `<div class="preview-slide theme-${theme}" style="background:${bgColor}">
            <div style="position:absolute;left:0;top:0;width:25px;height:100%;background:${accent}"></div>
            <div data-field="title" style="position:absolute;left:90px;top:100px;width:860px;font-size:36px;font-weight:700;color:#403F3E;font-family:${titleFont}">${t}</div>
            <div style="position:absolute;left:90px;top:265px;width:250px;height:4px;background:${ruleColor}"></div>
            ${sub ? `<div data-field="subtitle" style="position:absolute;left:90px;top:285px;font-size:15px;color:#555">${sub}</div>` : ''}
            ${auth ? `<div data-field="author" style="position:absolute;left:90px;top:410px;font-size:12px;color:#555;font-weight:600">${auth}</div>` : ''}
            ${dt ? `<div data-field="date" style="position:absolute;left:90px;top:445px;font-size:11px;color:#555">${dt}</div>` : ''}
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

        /* slick / editorial */
        const bgColor = theme === 'editorial' ? '#044014' : accent;
        return `<div class="preview-slide theme-${theme}" style="background:${bgColor}">
            <div data-field="title" style="position:absolute;left:50px;top:140px;width:900px;font-size:42px;font-weight:700;color:#fff;text-align:center;font-family:${theme === 'editorial' ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${t}</div>
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

        /* slick / editorial */
        let rows = '';
        const rowH = 52;
        const startY = 80;
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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

        /* slick / editorial */
        let rows = '';
        const rowH = 44;
        const startY = 80;
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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

        return `<div class="preview-slide theme-${theme}" style="background:${accent}">
            ${num ? `<div data-field="sectionNumber" style="position:absolute;left:90px;top:100px;font-size:72px;font-weight:700;color:rgba(255,255,255,0.2);font-family:${theme === 'editorial' ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${num}</div>` : ''}
            <div data-field="title" style="position:absolute;left:90px;top:${num ? 200 : 170}px;width:820px;font-size:36px;font-weight:700;color:#fff;font-family:${theme === 'editorial' ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${t}</div>
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
                <div data-field="stat" style="position:absolute;left:50px;top:140px;width:900px;text-align:center;font-size:96px;font-weight:700;color:${accent}">${stat}</div>
                ${headline ? `<div data-field="headline" style="position:absolute;left:50px;top:310px;width:900px;text-align:center;font-size:18px;font-weight:600;color:#333">${headline}</div>` : ''}
                ${detail ? `<div data-field="detail" style="position:absolute;left:100px;top:345px;width:800px;text-align:center;font-size:13px;color:#555">${detail}</div>` : ''}
                ${source ? `<div data-field="source" style="position:absolute;left:50px;top:510px;width:900px;text-align:center;font-size:10px;color:#999">Source: ${source}</div>` : ''}
            </div>`;
        }

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="stat" style="position:absolute;left:60px;top:120px;width:900px;text-align:center;font-size:96px;font-weight:700;color:${accent};font-family:${theme === 'editorial' ? 'Georgia, serif' : '"Rockwell", Georgia, serif'}">${stat}</div>
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div style="position:absolute;left:80px;top:80px;font-size:72px;color:${accent};font-family:Georgia,serif;line-height:1">\u201C</div>
            <div data-field="quote" style="position:absolute;left:100px;top:140px;width:820px;font-size:18px;font-style:italic;color:#333;line-height:1.6">${q || _placeholder('Quote text...')}</div>
            <div style="position:absolute;left:100px;top:380px;width:200px;height:3px;background:${theme === 'editorial' ? '#044014' : accent}"></div>
            ${attr ? `<div data-field="attribution" style="position:absolute;left:100px;top:395px;font-size:13px;font-weight:600;color:#333">\u2014 ${attr}</div>` : ''}
            ${ctx ? `<div data-field="context" style="position:absolute;left:100px;top:415px;font-size:11px;color:#777">${ctx}</div>` : ''}
        </div>`;
    },

    /* --------------------------------------------------------------------- */
    /*  8. COMPARISON                                                        */
    /* --------------------------------------------------------------------- */
    comparison(data, theme, sc) {
        const t = esc(data.title || 'Comparison');
        const lLabel = esc(data.leftLabel || 'Before');
        const rLabel = esc(data.rightLabel || 'After');
        const lItems = data.leftItems && data.leftItems.length ? data.leftItems : ['', '', ''];
        const rItems = data.rightItems && data.rightItems.length ? data.rightItems : ['', '', ''];
        const accent = _accent(sc);

        function listHtml(items, fieldKey) {
            return items.map((it, i) => {
                const text = esc(it) || _placeholder('Item ' + (i + 1));
                return `<div style="padding:5px 0;font-size:12px;color:#444;border-bottom:1px solid rgba(0,0,0,0.06)">${text}</div>`;
            }).join('');
        }

        if (theme === 'colorful') {
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="leftItems" style="position:absolute;left:30px;top:120px;width:440px;height:400px;background:#FFF5E6;border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:#D4A843"></div>
                    <div data-field="leftLabel" style="padding:12px 16px 6px;font-size:14px;font-weight:700;color:#D4A843">${lLabel}</div>
                    <div style="padding:0 16px">${listHtml(lItems, 'leftItems')}</div>
                </div>
                <div style="position:absolute;left:478px;top:290px;width:44px;height:44px;border-radius:50%;background:#eee;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:700;color:#999">vs</div>
                <div data-field="rightItems" style="position:absolute;left:530px;top:120px;width:440px;height:400px;background:#E2F0D9;border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:#368727"></div>
                    <div data-field="rightLabel" style="padding:12px 16px 6px;font-size:14px;font-weight:700;color:#368727">${rLabel}</div>
                    <div style="padding:0 16px">${listHtml(rItems, 'rightItems')}</div>
                </div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
        const lAccent = '#04547C';
        const rAccent = '#368727';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="leftItems" style="position:absolute;left:60px;top:80px;width:430px;height:440px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${lAccent}"></div>
                <div data-field="leftLabel" style="padding:14px 16px 6px 18px;font-size:14px;font-weight:700;color:${lAccent}">${lLabel}</div>
                <div style="padding:0 18px">${listHtml(lItems, 'leftItems')}</div>
            </div>
            <div data-field="rightItems" style="position:absolute;left:510px;top:80px;width:430px;height:440px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${rAccent}"></div>
                <div data-field="rightLabel" style="padding:14px 16px 6px 18px;font-size:14px;font-weight:700;color:${rAccent}">${rLabel}</div>
                <div style="padding:0 18px">${listHtml(rItems, 'rightItems')}</div>
            </div>
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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

        if (theme === 'colorful') {
            const cardW = Math.min(180, Math.floor((940 - (n - 1) * 36) / n));
            let cards = '';
            steps.forEach((step, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const x = 30 + i * (cardW + 36);
                const title = esc(step.title) || _placeholder('Step ' + (i + 1));
                const detail = esc(step.detail);
                cards += `<div style="position:absolute;left:${x}px;top:130px;width:${cardW}px;height:340px;background:${bg};border-radius:8px;overflow:hidden">
                    <div style="height:4px;background:${c}"></div>
                    <div style="position:absolute;left:50%;top:18px;transform:translateX(-50%);width:30px;height:30px;border-radius:50%;background:${c};color:#fff;font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                    <div style="padding:58px 12px 0;font-size:12px;font-weight:600;color:#333;text-align:center">${title}</div>
                    ${detail ? `<div style="padding:6px 12px 0;font-size:10px;color:#666;text-align:center">${detail}</div>` : ''}
                </div>`;
                if (i < n - 1) {
                    cards += `<div style="position:absolute;left:${x + cardW + 4}px;top:290px;width:28px;text-align:center;font-size:20px;color:#bbb">\u2192</div>`;
                }
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="steps">${cards}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
        const cardW = Math.min(180, Math.floor((900 - (n - 1) * 36) / n));
        let cards = '';
        steps.forEach((step, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const x = 60 + i * (cardW + 36);
            const title = esc(step.title) || _placeholder('Step ' + (i + 1));
            const detail = esc(step.detail);
            cards += `<div style="position:absolute;left:${x}px;top:80px;width:${cardW}px;height:380px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:50%;top:16px;transform:translateX(-50%);width:30px;height:30px;border-radius:50%;background:${c};color:#fff;font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center">${i + 1}</div>
                <div style="padding:58px 12px 0;font-size:12px;font-weight:600;color:#333;text-align:center">${title}</div>
                ${detail ? `<div style="padding:6px 12px 0;font-size:10px;color:#666;text-align:center">${detail}</div>` : ''}
            </div>`;
            if (i < n - 1) {
                cards += `<div style="position:absolute;left:${x + cardW + 4}px;top:260px;width:28px;text-align:center;font-size:20px;color:#bbb">\u2192</div>`;
            }
        });
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="steps">${cards}</div>
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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
            const rowH = 52;
            const startY = 120;
            hyps.forEach((h, i) => {
                const c = PREVIEW_COLORS[i % 5];
                const bg = PREVIEW_LIGHTS[i % 5];
                const y = startY + i * (rowH + 10);
                const text = esc(h.text) || _placeholder('Hypothesis ' + (i + 1));
                rows += `<div style="position:absolute;left:30px;top:${y}px;width:940px;height:${rowH}px;background:${bg};border-radius:6px;overflow:hidden">
                    <div style="position:absolute;left:0;top:0;width:940px;height:4px;background:${c}"></div>
                    <div style="position:absolute;left:14px;top:11px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center">H${i + 1}</div>
                    <div style="position:absolute;left:54px;top:16px;font-size:12px;color:#333;width:680px">${text}</div>
                    <div style="position:absolute;right:16px;top:14px">${badgeHtml(h.status)}</div>
                </div>`;
            });
            return `<div class="preview-slide theme-colorful">
                ${_colorfulBar(accent)}
                ${_colorfulTitle(t)}
                <div data-field="hypotheses">${rows}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
        let rows = '';
        const rowH = 52;
        const startY = 80;
        hyps.forEach((h, i) => {
            const c = PREVIEW_COLORS[i % 5];
            const y = startY + i * (rowH + 10);
            const text = esc(h.text) || _placeholder('Hypothesis ' + (i + 1));
            rows += `<div style="position:absolute;left:60px;top:${y}px;width:900px;height:${rowH}px;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${c}"></div>
                <div style="position:absolute;left:14px;top:11px;width:28px;height:28px;border-radius:50%;background:${c};color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center">H${i + 1}</div>
                <div style="position:absolute;left:54px;top:16px;font-size:12px;color:#333;width:660px">${text}</div>
                <div style="position:absolute;right:16px;top:14px">${badgeHtml(h.status)}</div>
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

        const cols = [
            { label: 'What', key: 'what', data: what, color: '#368727' },
            { label: 'So What', key: 'soWhat', data: soWhat, color: '#3880F3' },
            { label: 'Now What', key: 'nowWhat', data: nowWhat, color: '#5B2C8F' },
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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
        /* Same visual as wsn_dense -- shows all three columns */
        return PREVIEW_RENDERERS.wsn_dense(data, theme, sc);
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
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

        /* Build running takeaway list (all items) */
        let listHtml = '';
        takeaways.forEach((ta, i) => {
            const summary = esc(ta.summary) || esc(ta.headline) || _placeholder('Takeaway ' + (i + 1));
            const isCurrent = (i === lastIdx);
            listHtml += `<div style="padding:4px 0;font-size:11px;color:${isCurrent ? '#333' : '#888'};font-weight:${isCurrent ? '600' : '400'}">
                <span style="color:${PREVIEW_COLORS[i % 5]};margin-right:6px">\u2022</span>${summary}
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
                <div style="position:absolute;left:30px;top:312px;font-size:11px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:0.5px">Running Takeaways</div>
                <div style="position:absolute;left:30px;top:334px;width:940px">${listHtml}</div>
            </div>`;
        }

        /* slick / editorial */
        const bg = theme === 'editorial' ? '#FAF6F0' : '#fff';
        return `${_slideOpen(theme, bg)}
            ${_slickBar(accent)}
            ${_slickTitle(t, accent, theme)}
            <div data-field="takeaways" style="position:absolute;left:60px;top:74px;width:900px;height:220px;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
                <div style="position:absolute;left:0;top:0;width:5px;height:100%;background:${PREVIEW_COLORS[lastIdx % 5]}"></div>
                <div style="padding:14px 18px 6px;font-size:15px;font-weight:600;color:#333">${currentHeadline}</div>
                ${currentDetail ? `<div style="padding:0 18px;font-size:12px;color:#555;line-height:1.5">${currentDetail}</div>` : ''}
            </div>
            <div style="position:absolute;left:60px;top:310px;width:900px;height:1px;background:#ddd"></div>
            <div style="position:absolute;left:60px;top:322px;font-size:11px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:0.5px">Running Takeaways</div>
            <div style="position:absolute;left:60px;top:344px;width:900px">${listHtml}</div>
        </div>`;
    },
};
