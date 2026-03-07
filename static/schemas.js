/* schemas.js — field definitions for all 19 slide types */

const SLIDE_CATEGORIES = [
    { id: "opening", label: "Openers & Closers" },
    { id: "structure", label: "Structure" },
    { id: "content", label: "Key Points" },
    { id: "analysis", label: "Analysis & Evidence" },
    { id: "insights", label: "Frameworks & Insights" },
    { id: "discussion", label: "Conclusions" },
];

const SECTION_COLOR_OPTIONS = {
    slick: [
        { value: "green", label: "Green", hex: "#368727" },
        { value: "blue", label: "Blue", hex: "#3880F3" },
        { value: "purple", label: "Purple", hex: "#5B2C8F" },
        { value: "cobalt", label: "Cobalt", hex: "#04547C" },
        { value: "gold", label: "Gold", hex: "#D4A843" },
    ],
    colorful: [
        { value: "green", label: "Green", hex: "#368727" },
        { value: "blue", label: "Blue", hex: "#3880F3" },
        { value: "purple", label: "Purple", hex: "#5B2C8F" },
        { value: "cobalt", label: "Cobalt", hex: "#04547C" },
        { value: "gold", label: "Gold", hex: "#D4A843" },
    ],
    editorial: [
        { value: "green", label: "Green", hex: "#044014" },
        { value: "red", label: "Red", hex: "#C23B22" },
        { value: "teal", label: "Teal", hex: "#1B7A6E" },
        { value: "ochre", label: "Ochre", hex: "#CC7A2E" },
        { value: "slate", label: "Slate", hex: "#5A6A7A" },
        { value: "plum", label: "Plum", hex: "#8E4585" },
    ],
};

const SLIDE_SCHEMAS = {
    title: {
        label: "Title Slide",
        description: "Big title, subtitle, author/date. Use for the opening slide.",
        category: "opening",
        fields: [
            { key: "title", label: "Title", type: "text", default: "Presentation Title", required: true },
            { key: "subtitle", label: "Subtitle", type: "text", default: "" },
            { key: "author", label: "Author", type: "text", default: "" },
            { key: "date", label: "Date", type: "text", default: "" },
        ]
    },
    closer: {
        label: "Closer / Thank You",
        description: "Closing slide with contact info.",
        category: "opening",
        fields: [
            { key: "title", label: "Title", type: "text", default: "Thank You" },
            { key: "subtitle", label: "Subtitle", type: "text", default: "" },
            { key: "contact", label: "Contact Info", type: "text", default: "" },
        ]
    },
    agenda: {
        label: "Agenda",
        description: "Numbered list of topics with optional timing.",
        category: "structure",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Agenda" },
            {
                key: "items", label: "Agenda Items", type: "repeating",
                minItems: 2, maxItems: 8, startItems: 3,
                subfields: [
                    { key: "title", label: "Topic", type: "text", default: "" },
                    { key: "detail", label: "Detail / Time", type: "text", default: "" },
                ]
            },
        ]
    },
    in_brief: {
        label: "In Brief (Bullets)",
        description: "3-5 key bullet points with accent bars.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "In Brief" },
            {
                key: "bullets", label: "Bullet Points", type: "repeating-simple",
                minItems: 2, maxItems: 6, startItems: 3,
                itemLabel: "Point", default: ""
            },
        ]
    },
    stat_callout: {
        label: "Stat Callout",
        description: "One huge number/percentage with headline. Maximum impact.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Key Metric" },
            { key: "stat", label: "Big Number", type: "text", default: "67%" },
            { key: "headline", label: "Headline", type: "text", default: "" },
            { key: "detail", label: "Supporting Detail", type: "textarea", default: "" },
            { key: "source", label: "Source", type: "text", default: "" },
        ]
    },
    quote: {
        label: "Quote",
        description: "Large italic quote with attribution.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "In Their Words" },
            { key: "quote", label: "Quote Text", type: "textarea", default: "" },
            { key: "attribution", label: "Attribution", type: "text", default: "" },
            { key: "context", label: "Context", type: "text", default: "" },
        ]
    },
    comparison: {
        label: "Comparison (Left / Right)",
        description: "Side-by-side cards for before/after or contrast.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Comparison" },
            { key: "leftLabel", label: "Left Label", type: "text", default: "Before" },
            {
                key: "leftItems", label: "Left Items", type: "repeating-simple",
                minItems: 1, maxItems: 5, startItems: 3, itemLabel: "Item", default: ""
            },
            { key: "rightLabel", label: "Right Label", type: "text", default: "After" },
            {
                key: "rightItems", label: "Right Items", type: "repeating-simple",
                minItems: 1, maxItems: 5, startItems: 3, itemLabel: "Item", default: ""
            },
        ]
    },
    text_graph: {
        label: "Text + Graph",
        description: "Text on left, chart placeholder on right.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Title" },
            {
                key: "text", label: "Text Paragraphs", type: "repeating-simple",
                minItems: 1, maxItems: 3, startItems: 2, itemLabel: "Paragraph", default: ""
            },
        ]
    },
    process_flow: {
        label: "Process Flow (Steps)",
        description: "3-5 sequential step cards with arrows.",
        category: "analysis",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Process" },
            {
                key: "steps", label: "Steps", type: "repeating",
                minItems: 2, maxItems: 5, startItems: 3,
                subfields: [
                    { key: "title", label: "Step Title", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
        ]
    },
    matrix: {
        label: "2\u00d72 Matrix",
        description: "Four-quadrant framework with axis labels.",
        category: "analysis",
        layout: "grid-2x2",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Framework" },
            { key: "xAxis", label: "Horizontal Axis \u2192", type: "text", default: "" },
            { key: "yAxis", label: "\u2190 Vertical Axis", type: "text", default: "" },
            {
                key: "quadrants", label: "Quadrants", type: "fixed-list", count: 4,
                subfields: [
                    { key: "label", label: "Label", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ],
                itemLabels: ["Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right"]
            },
        ]
    },
    methods: {
        label: "Methods / Key-Value",
        description: "Label-value pairs stacked vertically.",
        category: "analysis",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Approach" },
            {
                key: "fields", label: "Fields", type: "repeating",
                minItems: 2, maxItems: 6, startItems: 3,
                subfields: [
                    { key: "label", label: "Label", type: "text", default: "" },
                    { key: "value", label: "Value", type: "text", default: "" },
                ]
            },
        ]
    },
    hypotheses: {
        label: "Hypotheses",
        description: "Numbered hypotheses with confirmed/rejected badges.",
        category: "analysis",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Hypotheses" },
            {
                key: "hypotheses", label: "Hypotheses", type: "repeating",
                minItems: 1, maxItems: 5, startItems: 3,
                subfields: [
                    { key: "text", label: "Hypothesis", type: "text", default: "" },
                    { key: "status", label: "Status", type: "select", options: ["", "Confirmed", "Rejected", "Partial"], default: "" },
                ]
            },
        ]
    },
    wsn_dense: {
        label: "What / So What / Now What",
        description: "Three-column single-slide insight.",
        category: "insights",
        layout: "wsn",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Key Finding" },
            {
                key: "what", label: "What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
            {
                key: "soWhat", label: "So What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
            {
                key: "nowWhat", label: "Now What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
        ]
    },
    wsn_reveal: {
        label: "WSN Progressive Reveal",
        description: "Builds across 3 slides: What \u2192 So What \u2192 Now What.",
        category: "insights",
        layout: "wsn",
        generatesMultipleSlides: 3,
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Key Finding" },
            {
                key: "what", label: "What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
            {
                key: "soWhat", label: "So What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
            {
                key: "nowWhat", label: "Now What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
        ]
    },
    findings_recs: {
        label: "Findings & Recommendations",
        description: "Paired finding \u2192 recommendation cards (up to 5).",
        category: "insights",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Findings & Recommendations" },
            {
                key: "items", label: "Items", type: "repeating",
                minItems: 1, maxItems: 5, startItems: 3,
                subfields: [
                    { key: "finding", label: "Finding", type: "text", default: "" },
                    { key: "recommendation", label: "Recommendation", type: "text", default: "" },
                ]
            },
        ]
    },
    findings_recs_dense: {
        label: "Findings & Recs (Dense)",
        description: "Compact paired rows \u2014 fits up to 8.",
        category: "insights",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Complete Findings" },
            {
                key: "items", label: "Items", type: "repeating",
                minItems: 1, maxItems: 8, startItems: 4,
                subfields: [
                    { key: "finding", label: "Finding", type: "text", default: "" },
                    { key: "recommendation", label: "Recommendation", type: "text", default: "" },
                ]
            },
        ]
    },
    open_questions: {
        label: "Open Questions",
        description: "2\u00d72 grid of question cards.",
        category: "discussion",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Open Questions" },
            {
                key: "questions", label: "Questions", type: "repeating-simple",
                minItems: 1, maxItems: 4, startItems: 3,
                itemLabel: "Question", default: ""
            },
        ]
    },
    progressive_reveal: {
        label: "Progressive Reveal",
        description: "One slide per takeaway with running summary list.",
        category: "discussion",
        generatesMultipleSlides: "N",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Building the Picture" },
            {
                key: "takeaways", label: "Takeaways", type: "repeating",
                minItems: 2, maxItems: 5, startItems: 2,
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                    { key: "summary", label: "Short Summary (for running list)", type: "text", default: "" },
                ]
            },
        ]
    },
};

/* Utility: build default data for a slide type */
function buildDefaultData(type) {
    const schema = SLIDE_SCHEMAS[type];
    if (!schema) return {};
    const data = {};
    for (const field of schema.fields) {
        if (field.type === "repeating") {
            const n = field.startItems || field.minItems || 1;
            const items = [];
            for (let i = 0; i < n; i++) {
                const item = {};
                for (const sf of field.subfields) item[sf.key] = sf.default || "";
                items.push(item);
            }
            data[field.key] = items;
        } else if (field.type === "repeating-simple") {
            const n = field.startItems || field.minItems || 1;
            data[field.key] = Array(n).fill(field.default || "");
        } else if (field.type === "group") {
            const group = {};
            for (const sf of field.subfields) group[sf.key] = sf.default || "";
            data[field.key] = group;
        } else if (field.type === "fixed-list") {
            const items = [];
            for (let i = 0; i < field.count; i++) {
                const item = {};
                for (const sf of field.subfields) item[sf.key] = sf.default || "";
                items.push(item);
            }
            data[field.key] = items;
        } else {
            data[field.key] = field.default !== undefined ? field.default : "";
        }
    }
    return data;
}
