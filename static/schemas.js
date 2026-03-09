/* schemas.js — field definitions for all 19 slide types */

const SLIDE_CATEGORIES = [
    { id: "opening", label: "Openers & Closers" },
    { id: "structure", label: "Structure" },
    { id: "content", label: "Key Points" },
    { id: "analysis", label: "Analysis & Evidence" },
    { id: "insights", label: "Frameworks & Insights" },
    { id: "personas", label: "Personas" },
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
    bold: [
        { value: "green", label: "Green", hex: "#368727" },
        { value: "blue", label: "Blue", hex: "#3880F3" },
        { value: "purple", label: "Purple", hex: "#5B2C8F" },
        { value: "cobalt", label: "Cobalt", hex: "#04547C" },
        { value: "gold", label: "Gold", hex: "#D4A843" },
    ],
    noir: [
        { value: "green", label: "Green", hex: "#368727" },
        { value: "blue", label: "Blue", hex: "#3880F3" },
        { value: "purple", label: "Purple", hex: "#5B2C8F" },
        { value: "cobalt", label: "Cobalt", hex: "#04547C" },
        { value: "gold", label: "Gold", hex: "#D4A843" },
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
            { key: "imagePath", label: "Image (optional)", type: "text", default: "" },
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
    process_flow_reveal: {
        label: "Process Flow (Reveal)",
        description: "Builds step-by-step across multiple slides.",
        category: "analysis",
        generatesMultipleSlides: "N",
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
                key: "labels", label: "Column Labels", type: "repeating-simple",
                minItems: 3, maxItems: 3, startItems: 3,
                itemLabel: "Label", default: ""
            },
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
                key: "labels", label: "Column Labels", type: "repeating-simple",
                minItems: 3, maxItems: 3, startItems: 3,
                itemLabel: "Label", default: ""
            },
            {
                key: "what", label: "What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                    { key: "summary", label: "Short Summary", type: "text", default: "" },
                ]
            },
            {
                key: "soWhat", label: "So What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                    { key: "summary", label: "Short Summary", type: "text", default: "" },
                ]
            },
            {
                key: "nowWhat", label: "Now What", type: "group",
                subfields: [
                    { key: "headline", label: "Headline", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                    { key: "summary", label: "Short Summary", type: "text", default: "" },
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
                minItems: 1, maxItems: 5, startItems: 4,
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
    timeline: {
        label: "Timeline",
        description: "Horizontal timeline with milestones and status indicators.",
        category: "structure",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Timeline" },
            {
                key: "milestones", label: "Milestones", type: "repeating",
                minItems: 2, maxItems: 6, startItems: 3,
                subfields: [
                    { key: "date", label: "Date", type: "text", default: "" },
                    { key: "title", label: "Title", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "text", default: "" },
                    { key: "status", label: "Status", type: "select", options: ["upcoming", "current", "complete"], default: "upcoming" },
                ]
            },
        ]
    },
    data_table: {
        label: "Data Table",
        description: "Tabular data with headers and optional column highlighting.",
        category: "analysis",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Data" },
            {
                key: "headers", label: "Column Headers", type: "repeating-simple",
                minItems: 2, maxItems: 6, startItems: 3,
                itemLabel: "Header", default: ""
            },
            {
                key: "rows", label: "Rows", type: "repeating",
                minItems: 1, maxItems: 10, startItems: 3,
                subfields: [
                    { key: "0", label: "Col 1", type: "text", default: "" },
                    { key: "1", label: "Col 2", type: "text", default: "" },
                    { key: "2", label: "Col 3", type: "text", default: "" },
                ]
            },
            { key: "highlightCol", label: "Highlight Column (0-based index)", type: "text", default: "" },
            { key: "note", label: "Footnote", type: "text", default: "" },
        ]
    },
    multi_stat: {
        label: "Multi-Stat",
        description: "2-4 key metrics side by side with values and labels.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Key Metrics" },
            {
                key: "stats", label: "Statistics", type: "repeating",
                minItems: 2, maxItems: 4, startItems: 3,
                subfields: [
                    { key: "value", label: "Value", type: "text", default: "" },
                    { key: "label", label: "Label", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "text", default: "" },
                ]
            },
            { key: "source", label: "Source", type: "text", default: "" },
        ]
    },
    persona: {
        label: "Persona",
        description: "User persona with name, archetype, traits, and strategy.",
        category: "personas",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Persona" },
            { key: "name", label: "Name", type: "text", default: "" },
            { key: "archetype", label: "Archetype", type: "text", default: "" },
            {
                key: "traits", label: "Traits", type: "repeating-simple",
                minItems: 1, maxItems: 5, startItems: 3,
                itemLabel: "Trait", default: ""
            },
            { key: "strategy", label: "Strategy", type: "textarea", default: "" },
            { key: "detail", label: "Additional Detail", type: "textarea", default: "" },
        ]
    },
    risk_tradeoff: {
        label: "Risk / Reward",
        description: "Side-by-side risks and rewards with severity levels.",
        category: "analysis",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Risk & Reward" },
            {
                key: "risks", label: "Risks", type: "repeating",
                minItems: 1, maxItems: 5, startItems: 2,
                subfields: [
                    { key: "label", label: "Risk", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "text", default: "" },
                    { key: "severity", label: "Severity", type: "select", options: ["low", "medium", "high"], default: "medium" },
                ]
            },
            {
                key: "rewards", label: "Rewards", type: "repeating",
                minItems: 1, maxItems: 5, startItems: 2,
                subfields: [
                    { key: "label", label: "Reward", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "text", default: "" },
                ]
            },
        ]
    },
    appendix: {
        label: "Appendix",
        description: "Reference material organized in labeled sections.",
        category: "discussion",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Appendix" },
            {
                key: "sections", label: "Sections", type: "repeating",
                minItems: 1, maxItems: 4, startItems: 2,
                subfields: [
                    { key: "label", label: "Section Label", type: "text", default: "" },
                    { key: "content", label: "Content", type: "textarea", default: "" },
                ]
            },
        ]
    },
    before_after: {
        label: "Before / After",
        description: "Transformation view with before state, intervention, and after state.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Transformation" },
            {
                key: "before", label: "Before", type: "group",
                subfields: [
                    { key: "label", label: "Label", type: "text", default: "Before" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
            { key: "intervention", label: "Intervention / Change", type: "textarea", default: "" },
            {
                key: "after", label: "After", type: "group",
                subfields: [
                    { key: "label", label: "Label", type: "text", default: "After" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
        ]
    },
    summary: {
        label: "Summary",
        description: "Column-based summary with headings and bullet points.",
        category: "discussion",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Summary" },
            {
                key: "sections", label: "Columns", type: "repeating",
                minItems: 2, maxItems: 4, startItems: 3,
                subfields: [
                    { key: "heading", label: "Heading", type: "text", default: "" },
                    {
                        key: "points", label: "Points", type: "repeating-simple",
                        minItems: 1, maxItems: 5, startItems: 2,
                        itemLabel: "Point", default: ""
                    },
                ]
            },
        ]
    },
    quote_full: {
        label: "Full-Bleed Quote",
        description: "Dramatic full-slide quote on colored background.",
        category: "content",
        fields: [
            { key: "quote", label: "Quote Text", type: "textarea", default: "" },
            { key: "attribution", label: "Attribution", type: "text", default: "" },
            { key: "context", label: "Context", type: "text", default: "" },
        ]
    },
    stat_hero: {
        label: "Stat Hero",
        description: "One hero stat with supporting metrics below.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Key Metric" },
            {
                key: "hero", label: "Hero Stat", type: "group",
                subfields: [
                    { key: "value", label: "Value", type: "text", default: "" },
                    { key: "label", label: "Label", type: "text", default: "" },
                ]
            },
            {
                key: "supporting", label: "Supporting Stats", type: "repeating",
                minItems: 0, maxItems: 4, startItems: 2,
                subfields: [
                    { key: "value", label: "Value", type: "text", default: "" },
                    { key: "label", label: "Label", type: "text", default: "" },
                ]
            },
            { key: "source", label: "Source", type: "text", default: "" },
        ]
    },
    in_brief_featured: {
        label: "In Brief (Featured)",
        description: "One featured callout with supporting bullet points.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "In Brief" },
            { key: "featured", label: "Featured Point", type: "textarea", default: "" },
            {
                key: "supporting", label: "Supporting Points", type: "repeating-simple",
                minItems: 1, maxItems: 4, startItems: 2,
                itemLabel: "Point", default: ""
            },
        ]
    },
    persona_duo: {
        label: "Persona Duo",
        description: "Side-by-side comparison of two personas.",
        category: "personas",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Archetype Comparison" },
            {
                key: "personas", label: "Personas", type: "fixed-list", count: 2,
                subfields: [
                    { key: "name", label: "Name", type: "text", default: "" },
                    { key: "archetype", label: "Archetype", type: "text", default: "" },
                    {
                        key: "traits", label: "Traits", type: "repeating-simple",
                        minItems: 1, maxItems: 4, startItems: 2,
                        itemLabel: "Trait", default: ""
                    },
                    { key: "strategy", label: "Strategy", type: "textarea", default: "" },
                ],
                itemLabels: ["Persona A", "Persona B"]
            },
        ]
    },
    process_flow_vertical: {
        label: "Process Flow (Vertical)",
        description: "Vertical step-by-step flow with arrows between steps.",
        category: "analysis",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Process" },
            {
                key: "steps", label: "Steps", type: "repeating",
                minItems: 2, maxItems: 3, startItems: 3,
                subfields: [
                    { key: "title", label: "Step Title", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
        ]
    },
    text_cards: {
        label: "Text Cards",
        description: "Grid of cards with titles and details.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Key Points" },
            {
                key: "items", label: "Cards", type: "repeating",
                minItems: 2, maxItems: 6, startItems: 4,
                subfields: [
                    { key: "title", label: "Card Title", type: "text", default: "" },
                    { key: "detail", label: "Detail", type: "textarea", default: "" },
                ]
            },
        ]
    },
    text_columns: {
        label: "Text Columns",
        description: "2-3 columns of text with optional headings.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Overview" },
            {
                key: "columns", label: "Columns", type: "repeating",
                minItems: 2, maxItems: 3, startItems: 2,
                subfields: [
                    { key: "heading", label: "Heading", type: "text", default: "" },
                    { key: "body", label: "Body", type: "textarea", default: "" },
                ]
            },
        ]
    },
    text_narrative: {
        label: "Text Narrative",
        description: "Featured lead paragraph with flowing body text.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Overview" },
            { key: "lede", label: "Lead Paragraph", type: "textarea", default: "" },
            { key: "body", label: "Body Text", type: "textarea", default: "" },
        ]
    },
    text_nested: {
        label: "Text Nested",
        description: "Hierarchical content with colored label blocks and child items.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Detail" },
            {
                key: "items", label: "Sections", type: "repeating",
                minItems: 2, maxItems: 4, startItems: 3,
                subfields: [
                    { key: "text", label: "Section Label", type: "text", default: "" },
                    {
                        key: "children", label: "Children", type: "repeating-simple",
                        minItems: 1, maxItems: 4, startItems: 2,
                        itemLabel: "Item", default: ""
                    },
                ]
            },
        ]
    },
    text_split: {
        label: "Text Split",
        description: "Headline on left, supporting points on right.",
        category: "content",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Key Point" },
            { key: "headline", label: "Headline", type: "textarea", default: "" },
            { key: "detail", label: "Detail", type: "textarea", default: "" },
            {
                key: "points", label: "Supporting Points", type: "repeating-simple",
                minItems: 1, maxItems: 6, startItems: 3,
                itemLabel: "Point", default: ""
            },
        ]
    },
    text_annotated: {
        label: "Text Annotated",
        description: "Labeled rows with colored category tags and descriptions.",
        category: "analysis",
        fields: [
            { key: "title", label: "Slide Title", type: "text", default: "Analysis" },
            {
                key: "items", label: "Items", type: "repeating",
                minItems: 2, maxItems: 5, startItems: 3,
                subfields: [
                    { key: "label", label: "Label/Tag", type: "text", default: "" },
                    { key: "text", label: "Description", type: "textarea", default: "" },
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
    if (type === 'wsn_dense' && data.labels && data.labels.every(l => l === '')) {
        data.labels = ['What', 'So What', 'Now What'];
    }
    if (type === 'wsn_reveal' && data.labels && data.labels.every(l => l === '')) {
        data.labels = ['What', 'So What', 'Now What'];
    }
    return data;
}
