"""
Everything I Learned About Storytelling I Learned from Gilmore Girls
- Content mapped to all 38 slide builder types (format-corrected)
"""

SLIDES = [
    ("title", {
        "title": "Everything I Learned About\nStorytelling I Learned\nfrom Gilmore Girls",
        "subtitle": "A framework for narrative craft, disguised as a show about coffee",
        "author": "Stars Hollow School of Communication",
        "date": "2025",
    }),

    ("in_brief", {
        "title": "Why Gilmore Girls Is a Storytelling Masterclass",
        "bullets": [
            "153 episodes averaging 8,000+ words of dialogue each — more than most films",
            "Pioneered the 'walk and talk' as a pacing device for television",
            "Built an entire fictional economy of recurring characters across 7 seasons",
            "Proved that verbal density and emotional depth aren't trade-offs",
            "Created one of TV's most debated narrative choices with the final four words",
        ],
    }),

    ("section_divider", {
        "title": "Act I: Pacing & Rhythm",
        "subtitle": "What 300 words per minute teaches you about holding attention",
        "sectionNumber": "01",
    }),

    ("stat_callout", {
        "title": "The Gilmore Speed Advantage",
        "stat": "160%",
        "headline": "Faster than average TV dialogue",
        "detail": "Gilmore Girls scripts ran roughly 75-80 pages per episode vs. the standard 45-50. Amy Sherman-Palladino wrote dialogue at ~300 words/minute compared to the TV average of ~120. This forced actors to rehearse like stage performers — the pacing itself became a storytelling tool.",
        "source": "Vanity Fair, 2016",
    }),

    ("quote", {
        "title": "On Pacing as Character",
        "quote": "I speak fast. I think fast. I eat fast. I have to because there's just so much to do and so little time.",
        "attribution": "Lorelai Gilmore, Season 1",
        "context": "The speed of Gilmore dialogue isn't a stylistic quirk — it's characterization. Lorelai's rapid-fire delivery signals anxiety masked as confidence. When she slows down, something is wrong.",
    }),

    ("comparison", {
        "title": "Two Schools of TV Dialogue",
        "leftLabel": "Gilmore Girls (Density)",
        "leftItems": [
            "Dialogue carries plot, character, and setting simultaneously",
            "References function as shorthand for emotional states",
            "Silence is rare and therefore devastating when used",
            "The audience must keep up — no hand-holding",
        ],
        "rightLabel": "Standard TV (Clarity)",
        "rightItems": [
            "Dialogue serves one function per exchange",
            "Exposition is delivered explicitly",
            "Silence is common and used for pacing",
            "The audience is guided through each beat",
        ],
    }),

    ("text_graph", {
        "title": "Dialogue Density Across TV Eras",
        "text": "Gilmore Girls pioneered a style that later influenced Bunheads, The Marvelous Mrs. Maisel, and Succession. The trend toward density has accelerated, but GG remains the benchmark for verbal velocity in a drama format.",
        "chartType": "bar",
        "chartData": [
            {"label": "Friends", "value": 110},
            {"label": "West Wing", "value": 185},
            {"label": "Gilmore Girls", "value": 300},
            {"label": "Mrs. Maisel", "value": 280},
            {"label": "Succession", "value": 200},
        ],
        "note": "Approximate words per minute of dialogue",
    }),

    ("process_flow", {
        "title": "How a Gilmore Rant Builds to a Punchline",
        "steps": [
            {"label": "Hook", "detail": "Start with an absurd premise or observation that seems throwaway"},
            {"label": "Escalate", "detail": "Layer references and tangents that appear to wander off-topic"},
            {"label": "Callback", "detail": "Circle back to the original premise with a twist"},
            {"label": "Land", "detail": "Deliver the emotional truth underneath the comedy"},
        ],
    }),

    ("matrix", {
        "title": "The Stars Hollow Character Grid",
        "quadrants": [
            {"label": "High Warmth / High Chaos", "content": "Lorelai, Sookie, Babette — charm as a weapon, unpredictable but always landing on their feet"},
            {"label": "High Warmth / Low Chaos", "content": "Luke, Lane, Rory (early) — steady presences who anchor scenes and give others room to perform"},
            {"label": "Low Warmth / High Chaos", "content": "Paris, Emily, Taylor — friction generators who create the conflict that drives episodes"},
            {"label": "Low Warmth / Low Chaos", "content": "Richard, Headmaster Charleston — institutional weight, used sparingly for maximum dramatic impact"},
        ],
    }),

    ("methods", {
        "title": "Sherman-Palladino's Storytelling Toolkit",
        "fields": [
            {"label": "Walk & Talk", "text": "Camera follows characters through physical spaces, using movement to externalize internal momentum. Borrowed from Sorkin but made domestic."},
            {"label": "Pop Culture Shorthand", "text": "References to films, books, and music function as emotional vocabulary — characters say 'Casablanca' when they mean 'doomed love.'"},
            {"label": "Town as Greek Chorus", "text": "Stars Hollow residents comment on, interfere with, and mirror the main characters' arcs. Every town meeting is exposition disguised as comedy."},
            {"label": "The Friday Dinner", "text": "A recurring structural device that forces characters into confined spaces where avoidance is impossible."},
        ],
    }),

    ("hypotheses", {
        "title": "Storytelling Hypotheses from Stars Hollow",
        "hypotheses": [
            {"text": "Audiences will tolerate complexity if the emotional throughline is clear", "status": "Confirmed"},
            {"text": "Recurring settings reduce cognitive load, freeing attention for dialogue", "status": "Confirmed"},
            {"text": "Cultural references age well if they serve character rather than currency", "status": ""},
            {"text": "A strong ensemble can sustain a show even when the leads falter", "status": "Confirmed"},
        ],
    }),

    ("wsn_dense", {
        "title": "The Friday Dinner Problem",
        "what": {"headline": "The Obligatory Gathering", "detail": "Every Friday, Rory and Lorelai attend dinner at Emily and Richard's house in Hartford. These dinners are tense, performative, and loaded with subtext. Established in the pilot as a condition of Rory's Chilton tuition."},
        "soWhat": {"headline": "Structure Creates Drama", "detail": "The Friday dinner forces the central conflict (class, independence, obligation) into a confined space on a reliable schedule. Every major revelation and turning point either happens at or is triggered by a Friday dinner."},
        "nowWhat": {"headline": "Use Recurring Obligation as Engine", "detail": "When characters can't leave, they have to deal. The constraint creates the drama. You don't need external plot devices if the structure itself generates tension."},
    }),

    ("wsn_reveal", {
        "title": "Luke's Diner: Setting as Character",
        "what": {"headline": "A Space with Its Own Rules", "detail": "Luke's Diner appears in 93% of all episodes. It's where Lorelai gets coffee, where town gossip circulates, and where most casual conversations happen. Unwritten rules (no cell phones, counter seats for regulars) function like character traits."},
        "soWhat": {"headline": "Repetition Earns Emotional Weight", "detail": "The diner isn't a backdrop. Scenes set at Luke's establish normalcy, so when something disrupts the routine (a fight, a closure, a renovation), the audience feels it viscerally."},
        "nowWhat": {"headline": "Build a Home Base, Then Disrupt It", "detail": "Give your narrative anchor rules, regulars, and rituals. When you need dramatic impact, disrupt those rituals rather than introducing something new."},
    }),

    ("findings_recs", {
        "title": "What the Revival Got Right and Wrong",
        "items": [
            {"finding": "The musical episode ran 12 minutes — audiences checked out", "recommendation": "Self-indulgence has a cost. Cut the bits you love most if they don't serve the story."},
            {"finding": "Emily's grief arc was the revival's strongest storyline", "recommendation": "Characters with the most to lose create the most compelling arcs."},
            {"finding": "Rory's career stagnation felt unearned after 7 seasons of promise", "recommendation": "Subverting expectations only works if the new direction is as interesting as the old one."},
        ],
    }),

    ("findings_recs_dense", {
        "title": "Lessons from the Gilmore Writers' Room",
        "items": [
            {"finding": "Scripts were locked — no ad-libbing allowed", "recommendation": "Precision in language matters. When every word is chosen, dialogue becomes the performance."},
            {"finding": "Each episode was written as a mini-play with 3 acts", "recommendation": "TV episodes that feel like complete stories within a series arc retain audiences better."},
            {"finding": "Rehearsals ran like theatre — full table reads before shooting", "recommendation": "Invest in preparation. The polish shows even if the audience can't articulate why."},
            {"finding": "Pop culture references were never explained in-script", "recommendation": "Trust your audience. Explaining the joke kills it."},
            {"finding": "Town meetings were used to reset or advance the B-plot", "recommendation": "Build a device that lets you shift gears without a jarring transition."},
        ],
    }),

    ("open_questions", {
        "title": "Unresolved Questions in Narrative Craft",
        "questions": [
            "Does the 'final four words' structure (planned ending held for years) improve or constrain a long-running story?",
            "Can verbal density work outside of comedy-drama? Would a thriller sustain 300 wpm?",
            "Is Stars Hollow's insularity a feature or a bug for modern storytelling?",
            "When does an ensemble become too large to sustain meaningful arcs?",
        ],
    }),

    ("agenda", {
        "title": "A Storytelling Workshop in 7 Seasons",
        "items": [
            {"title": "S1: Establishing Voice", "detail": "Build a world in 10 minutes"},
            {"title": "S2: Deepening Conflict", "detail": "The Jess arc as controlled disruption"},
            {"title": "S3: Peak Ensemble", "detail": "Every subplot connects to the spine"},
            {"title": "S4: The College Problem", "detail": "Evolve without losing the core"},
            {"title": "S5: Raising the Floor", "detail": "Worst episodes still good"},
            {"title": "S6: Losing the Thread", "detail": "Vision vs. network divergence"},
            {"title": "S7: The Handoff", "detail": "Can a story survive without its voice?"},
        ],
    }),

    ("progressive_reveal", {
        "title": "The Real Lesson of Gilmore Girls",
        "takeaways": [
            {"headline": "It's not about the coffee or the pop culture references", "detail": "Those are the surface pleasures — the hooks that get you in the door. But they're not what keeps people rewatching 20 years later.", "summary": "Not about the surface pleasures"},
            {"headline": "It's not even about mother-daughter relationships", "detail": "Though that's closer. The Lorelai-Rory dynamic is the spine, but it's not the theme.", "summary": "Not about the family structure"},
            {"headline": "It's about the gap between performance and truth", "detail": "Every major character in Gilmore Girls is performing a version of themselves that hides what they actually feel.", "summary": "The gap between who we perform and who we are"},
            {"headline": "Lorelai performs independence; she's terrified of needing anyone", "detail": "Her jokes are armor. Her speed is avoidance. Every relationship she sabotages is proof.", "summary": "Lorelai: independence masking fear"},
            {"headline": "The storytelling lesson: build the performance first", "detail": "Build your character's public self first, then slowly reveal what's underneath. The revelation is the drama.", "summary": "Build the mask, then slowly remove it"},
        ],
    }),

    ("closer", {
        "title": "Where You Lead, I Will Follow",
        "subtitle": "The best stories don't tell you what to think. They walk fast enough that you have to run to keep up — and then you realize you've been somewhere important the whole time.",
        "contact": "Stars Hollow Gazette | Friday Dinners by Appointment Only",
    }),

    ("timeline", {
        "title": "Gilmore Girls: A Narrative Timeline",
        "milestones": [
            {"date": "Oct 2000", "title": "Pilot Airs", "detail": "WB debut; establishes the Friday dinner structure in 44 minutes", "status": "complete"},
            {"date": "2001", "title": "Jess Arrives", "detail": "The first true narrative disruptor", "status": "complete"},
            {"date": "2003", "title": "Rory Enters Yale", "detail": "Setting shift tests the show's voice", "status": "complete"},
            {"date": "2005", "title": "Rory Drops Out", "detail": "The most divisive arc", "status": "complete"},
            {"date": "2006", "title": "ASP Departs", "detail": "Sherman-Palladino leaves after S6", "status": "complete"},
            {"date": "2016", "title": "Revival Airs", "detail": "The final four words delivered", "status": "current"},
        ],
    }),

    ("data_table", {
        "title": "Gilmore Girls by the Numbers",
        "headers": ["Metric", "Value", "Context"],
        "rows": [
            ["Episodes", "153 + 4", "7 seasons plus Netflix revival"],
            ["Avg Script Length", "75-80 pages", "Standard TV: 45-50 pages"],
            ["Dialogue Speed", "~300 wpm", "TV average: ~120 wpm"],
            ["Recurring Characters", "50+", "Named characters with arcs"],
            ["Pop Culture Refs/Ep", "~30", "Verified by fan databases"],
            ["Friday Dinners", "~100", "Central structural device"],
        ],
        "highlightCol": 1,
        "note": "Sources: Vanity Fair, The Gilmore Girls Companion, fan wikis",
    }),

    ("multi_stat", {
        "title": "The Cultural Footprint",
        "stats": [
            {"value": "7", "label": "Seasons on the WB/CW"},
            {"value": "4", "label": "Revival episodes on Netflix"},
            {"value": "50+", "label": "Recurring named characters"},
            {"value": "30+", "label": "Pop culture refs per episode"},
        ],
        "source": "The Gilmore Girls Companion & fan databases",
    }),

    ("persona", {
        "title": "Character as Storytelling Device",
        "name": "Emily Gilmore",
        "archetype": "The Antagonist Who's Right",
        "traits": ["Exacting standards", "Weaponized etiquette", "Fierce loyalty masked as control", "Devastating one-liners"],
        "detail": "Emily is the show's most complex character because she's simultaneously the obstacle and the mirror. Every criticism she levels at Lorelai is partially valid, and every act of control is rooted in genuine love. She makes the audience uncomfortable because she forces us to ask: what if the villain has a point?",
        "strategy": "Use Emily's archetype when you need an antagonist audiences argue about, not root against.",
    }),

    ("risk_tradeoff", {
        "title": "The Rory Problem: Risks of a Perfect Protagonist",
        "risks": [
            {"label": "Audience Fatigue", "detail": "A character who always succeeds gives the audience nothing to worry about"},
            {"label": "Stakes Deflation", "detail": "When failure finally comes (the yacht theft), it feels jarring rather than earned"},
            {"label": "Protagonist Immunity", "detail": "Other characters are more interesting because they face real consequences"},
        ],
        "rewards": [
            {"label": "Aspirational Investment", "detail": "Audiences project their own ambitions onto Rory's trajectory"},
            {"label": "Long-Arc Payoff", "detail": "The eventual fall hits harder because of the height she fell from"},
            {"label": "Ensemble Freedom", "detail": "A stable protagonist lets supporting characters take bigger risks"},
        ],
    }),

    ("appendix", {
        "title": "Further Reading & References",
        "label": "Appendix A",
        "sections": [
            {"label": "Books", "content": "The Gilmore Girls Companion (2010); Coffee at Luke's (2007); Feminist Subtext in Gilmore Girls (academic collection, 2019)"},
            {"label": "Key Episodes", "content": "S1E1 Pilot — world-building; S2E5 Nick & Nora — peak walk-and-talk; S5E13 Wedding Bell Blues — pressure cooker; S6E22 Partings — cost of lost voice"},
            {"label": "Related Shows", "content": "The West Wing (pacing), Bunheads (same creator), The Marvelous Mrs. Maisel (spiritual successor), Succession (verbal density in drama)"},
        ],
    }),

    ("before_after", {
        "title": "Before and After: The Sherman-Palladino Effect",
        "before": {"label": "Season 7 (Rosenthal)", "detail": "Conventional pacing, resolved conflicts neatly, dialogue slowed to standard TV speed, town characters became background decoration, emotional beats stated explicitly"},
        "intervention": "Sherman-Palladino departs after Season 6 due to contract disputes",
        "after": {"label": "Seasons 1-6 (ASP)", "detail": "Relentless pacing, conflicts lingered across seasons, dialogue at 300 wpm, every townie served narrative purpose, emotional beats embedded in subtext and pop culture references"},
    }),

    ("summary", {
        "title": "Key Takeaways for Storytellers",
        "heading": "Seven Lessons from Stars Hollow",
        "sections": [
            {"heading": "Pace > Plot", "points": ["Rhythm and velocity can compensate for thin plot if the character work is strong"]},
            {"heading": "Build Recurring Structures", "points": ["Friday dinners, town meetings, and Luke's Diner aren't repetitive — they're load-bearing"]},
            {"heading": "Trust Your Audience", "points": ["Don't explain the reference. The ones who get it feel like insiders."]},
        ],
        "points": [
            "Constraint creates drama — the best scenes happen when characters can't leave",
            "Your antagonist should be partially right",
            "Ensemble isn't decoration — every recurring character should serve the story's spine",
            "Know your ending, but hold it loosely",
        ],
    }),

    ("quote_full", {
        "quote": "I live in two worlds. One is a world of books. I've been a resident of Faulkner's Yoknapatawpha County, hunted whales with Ahab, fought alongside Napoleon. The other is Stars Hollow.",
        "attribution": "Rory Gilmore, Chilton Valedictorian Speech, S3",
        "context": "This speech works because it collapses the distance between fiction and life — the same move Gilmore Girls makes in every episode.",
    }),

    ("stat_hero", {
        "title": "The Speed of Connection",
        "hero": {"value": "300", "label": "Words per minute — the speed at which the Gilmores spoke and audiences learned to listen"},
        "supporting": [
            {"value": "80", "label": "Script pages per episode"},
            {"value": "153", "label": "Episodes across 7 seasons"},
            {"value": "50+", "label": "Recurring characters"},
        ],
        "source": "Script analyses; The Gilmore Girls Companion",
    }),

    ("in_brief_featured", {
        "title": "The Core Storytelling Principle",
        "featured": "Make them keep up. The fastest way to respect your audience is to refuse to slow down for them — and then reward them for staying with you.",
        "supporting": [
            "Gilmore Girls never paused to explain a reference",
            "The show trusted viewers to catch up or rewatch",
            "This created a cult following that still dissects episodes 20 years later",
            "The lesson: density is a form of generosity, not gatekeeping",
        ],
    }),

    ("persona_duo", {
        "title": "Two Models of Storytelling Leadership",
        "personas": [
            {
                "name": "Lorelai Gilmore",
                "archetype": "The Improviser",
                "traits": ["Leads with charm and instinct", "Avoids planning until forced", "Turns every crisis into a bit", "Builds loyalty through warmth"],
                "detail": "Lorelai's leadership style is performative spontaneity — she appears to wing it, but her improvisations are built on deep knowledge of her people.",
            },
            {
                "name": "Emily Gilmore",
                "archetype": "The Architect",
                "traits": ["Leads with structure and expectation", "Plans every detail in advance", "Controls environment to control outcomes", "Builds loyalty through obligation"],
                "detail": "Emily's leadership style is institutional — she believes that if the structure is right, the people within it will perform.",
            },
        ],
    }),

    ("process_flow_vertical", {
        "title": "Building a Scene the Gilmore Way",
        "steps": [
            {"title": "Set the Table", "detail": "Establish the physical space and who's in it"},
            {"title": "Start Mid-Conversation", "detail": "Never begin at the beginning. Drop in where it's interesting."},
            {"title": "Layer the Subtext", "detail": "What they're talking about is never what they're talking about."},
            {"title": "Interrupt the Pattern", "detail": "Introduce a disruption that forces the subtext to the surface"},
            {"title": "Land the Emotion", "detail": "End on the feeling, not the information."},
        ],
    }),

    ("text_cards", {
        "title": "Storytelling Archetypes of Stars Hollow",
        "items": [
            {"title": "The Trickster (Lorelai)", "body": "Uses humor to deflect, connect, and survive. Every joke is a bridge and a wall simultaneously."},
            {"title": "The Scholar (Rory)", "body": "Processes the world through narrative frameworks absorbed from books. Crisis comes when life stops following the plot."},
            {"title": "The Sentinel (Luke)", "body": "Guards the space where others can be themselves. His diner is the safest room in Stars Hollow."},
            {"title": "The Monarch (Emily)", "body": "Rules through ritual and expectation. Her power comes from the world she's built, and she defends it absolutely."},
        ],
    }),

    ("text_columns", {
        "title": "Three Pillars of the Gilmore Voice",
        "columns": [
            {"heading": "Velocity", "body": "Speak faster than the audience expects. Force them to lean in. The speed communicates urgency, intelligence, and confidence — even when the character feels none of those things."},
            {"heading": "Density", "body": "Pack every line with more than one job. A joke that also reveals character. A reference that also advances plot. A throwaway that pays off three episodes later."},
            {"heading": "Specificity", "body": "Never say 'movie' when you can say 'Casablanca.' Never say 'food' when you can say 'Al's Pancake World.' The specific detail does more work than the general category ever could."},
        ],
    }),

    ("text_narrative", {
        "title": "The Coffee Metaphor",
        "lede": "Lorelai Gilmore drinks coffee the way other characters drink alcohol — compulsively, socially, and as a marker of identity.",
        "body": "But here's what's interesting from a storytelling perspective: the coffee isn't just a character quirk. It's a structural device. Coffee scenes happen at Luke's, which means they happen in the show's emotional center. Every major relationship has a coffee scene as a turning point — Luke and Lorelai's slow burn is literally measured in cups served. The coffee is the show's love language, its nervous habit, and its clock all at once. When you're building a story, find your coffee. Find the small, repeatable, specific thing that your characters do together that can carry weight without ever calling attention to itself.",
    }),

    ("text_nested", {
        "title": "The Gilmore Reference System",
        "text": "Pop culture references in Gilmore Girls serve three distinct narrative functions.",
        "items": [
            {"text": "Surface Level: Humor", "children": ["The reference is the joke. Funny if you get it, still funny from context if you don't."]},
            {"text": "Character Level: Identity", "children": ["What someone references tells you who they are. Lorelai quotes movies; Rory quotes books; Emily quotes social customs."]},
            {"text": "Thematic Level: Commentary", "children": ["The show references tragic love stories when Lorelai makes romantic mistakes. The references are the show's Greek chorus."]},
        ],
    }),

    ("text_split", {
        "title": "Why the Final Four Words Matter",
        "headline": "The ending Amy Sherman-Palladino planned from day one was always going to be divisive — and that's precisely why it works as a storytelling case study.",
        "detail": "Sherman-Palladino revealed early in the show's run that she had planned the final four words: 'Mom?' 'Yeah?' 'I'm pregnant.' She held this ending through 7 seasons, a departure, and a revival.",
        "points": [
            "The cycle of Gilmore women repeats — Rory mirrors Lorelai mirrors Emily",
            "Every choice Rory made was leading here, even when it didn't seem like it",
            "The audience debates it because it's ambiguous enough to mean different things — that's a feature",
        ],
    }),

    ("text_annotated", {
        "title": "Key Scenes for Storytelling Study",
        "items": [
            {"label": "S1E1: Pilot", "text": "Watch the first Luke's Diner scene. In 90 seconds, it establishes setting, central relationship, power dynamic, humor style, and the coffee motif."},
            {"label": "S3E22: Graduation", "text": "Rory's valedictorian speech works because it's earned. 67 episodes of setup make 3 minutes of monologue land with precision."},
            {"label": "S5E13: Wedding Bells", "text": "The rehearsal dinner is the show's best pressure-cooker scene. Every character is trapped, alcohol is flowing, and subtext detonates."},
            {"label": "S6E22: Partings", "text": "ASP's final episode. Watch how a showrunner says goodbye — every scene is both a story beat and a meta-commentary on endings."},
        ],
    }),
]
