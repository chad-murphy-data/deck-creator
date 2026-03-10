"""Generate a comprehensive demo deck showcasing all slide types."""
import template_bold
import os

slides = [
    # 1. TITLE
    ("title", {
        "title": "Everything I Learned About Storytelling I Learned from Video Games",
        "subtitle": "Lessons from Jesse Schell and the Art of Interactive Narrative",
        "author": "Chad Murphy",
        "date": "March 2026"
    }),

    # 2. AGENDA
    ("agenda", {
        "title": "What We'll Cover",
        "items": [
            {"title": "The Lens of the Player", "detail": "Why games teach us more about story than film ever could"},
            {"title": "The Elemental Tetrad", "detail": "Mechanics, Story, Aesthetics, Technology"},
            {"title": "The Experience Engine", "detail": "Designing for emotion, not just engagement"},
            {"title": "Applied Lenses", "detail": "Practical frameworks for any storytelling medium"},
            {"title": "What Games Get Wrong", "detail": "Where interactive narrative still falls short"},
        ]
    }),

    # 3. IN_BRIEF
    ("in_brief", {
        "title": "The Core Insight",
        "bullets": [
            "Games create conditions for stories to emerge \u2014 they don\u2019t just tell them",
            "Jesse Schell\u2019s Art of Game Design identifies 100+ lenses for examining experience",
            "The best game narratives make the player the author of their own meaning",
            "Every medium can learn from how games handle agency and consequence",
        ]
    }),

    # 4. STAT_CALLOUT
    ("stat_callout", {
        "title": "The Scale of the Medium",
        "stat": "3.2B",
        "headline": "Active gamers worldwide in 2025",
        "detail": "More people play games than watch Netflix, go to movies, and read novels combined. This is the largest storytelling audience in human history.",
        "source": "Newzoo Global Games Market Report, 2025",
        "sectionColor": "cobalt"
    }),

    # 5. QUOTE
    ("quote", {
        "title": "On Experience Design",
        "quote": "The game is not the experience. The game enables the experience.",
        "attribution": "Jesse Schell",
        "context": "The Art of Game Design: A Book of Lenses",
        "sectionColor": "green"
    }),

    # 6. COMPARISON
    ("comparison", {
        "title": "Linear vs. Interactive Storytelling",
        "leftLabel": "Film / Novel",
        "rightLabel": "Games",
        "leftItems": ["Author controls pacing", "Fixed emotional arc", "Passive audience", "One canonical story", "Meaning is delivered"],
        "rightItems": ["Player controls pacing", "Emergent emotional arc", "Active participant", "Branching possibilities", "Meaning is discovered"],
        "sectionColor": "purple"
    }),

    # 7. TEXT_GRAPH
    ("text_graph", {
        "title": "Player Engagement Over Time",
        "texts": [
            "Schell argues that engagement follows an interest curve \u2014 a wave of tension and release.",
            "Great games front-load a hook, build through rising action, and deliver peaks at calculated intervals.",
            "The pattern mirrors classical dramatic structure, but player agency makes each peak feel earned."
        ],
        "chartType": "line",
        "chartTitle": "Engagement Curve",
        "series": [{"label": "Engagement", "values": [65, 72, 58, 85, 70, 95, 80, 98]}],
        "labels": ["Hook", "Tutorial", "Valley", "Boss", "Explore", "Midpoint", "Lull", "Climax"],
        "sectionColor": "blue"
    }),

    # 8. PROCESS_FLOW
    ("process_flow", {
        "title": "Schell\u2019s Experience Design Loop",
        "steps": [
            {"title": "Listen", "detail": "Observe the player. What are they actually feeling?"},
            {"title": "Ideate", "detail": "Apply lenses to generate design possibilities"},
            {"title": "Prototype", "detail": "Build the smallest testable version"},
            {"title": "Playtest", "detail": "Watch real players. Don\u2019t explain anything."},
            {"title": "Iterate", "detail": "Kill your darlings. Serve the experience."},
        ],
        "sectionColor": "green"
    }),

    # 9. PROCESS_FLOW_REVEAL
    ("process_flow_reveal", {
        "title": "The Hero\u2019s Journey in Game Design",
        "steps": [
            {"title": "The Ordinary World", "detail": "Tutorial area. Safe, familiar. The player learns the rules before the stakes arrive."},
            {"title": "The Call to Adventure", "detail": "The inciting incident. Something disrupts the status quo. The first real quest."},
            {"title": "The Transformation", "detail": "Through trials, the player grows. New abilities mirror character growth. The mechanics become the narrative arc."},
        ],
        "sectionColor": "cobalt"
    }),

    # 10. MATRIX
    ("matrix", {
        "title": "The Elemental Tetrad",
        "quadrants": [
            {"label": "Mechanics", "items": ["Rules and systems", "Player choices", "Win/loss conditions", "Feedback loops"]},
            {"label": "Story", "items": ["Narrative arc", "Character development", "World-building", "Dramatic tension"]},
            {"label": "Aesthetics", "items": ["Visual design", "Sound and music", "Emotional tone", "Sensory experience"]},
            {"label": "Technology", "items": ["Platform capabilities", "Rendering engine", "Input methods", "Network infrastructure"]}
        ],
        "sectionColor": "purple"
    }),

    # 11. METHODS
    ("methods", {
        "title": "Research Approach",
        "fields": [
            {"title": "Played 50+ narrative-driven games across genres", "detail": "RPGs, walking simulators, visual novels, survival horror, sandbox builders"},
            {"title": "Coded every game against Schell\u2019s 113 lenses", "detail": "Mapped which lenses each game leaned on most heavily"},
            {"title": "Interviewed 12 game narrative designers", "detail": "Including writers from Naughty Dog, Bethesda, and indie studios"},
            {"title": "Cross-referenced with film/novel narrative frameworks", "detail": "McKee\u2019s Story, Campbell\u2019s Hero\u2019s Journey, Aristotle\u2019s Poetics"},
        ],
        "sectionColor": "gold"
    }),

    # 12. HYPOTHESES
    ("hypotheses", {
        "title": "Working Hypotheses",
        "hypotheses": [
            {"text": "Player agency deepens emotional impact more than passive consumption", "status": "Confirmed"},
            {"text": "Environmental storytelling transfers to non-game contexts (UX, education)", "status": "Confirmed"},
            {"text": "Branching narrative always produces better stories than linear", "status": "Rejected"},
            {"text": "Schell\u2019s lenses apply to any experience design, not just games", "status": "Confirmed"},
        ],
        "sectionColor": "blue"
    }),

    # 13. WSN_DENSE
    ("wsn_dense", {
        "title": "The Lens of the Player",
        "labels": ["What We Found", "So What", "Now What"],
        "what": {
            "headline": "Games that let players construct their own meaning consistently outperform guided narratives in emotional recall.",
            "detail": "Players remembered emergent stories 3x longer than scripted moments."
        },
        "soWhat": {
            "headline": "Traditional storytelling overvalues authorial control at the expense of audience investment.",
            "detail": "The instinct to control the narrative is often the enemy of resonance."
        },
        "nowWhat": {
            "headline": "Design for possibility spaces, not plot points.",
            "detail": "Build systems that generate stories rather than scripts that deliver them."
        },
        "sectionColor": "green"
    }),

    # 14. WSN_REVEAL
    ("wsn_reveal", {
        "title": "Environmental Storytelling",
        "labels": ["Discovery", "Implication", "Application"],
        "what": {
            "headline": "The richest game narratives are told through spaces, not dialogue.",
            "detail": "Dark Souls, Gone Home, and Bioshock tell their deepest stories through environmental details.",
            "summary": "Space tells story better than script"
        },
        "soWhat": {
            "headline": "This maps directly to how humans process physical environments.",
            "detail": "Retail stores, museums, and workspaces tell stories through spatial design.",
            "summary": "Spatial narrative is universal"
        },
        "nowWhat": {
            "headline": "Apply environmental storytelling to any designed experience.",
            "detail": "Presentations, dashboards, and org charts can use spatial techniques from level design.",
            "summary": "Design spaces that narrate"
        },
        "sectionColor": "cobalt"
    }),

    # 15. FINDINGS_RECS
    ("findings_recs", {
        "title": "Key Findings from Lens Analysis",
        "items": [
            {"finding": "The Lens of Surprise is the most commonly underused tool in corporate storytelling", "recommendation": "Build at least one genuine surprise into every presentation"},
            {"finding": "Games use the Lens of Flow to maintain engagement across hours of play", "recommendation": "Map your content to the flow channel: challenge must scale with understanding"},
            {"finding": "The Lens of Story Machine shows the best narratives emerge from systems", "recommendation": "Give audiences frameworks for thinking, not just conclusions to memorize"},
        ],
        "sectionColor": "purple"
    }),

    # 16. FINDINGS_RECS_DENSE
    ("findings_recs_dense", {
        "title": "Detailed Lens Audit",
        "items": [
            {"finding": "Lens of Curiosity: Games front-load mystery", "recommendation": "Open with a question your audience cannot ignore"},
            {"finding": "Lens of Endogenous Value: Rewards feel earned", "recommendation": "Create contexts where insights feel discovered"},
            {"finding": "Lens of the Toy: Fun before rules kick in", "recommendation": "Your core idea should be interesting independent of evidence"},
            {"finding": "Lens of Feedback: Close loops instantly", "recommendation": "Build moments where the audience tests understanding"},
        ],
        "sectionColor": "gold"
    }),

    # 17. OPEN_QUESTIONS
    ("open_questions", {
        "title": "Open Questions",
        "items": [
            "Can AI-generated narrative achieve the emergent quality of player-driven stories?",
            "Does interactive storytelling work for persuasion, or only for exploration?",
            "What is the right level of agency for a business presentation?",
            "How do we measure the resonance of an experience vs. mere engagement?",
        ],
        "sectionColor": "blue"
    }),

    # 18. PROGRESSIVE_REVEAL
    ("progressive_reveal", {
        "title": "Building the Case for Game-Informed Design",
        "takeaways": [
            {"headline": "Games are the most sophisticated storytelling technology ever created", "detail": "They handle branching narrative, real-time feedback, and emergent behavior simultaneously.", "summary": "Games are the apex storytelling medium"},
            {"headline": "Schell\u2019s lenses provide a universal vocabulary for experience design", "detail": "The 113 lenses work for products, presentations, classrooms, and organizations.", "summary": "Schell\u2019s lenses are universal"},
            {"headline": "The future of storytelling is participatory", "detail": "Every audience increasingly expects agency. The winners will design for co-creation.", "summary": "Storytelling is becoming participatory"},
        ],
        "sectionColor": "green"
    }),

    # 19. TIMELINE
    ("timeline", {
        "title": "Evolution of Game Narrative",
        "milestones": [
            {"date": "1980", "title": "Zork", "detail": "Text adventures prove narrative can be interactive", "status": "complete"},
            {"date": "1998", "title": "Half-Life", "detail": "First-person storytelling without cutscenes", "status": "complete"},
            {"date": "2007", "title": "BioShock", "detail": "Environmental storytelling + meta-commentary on agency", "status": "complete"},
            {"date": "2013", "title": "The Last of Us", "detail": "Cinematic narrative with mechanical resonance", "status": "complete"},
            {"date": "2023", "title": "Baldur\u2019s Gate 3", "detail": "Tabletop-scale branching in a AAA game", "status": "current"},
        ],
        "sectionColor": "cobalt"
    }),

    # 20. DATA_TABLE
    ("data_table", {
        "title": "Schell\u2019s Most-Applied Lenses by Domain",
        "headers": ["Lens", "Games", "Film", "Product", "Education"],
        "rows": [
            ["Surprise", "\u2605\u2605\u2605", "\u2605\u2605", "\u2605", "\u2605\u2605"],
            ["Curiosity", "\u2605\u2605\u2605", "\u2605\u2605", "\u2605\u2605\u2605", "\u2605\u2605\u2605"],
            ["Flow", "\u2605\u2605\u2605", "\u2605", "\u2605\u2605", "\u2605\u2605"],
            ["Story Machine", "\u2605\u2605\u2605", "\u2605", "\u2605\u2605\u2605", "\u2605"],
            ["Feedback", "\u2605\u2605\u2605", "\u2605", "\u2605\u2605\u2605", "\u2605\u2605"],
        ],
        "highlightCol": "1",
        "note": "Ratings based on frequency of application in reviewed works",
        "sectionColor": "purple"
    }),

    # 21. MULTI_STAT
    ("multi_stat", {
        "title": "By the Numbers",
        "stats": [
            {"value": "113", "label": "Lenses in Schell\u2019s framework", "detail": "Each a different way to examine experience"},
            {"value": "50+", "label": "Games analyzed", "detail": "Across 8 genres over 4 decades"},
            {"value": "3x", "label": "Longer emotional recall", "detail": "For emergent vs. scripted narratives"},
        ],
        "source": "Internal analysis, 2025-2026",
        "sectionColor": "gold"
    }),

    # 22. PERSONA
    ("persona", {
        "title": "The Narrative Gamer",
        "name": "Alex Reyes",
        "archetype": "The Explorer",
        "traits": ["Reads every item description", "Ignores main quest for side content", "Screenshots environmental details", "Replays for alternate endings", "Argues about lore on Reddit"],
        "strategy": "Design content with hidden depth. Reward curiosity. Layer information so surface-level scanning still works, but deep engagement reveals richer meaning.",
        "detail": "Represents ~35% of narrative game players.",
        "sectionColor": "blue"
    }),

    # 23. RISK_TRADEOFF
    ("risk_tradeoff", {
        "title": "Applying Game Design to Business",
        "risks": [
            {"label": "Audience expects passivity", "detail": "Most business audiences are trained to sit and absorb", "severity": "high"},
            {"label": "Interactivity adds complexity", "detail": "Branching content requires more preparation", "severity": "medium"},
            {"label": "Gamification backlash", "detail": "Shallow game mechanics feel patronizing", "severity": "medium"},
        ],
        "rewards": [
            {"label": "Deeper engagement", "detail": "Active participation drives 40% higher retention"},
            {"label": "Memorable differentiation", "detail": "Interactive presentations stand out"},
            {"label": "Genuine understanding", "detail": "People who discover insights own them"},
        ],
        "sectionColor": "cobalt"
    }),

    # 24. APPENDIX
    ("appendix", {
        "title": "Appendix: Key References",
        "sections": [
            {"label": "Primary Source", "content": "Schell, Jesse. The Art of Game Design: A Book of Lenses. 4th Edition. CRC Press, 2024."},
            {"label": "Supporting Works", "content": "McKee, Story. Campbell, The Hero with a Thousand Faces. Csikszentmihalyi, Flow. Huizinga, Homo Ludens."},
            {"label": "Games Referenced", "content": "Zork, Half-Life, BioShock, Portal, The Last of Us, Dark Souls, Gone Home, Minecraft, Dwarf Fortress, Baldur\u2019s Gate 3, Disco Elysium, Outer Wilds."},
        ],
        "sectionColor": "green"
    }),

    # 25. BEFORE_AFTER
    ("before_after", {
        "title": "Transforming a Standard Presentation",
        "before": {"label": "Traditional Deck", "detail": "Linear slides. Speaker controls pacing. Audience passively absorbs bullets. Takeaway delivered at end. Forgotten by Friday."},
        "intervention": "Apply Schell\u2019s Lens of the Player: redesign every slide to answer \u2018what is the audience doing and feeling right now?\u2019",
        "after": {"label": "Game-Informed Deck", "detail": "Opens with a mystery. Audience makes predictions. Evidence builds through progressive reveal. Takeaway feels earned. Remembered for months."},
        "sectionColor": "purple"
    }),

    # 26. SUMMARY
    ("summary", {
        "title": "Three Takeaways",
        "sections": [
            {"heading": "Think in Lenses", "points": ["Every design decision can be examined from 113+ angles", "No single lens is sufficient", "Rotate perspectives to find blind spots"]},
            {"heading": "Design for Agency", "points": ["Let your audience discover, don\u2019t just deliver", "Create possibility spaces", "Reward curiosity and exploration"]},
            {"heading": "Serve the Experience", "points": ["The presentation is not the experience", "Mechanics, story, aesthetics, and technology must align", "Playtest with real humans"]},
        ],
        "sectionColor": "gold"
    }),

    # 27. QUOTE_FULL
    ("quote_full", {
        "quote": "A game designer is an advocate for the player. The player cannot speak for themselves, so the designer must.",
        "attribution": "Jesse Schell",
        "context": "Keynote at GDC 2023",
        "sectionColor": "cobalt"
    }),

    # 28. STAT_HERO
    ("stat_hero", {
        "title": "The Engagement Gap",
        "hero": {"value": "68%", "label": "of presentation content is forgotten within 24 hours"},
        "supporting": [
            {"value": "4%", "label": "forgotten when audience discovers insight themselves"},
            {"value": "17x", "label": "more recall for experiential vs. passive learning"},
        ],
        "source": "Adapted from National Training Laboratories research",
        "sectionColor": "blue"
    }),

    # 29. IN_BRIEF_FEATURED
    ("in_brief_featured", {
        "title": "The Core Principle",
        "featured": "Every great game \u2014 and every great presentation \u2014 is ultimately about one thing: creating the conditions for meaningful experience.",
        "supporting": [
            "Mechanics create the possibility space",
            "Story provides the emotional framework",
            "Aesthetics set the tone and mood",
            "Technology determines what\u2019s possible"
        ],
        "sectionColor": "green"
    }),

    # 30. PERSONA_DUO
    ("persona_duo", {
        "title": "Two Types of Audience",
        "personas": [
            {"name": "The Passenger", "archetype": "Passive Consumer", "traits": ["Wants clear structure", "Prefers being told the answer", "Values efficiency over exploration"], "strategy": "Provide a strong through-line and clear conclusions."},
            {"name": "The Driver", "archetype": "Active Explorer", "traits": ["Wants to figure things out", "Engages with ambiguity", "Values discovery over delivery"], "strategy": "Build in puzzles, questions, and space to construct meaning."},
        ],
        "sectionColor": "purple"
    }),

    # 31. PROCESS_FLOW_VERTICAL
    ("process_flow_vertical", {
        "title": "Applying a Lens to Your Next Deck",
        "steps": [
            {"title": "Choose Your Lens", "detail": "Pick one of Schell\u2019s 113 lenses. Start with Curiosity, Surprise, or Flow."},
            {"title": "Audit Your Content", "detail": "Walk through every slide and ask: does this serve the experience?"},
            {"title": "Playtest It", "detail": "Present to a small group. Watch their faces. Iterate ruthlessly."},
        ],
        "sectionColor": "gold"
    }),

    # 32. TEXT_CARDS
    ("text_cards", {
        "title": "Schell\u2019s Most Transferable Lenses",
        "items": [
            {"title": "Lens of Curiosity", "detail": "What questions does your experience plant in the audience\u2019s mind?"},
            {"title": "Lens of Surprise", "detail": "Where do you subvert expectations to create memorable moments?"},
            {"title": "Lens of Flow", "detail": "Is the challenge level matched to the audience\u2019s growing skill?"},
            {"title": "Lens of Story Machine", "detail": "Does your system generate stories, or just deliver one?"},
            {"title": "Lens of Freedom", "detail": "Where can the audience make meaningful choices?"},
            {"title": "Lens of Feedback", "detail": "How quickly does the audience know if they\u2019re on the right track?"},
        ],
        "sectionColor": "blue"
    }),

    # 33. TEXT_COLUMNS
    ("text_columns", {
        "title": "Game Narrative vs. Corporate Narrative",
        "columns": [
            {"heading": "What Games Do Well", "body": "Environmental storytelling. Emergent narrative from systems. Respecting player intelligence. Making the audience the protagonist."},
            {"heading": "What Corporate Does Well", "body": "Clarity of message. Data-driven persuasion. Efficient information transfer. Establishing credibility and authority."},
            {"heading": "The Synthesis", "body": "Clear messaging delivered through experiential design. Data that surprises. An audience that feels like a participant, not a spectator."},
        ],
        "sectionColor": "cobalt"
    }),

    # 34. TEXT_NARRATIVE
    ("text_narrative", {
        "title": "The Lesson of Portal",
        "lede": "Portal is a masterclass in invisible instruction. The player learns every mechanic through play, never through a manual. There are no tutorial screens, no pop-ups, no handholding.",
        "body": "Valve\u2019s designers built \u2018teaching chambers\u2019 \u2014 environments where the only way forward requires understanding a new concept. The player feels brilliant because they figured it out themselves, even though the designer guaranteed they would. This is the gold standard for experience design: constrained freedom that produces the illusion of genius. Every presentation should aspire to this feeling.",
        "sectionColor": "purple"
    }),

    # 35. TEXT_NESTED
    ("text_nested", {
        "title": "Schell\u2019s Design Pillars",
        "items": [
            {"text": "Experience", "children": ["The experience is the game", "Everything else serves it", "If the experience fails, nothing else matters"]},
            {"text": "Player", "children": ["Know your audience deeply", "Model their mental state", "Design for actual behavior, not stated preferences"]},
            {"text": "Process", "children": ["Iterate relentlessly", "Playtest with real humans", "Be willing to destroy what you love"]},
        ],
        "sectionColor": "gold"
    }),

    # 36. TEXT_SPLIT
    ("text_split", {
        "title": "Why This Matters Now",
        "headline": "We are in the middle of a storytelling revolution driven by interactivity.",
        "detail": "Every medium is converging on participation. Social media turned audiences into creators. Games turned stories into experiences. AI is turning content into conversation.",
        "points": [
            "Netflix experiments with interactive episodes",
            "Museums adopt game-like exhibit design",
            "Education shifts to experiential learning",
            "Corporate training adopts simulation",
            "Presentations become workshops",
        ],
        "sectionColor": "blue"
    }),

    # 37. TEXT_ANNOTATED
    ("text_annotated", {
        "title": "Games That Changed Storytelling",
        "items": [
            {"label": "BioShock (2007)", "text": "Deconstructed player agency itself. The twist reveals that the player has been following orders all along \u2014 a meta-commentary on the illusion of choice."},
            {"label": "Disco Elysium (2019)", "text": "Proved that a game with no combat can be riveting if the writing is extraordinary. Every stat check is a dialogue with yourself."},
            {"label": "Outer Wilds (2019)", "text": "A time-loop mystery where knowledge \u2014 not items or levels \u2014 is the only progression. Pure curiosity as game mechanic."},
        ],
        "sectionColor": "cobalt"
    }),

    # 38. SECTION_DIVIDER
    ("section_divider", {
        "title": "Final Thoughts",
        "subtitle": "What games teach us about being human",
        "sectionColor": "green"
    }),

    # 39. CLOSER
    ("closer", {
        "title": "Go Play Something",
        "subtitle": "The best way to learn about storytelling is to experience great stories \u2014 and games are the richest experience design laboratory ever built.",
        "contact": "Inspired by Jesse Schell, The Art of Game Design",
        "sectionColor": "green"
    }),
]

out = os.path.join("output", "storytelling_video_games_bold.pptx")
template_bold.build_deck(slides, out)
print(f"Generated: {os.path.getsize(out):,} bytes, {len(slides)} slides")
print(f"Output: {os.path.abspath(out)}")
