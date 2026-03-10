"""Build a demo deck: Everything I Learned About Storytelling I Learned from Video Games"""
import json, requests, sys

slides = [
    ["title", {
        "title": "Everything I Learned About Storytelling I Learned from Video Games",
        "subtitle": "Applying Jesse Schell's Lenses to Narrative Design",
        "author": "A Fellow Gamer",
        "date": "2026",
        "sectionColor": "green"
    }],

    ["section_divider", {"title": "Setting the Stage", "sectionNumber": "01", "sectionColor": "green"}],

    ["agenda", {"title": "Today's Quest Log", "items": [
        {"title": "The Lens of Essential Experience", "detail": "What makes a story feel real?"},
        {"title": "The Lens of Surprise", "detail": "Plot twists and player agency"},
        {"title": "The Lens of Story", "detail": "Narrative arcs in interactive media"},
        {"title": "The Lens of the Hero's Journey", "detail": "From tutorial to final boss"},
        {"title": "The Lens of Interest Curves", "detail": "Pacing that keeps players hooked"},
    ], "sectionColor": "green"}],

    ["in_brief", {"title": "Why Video Games?", "bullets": [
        "Games are the only medium where the audience IS the protagonist",
        "Player agency forces stories to earn their emotional beats",
        "Failure states create stakes that passive media can't match",
        "Schell's 100+ lenses provide a universal storytelling toolkit",
    ], "sectionColor": "green"}],

    ["section_divider", {"title": "Core Lenses", "sectionNumber": "02", "sectionColor": "blue"}],

    ["stat_callout", {
        "title": "The Lens of Essential Experience",
        "stat": "113",
        "headline": "Lenses in The Art of Game Design",
        "detail": "Jesse Schell's book catalogs over 113 lenses \u2014 each a different way to examine your creative work. Many apply directly to any form of storytelling.",
        "source": "The Art of Game Design, 3rd Edition",
        "sectionColor": "blue"
    }],

    ["quote", {
        "title": "The Lens of Story",
        "quote": "A game is not a story. A game is an experience. But stories can be a powerful part of that experience.",
        "attribution": "Jesse Schell",
        "context": "The Art of Game Design",
        "sectionColor": "blue"
    }],

    ["comparison", {
        "title": "Linear vs. Interactive Storytelling",
        "leftLabel": "Film / Novel",
        "leftItems": ["Fixed narrative arc", "Author controls pacing", "Audience is passive observer", "Single canonical ending"],
        "rightLabel": "Video Games",
        "rightItems": ["Branching possibilities", "Player controls pacing", "Audience is active participant", "Multiple outcomes possible"],
        "sectionColor": "blue"
    }],

    ["text_graph", {
        "title": "The Interest Curve",
        "text": [
            "Schell's Lens of Interest Curves shows that great stories follow a pattern of rising tension, peaks of excitement, and valleys of rest.",
            "Games like The Last of Us master this by alternating combat sequences with quiet exploration and dialogue."
        ],
        "sectionColor": "blue"
    }],

    ["section_divider", {"title": "Frameworks", "sectionNumber": "03", "sectionColor": "purple"}],

    ["process_flow", {"title": "The Hero's Journey in Games", "steps": [
        {"title": "Tutorial Zone", "detail": "The ordinary world \u2014 learn the controls, meet allies"},
        {"title": "The Call", "detail": "Inciting incident \u2014 the village burns, the princess vanishes"},
        {"title": "Trials & Growth", "detail": "Dungeons, bosses, skill trees \u2014 transformation through challenge"},
        {"title": "The Final Boss", "detail": "Supreme ordeal \u2014 everything learned is tested"},
        {"title": "Return Home", "detail": "The ending cutscene \u2014 changed by the journey"},
    ], "sectionColor": "purple"}],

    ["matrix", {
        "title": "Story Engagement Matrix",
        "xAxis": "Player Agency",
        "yAxis": "Narrative Depth",
        "quadrants": [
            {"label": "Walking Simulator", "detail": "Deep story, limited choices. E.g. Gone Home, Firewatch"},
            {"label": "Branching Epic", "detail": "Deep story, high agency. E.g. Mass Effect, Disco Elysium"},
            {"label": "Arcade Classic", "detail": "Minimal story, minimal agency. E.g. Tetris, Pac-Man"},
            {"label": "Sandbox", "detail": "Minimal story, high agency. E.g. Minecraft, GTA"},
        ],
        "sectionColor": "purple"
    }],

    ["methods", {"title": "Schell's Storytelling Toolkit", "fields": [
        {"label": "Lens #1", "value": "Essential Experience \u2014 What experience do I want the player to have?"},
        {"label": "Lens #2", "value": "Surprise \u2014 What surprises does my story deliver?"},
        {"label": "Lens #9", "value": "Unification \u2014 Does every element serve the theme?"},
        {"label": "Lens #72", "value": "Interest Curves \u2014 Does my pacing have peaks and valleys?"},
        {"label": "Lens #78", "value": "Story \u2014 Does my game have a story worth telling?"},
    ], "sectionColor": "purple"}],

    ["hypotheses", {"title": "Storytelling Hypotheses", "hypotheses": [
        {"text": "Player agency enhances emotional investment in story outcomes", "status": "Confirmed"},
        {"text": "Branching narratives are always better than linear ones", "status": "Rejected"},
        {"text": "Failure states make narrative stakes feel more real", "status": "Confirmed"},
    ], "sectionColor": "purple"}],

    ["section_divider", {"title": "Insights", "sectionNumber": "04", "sectionColor": "cobalt"}],

    ["wsn_dense", {
        "title": "The Lens of Emotion",
        "what": {"headline": "Games create unique emotions", "detail": "Triumph after defeating a hard boss, guilt from an NPC death, wonder at exploring a new world \u2014 these feelings are uniquely interactive."},
        "soWhat": {"headline": "Emotion drives memory", "detail": "Players remember how a game made them FEEL, not the polygon count. Storytellers in any medium should prioritize emotional truth."},
        "nowWhat": {"headline": "Design for feelings first", "detail": "Start with the emotion you want your audience to feel, then build mechanics and narrative to deliver it."},
        "sectionColor": "cobalt"
    }],

    ["wsn_reveal", {
        "title": "The Lens of Curiosity",
        "what": {"headline": "Curiosity is a renewable resource", "detail": "Games like Breath of the Wild place something interesting on every horizon. Each answer spawns two new questions."},
        "soWhat": {"headline": "Stories must earn attention", "detail": "In an age of infinite content, curiosity is the only reliable hook. Schell calls this the engine of engagement."},
        "nowWhat": {"headline": "Plant questions, delay answers", "detail": "Every scene should raise a question the audience needs answered. Mystery sustains momentum."},
        "sectionColor": "cobalt"
    }],

    ["findings_recs", {"title": "What Games Teach Storytellers", "items": [
        {"finding": "Players skip cutscenes when story feels disconnected from gameplay", "recommendation": "Integrate narrative into mechanics \u2014 show, don't tell"},
        {"finding": "Open worlds with no direction cause player fatigue", "recommendation": "Use breadcrumbs and interest curves to guide without constraining"},
        {"finding": "Moral choices without consequences feel hollow", "recommendation": "Every decision should ripple \u2014 audiences need to see impact"},
    ], "sectionColor": "cobalt"}],

    ["findings_recs_dense", {"title": "Lens-by-Lens Takeaways", "items": [
        {"finding": "Essential Experience lens: many stories lack a clear core feeling", "recommendation": "Define your one-sentence emotional promise before writing"},
        {"finding": "Surprise lens: predictable plots lose audiences", "recommendation": "Subvert one expectation per act"},
        {"finding": "Story lens: not every game needs a story", "recommendation": "Only add narrative if it serves the experience"},
        {"finding": "Interest Curve lens: pacing problems kill engagement", "recommendation": "Map your beats to a tension graph"},
        {"finding": "Emotion lens: technical polish can't replace feeling", "recommendation": "Playtest for emotion, not just bugs"},
        {"finding": "Curiosity lens: front-loading exposition kills wonder", "recommendation": "Withhold information strategically"},
    ], "sectionColor": "cobalt"}],

    ["section_divider", {"title": "Wrap-Up", "sectionNumber": "05", "sectionColor": "gold"}],

    ["open_questions", {"title": "Open Questions", "questions": [
        "Can AI-generated narratives ever match hand-crafted storytelling?",
        "Will VR create entirely new storytelling lenses?",
        "How do we measure 'emotional impact' of interactive stories?",
        "What lens is missing from Schell's framework?",
    ], "sectionColor": "gold"}],

    ["progressive_reveal", {"title": "Key Takeaways", "takeaways": [
        {"headline": "Stories are experiences", "detail": "Schell's most fundamental insight: design the experience first, then find the story that delivers it.", "summary": "Design the experience first"},
        {"headline": "Agency amplifies emotion", "detail": "When the audience participates in the story, every beat hits harder. Give people choices that matter.", "summary": "Give choices that matter"},
        {"headline": "Curiosity is the engine", "detail": "The best stories \u2014 in games or anywhere \u2014 are driven by questions the audience needs answered.", "summary": "Lead with questions"},
    ], "sectionColor": "gold"}],

    ["closer", {
        "title": "Now Go Tell Better Stories",
        "subtitle": "Apply the lenses. Play more games. Never stop iterating.",
        "contact": "@agamedesigner",
        "sectionColor": "gold"
    }],
]

payload = {
    "designSystem": "slick",
    "deckTitle": "Storytelling from Video Games",
    "slides": slides,
}

r = requests.post("http://127.0.0.1:5001/build", json=payload)
print(f"Status: {r.status_code}")
print(f"Content-Type: {r.headers.get('Content-Type')}")
if r.status_code == 200:
    outpath = "slide-creator/output/storytelling_video_games.pptx"
    with open(outpath, "wb") as f:
        f.write(r.content)
    print(f"Saved to {outpath} ({len(r.content):,} bytes)")
else:
    print(r.text[:2000])
