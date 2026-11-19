// Builds voice-over recording scripts (.docx) from the episode markdown scripts.
// Usage: node build-vo-scripts.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, BorderStyle, LevelFormat, Footer, PageNumber,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const GOLD = "8A5A12", GOLD_FILL = "F3E3C3", GOD_FILL = "FBF4E6", PAUSE_FILL = "EEEEEE";
const W = { num: 950, voice: 1250, line: 4960, dir: 2200 }; // sums to 9360 (US Letter, 1" margins)

const EPISODES = [
  {
    md: "scripts/practical-ways-to-live-in-my-purpose.md",
    out: "word-docs/Texting-God-Ep1-Purpose-VOICE-SCRIPT.docx",
    code: "E1",
    title: "Episode 1: How Do I Actually Live My Purpose?",
    pron: [["Nehemiah", "nee-uh-MY-uh"], ["Ephesians", "ee-FEE-zhunz"], ["Jeremiah", "jair-uh-MY-uh"], ["Exodus", "EK-suh-dus"]],
    dir: {
      "Are you awake?": "Quiet, almost a whisper. Not sure anyone's there.",
      "I never sleep.": "Calm and warm. A tiny smile.",
      "I've been waiting for you to ask the real question.": "Gentle, knowing.",
      "What's my purpose? Like actually.": "Frustrated, honest, voice cracking slightly.",
      "You're not wasting it.": "Slow. Let it land.",
      "You're waiting on it.": "Stress 'waiting'. Then a beat.",
      "If you've ever felt like everyone else got a map": "Straight to camera, warm, like a voice note to a friend.",
      "Okay, that sounds nice.": "A bit sceptical. Real.",
      "Most people do the first six.": "Lower your voice. This is the hook.",
      "That's what you": "Gentle correction, a soft smile.",
      "…my phone?": "Confused, a little sheepish.",
      "You keep asking for the mountain.": "Kind but direct.",
      "Honestly? Kids who grow up": "Vulnerable, quieter. This is personal.",
      "I know. I was there.": "Very gentle, almost a whisper. Long pause after.",
      "Your pain is not your disqualification.": "Slow. Emphasise 'direction'.",
      "So comparing myself is literally": "Lightbulb moment.",
      "Now you're getting it.": "Proud, warm.",
      "Hold on—": "Distracted, cut yourself off.",
      "Go ahead. I'll wait.": "Playful, patient.",
      "…no. Keep going.": "Choosing. A small resolve.",
      "…my phone. 😬": "Caught out, sheepish laugh.",
      "Okay I see where this is going.": "Laughing at yourself.",
      "…why does one of them have a red flag": "Laugh-cry, half joking.",
      "You tell Me.": "Playful, then let it sit.",
      "I don't give floodlights. I give lamps.": "The key line. Slow and clear.",
      "So I'll never feel ready.": "Realising, a little scared.",
      "No. You'll feel": "Warm, confident. Stress 'sent'.",
      "What if it's too late?": "Breaking a little. Slow.",
      "Look at the time.": "Soft.",
      "It's late. But it's not over.": "Firm and tender together.",
      "I'm not starting over with you. I'm continuing.": "Slowest line in the video.",
      "Okay. I think I'm ready to try.": "Lighter, hopeful.",
      "Put the phone down.": "Gentle but clear.",
      "Wait. Are you leaving?": "Sudden worry.",
      "Never.": "Immediate, certain, warm.",
      "I'm coming with you.": "Smile in the voice.",
      "I don't know who needed this tonight.": "To camera, soft and sincere.",
      "Out of the 7, which one hit hardest?": "Brighter, inviting.",
      "Subscribe if you want to keep reading": "Warm, a hint of mystery on 'too scared'.",
    },
  },
  {
    md: "scripts/why-do-i-still-feel-alone.md",
    out: "word-docs/Texting-God-Ep2-Alone-VOICE-SCRIPT.docx",
    code: "E2",
    title: "Episode 2: Why Do I Still Feel Alone?",
    pron: [["Elijah", "ih-LY-juh"], ["Galatians", "guh-LAY-shunz"], ["Tolu", "TOH-loo"], ["Isaiah", "eye-ZAY-uh"]],
    dir: {
      "Can I ask You something": "Tired, guarded, a bit sarcastic.",
      "Ask Me the real one.": "Calm, inviting.",
      "Why do I still feel alone?": "Quiet. Raw. Barely above a whisper.",
      "I have people. I go to church.": "Listing it out, frustrated.",
      "Because being surrounded and being known": "Slow, tender. Stress 'surrounded' and 'known'.",
      "And I'm going to tell you something": "Lower your voice. The hook.",
      "If you've ever sat in a full room": "To camera, warm and honest.",
      "Every Friday it's the same.": "Tired honesty, small self-mocking laugh on the quotes.",
      "I know. I'm up at night too.": "Light, warm humour.",
      "Okay but seriously.": "Awkward laugh, then sincere.",
      "No. Loneliness is not a sign": "Reassuring, firm.",
      "It's a signal. Like hunger.": "Explaining gently.",
      "But it really feels like You left too.": "Hurt, quiet.",
      "Feelings are real. But they're not always right.": "Kind, steady.",
      "You can feel alone and still not be alone.": "Slow. Let it land.",
      "When did you last sleep properly?": "Caring, slightly knowing.",
      "…define properly.": "Dry humour.",
      "That's what I thought.": "Playful.",
      "Wait. You fed him before You fixed him?": "Surprised, moved.",
      "Exactly. A tired body tells lonely stories.": "Warm, clear.",
      "\"I'm good! You?\"": "Fake-cheerful, then let it drop.",
      "…not usually.": "Quiet admission.",
      "Then here's the hard truth.": "Gentle but direct.",
      "But I always reach out first.": "Defensive.",
      "…oh.": "Realisation. Small.",
      "You're not as forgotten as you think.": "Soft, kind.",
      "Go first. Not because they deserve it.": "Stress 'you' at the end.",
      "Wait. I didn't even know": "Surprised.",
      "You muted it in March.": "Deadpan, playful.",
      "…how do You know that?": "Half laughing.",
      "I know everything. Reply to Tolu.": "Teasing, warm.",
      "Me! What time?": "Bright, a bit nervous.",
      "Okay that was kind of beautiful.": "Touched, then practical.",
      "You're right. One brunch won't.": "Playful, then sincere.",
      "Do you know who else is awake": "Gentle, opening her eyes.",
      "I didn't even think about them.": "Quiet, a little ashamed.",
      "The cure for your loneliness": "Slow.",
      "Can I ask one more thing?": "Hesitant.",
      "Always.": "Immediate, warm.",
      "Do You even know what this feels like?": "Hurt, almost accusing.",
      "The night before I died": "Slow. Almost a whisper.",
      "They fell asleep. Three times.": "Whisper. Long pause after.",
      "The next day, on the cross": "Heavy, slow. Don't shout the quote.",
      "I know loneliness better than anyone": "Steady, tender.",
      "I was forsaken so that you would never have to be.": "Slowest line in the video.",
      "I don't know what to say.": "Emotional, quiet.",
      "You don't have to say anything.": "Very gentle.",
      "Okay. I'm going to bed.": "Lighter, a smile returning.",
      "You were never as alone as you felt tonight.": "Tender.",
      "I was here the whole time.": "Warm, a little playful on 'brunch'.",
      "You're coming to brunch?": "Laughing.",
      "I go everywhere you go.": "Smile in the voice, certain.",
      "If you felt seen in this video": "To camera, sincere.",
      "And then text one person.": "Gentle push.",
      "Subscribe for the next chat": "Warm, a hint of mystery.",
    },
  },
  {
    md: "scripts/why-havent-you-answered-my-prayer.md",
    out: "word-docs/Texting-God-Ep3-Prayer-VOICE-SCRIPT.docx",
    code: "E3",
    title: "Episode 3: Why Haven't You Answered My Prayer?",
    note: "This episode is the quietest in the series. She's whispering at 4 AM beside her sleeping mum. Record close to the mic and keep everything soft. The silences matter more than the words.",
    pron: [["Lazarus", "LAZ-uh-rus"], ["Corinthians", "kor-IN-thee-unz"], ["Isaiah", "eye-ZAY-uh"], ["Tolu", "TOH-loo"]],
    dir: {
      "She's finally asleep.": "Whisper. Exhausted relief.",
      "The doctor used a new word today.": "Flat, numb. Stress 'Comfortable'.",
      "I know what that means.": "Barely audible. Long pause after.",
      "I have prayed for my mum every single night": "Quiet, steady, holding back tears.",
      "Why haven't You answered me?": "The hurt breaks through. Still a whisper.",
      "I'm here.": "Very soft, close.",
      "I'm in the room.": "Slow. Let it land.",
      "Then why aren't You doing anything?": "Angry, but whispered. Don't shout.",
      "I'm doing more than you can see tonight.": "Gentle, unhurried.",
      "But I'm not going to pretend this doesn't hurt.": "Tender. A catch on 'I know'.",
      "I almost didn't post this one.": "To camera. Quiet, honest, no performance.",
      "Everyone keeps sending me Bible verses.": "Tired, a little bitter.",
      "I don't want a verse right now.": "Firm, small.",
      "Then I won't send you one.": "Kind, immediate.",
      "I'll just stay.": "Softest line so far. Then silence.",
      "…thank You.": "Whisper, almost crying.",
      "When My friend Lazarus died": "Slow, remembering.",
      "I still stood outside his tomb": "Emotional. Let your voice soften.",
      "I don't rush your grief.": "Tender, certain.",
      "Can I ask the thing I'm scared to ask?": "Hesitant, small.",
      "Always.": "Warm, immediate.",
      "Is it my fault?": "Barely a whisper.",
      "Did I not pray enough?": "Spilling out faster, guilt.",
      "No.": "Firm and quick. The only firm moment.",
      "Listen to Me. No.": "Firm, then softening on the second 'No'.",
      "Her sickness is not a punishment.": "Steady, deliberate.",
      "When My disciples asked whose sin": "Calm.",
      "You did not cause this.": "Slow. Stress 'did not'.",
      "I've been carrying that for eight months.": "Release. Quiet tears.",
      "I know. Put it down.": "Gentle, like taking something from her hands.",
      "Then why? You healed people in the Bible.": "Pleading.",
      "Yes. I could.": "Honest. No defensiveness.",
      "Then WHY?": "Breaking. Louder, but still held back.",
      "There are answers I can't give you tonight": "Very slow, tender.",
      "Right now you see through a dark glass.": "Gentle.",
      "Paul begged Me three times": "Steady.",
      "That sounds like a no.": "Bitter, tired.",
      "It sounds like a no. It was a promise.": "Stress 'promise'.",
      "I never promised you wouldn't go through the water.": "The key line. Slow and clear.",
      "I haven't answered her in two weeks.": "Ashamed, quiet.",
      "I don't want anyone to see me like this.": "Vulnerable.",
      "She's not coming to see you strong.": "Warm.",
      "Let her in.": "Gentle push. Small smile in the voice.",
      "Come up. Room 412.": "Small, brave. She's typing as she says it.",
      "Do You even hear her prayers?": "Quiet, a little desperate.",
      "I hear every word.": "Certain, soft.",
      "Do you want to know what she prays?": "Very gentle. Pause after.",
      "…yes.": "Barely audible.",
      "She stopped praying for herself": "Slow.",
      "Every night she prays for you.": "Stress 'you'.",
      "She says: \"Whatever happens to me": "The heart of the episode. Slowest line. Take your time.",
      "She never told me.": "Broken whisper.",
      "She's prayed some version of it": "Tender.",
      "And that prayer? I'm answering it right now.": "A small lift of hope.",
      "You're awake at 4 AM in a hospital chair.": "Warm, almost proud.",
      "Will You heal her?": "Whisper. Long pause after.",
      "I can't promise you tomorrow": "Honest, gentle.",
      "But I promise you this": "Steady, certain.",
      "And one day, I will wipe every tear": "Hope rising. Still soft.",
      "No more hospitals. No more 4 AMs.": "Tender, a gentle smile.",
      "Stay with me?": "Childlike, small.",
      "I'm in the room. I never left it.": "Final line. Slow, warm, certain.",
      "If someone you love is sick tonight": "To camera. Sincere and gentle.",
      "And if you're the one carrying it alone": "Soft encouragement.",
    },
  },
];

const DEFAULT_DIR = {
  ME: "Natural, honest. Her own voice.",
  GOD: "About 20% slower, softer, lower.",
  VO: "To camera, warm. Like a voice note to a friend.",
};
const EMOJI_DIR = { "😄": "smiling", "😬": "sheepish", "😂": "laughing", "😅": "awkward laugh", "🙂": "fake-cheerful", "😭": "laugh-cry" };

// ---------- parsing ----------
function cleanSpoken(raw) {
  let refs = [];
  let t = raw.replace(/\(([1-3]?\s?[A-Z][a-z]+\s\d+:\d+(?:[–-]\d+)?(?:,\s?\d+)?(?:,\s?NKJV)?)\)/g, (_, r) => { refs.push(r); return ""; });
  const emo = [];
  for (const [e, d] of Object.entries(EMOJI_DIR)) if (t.includes(e)) emo.push(d);
  t = t.replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}]/gu, "").replace(/\s+/g, " ").trim();
  return { text: t, refs, emo };
}

function parse(md) {
  const lines = fs.readFileSync(path.join(ROOT, md), "utf8").split("\n");
  const sections = [];
  let cur = null, started = false;
  for (const ln of lines) {
    const h = ln.match(/^## (\d+:\d+ – .+)$/);
    if (h) { started = true; cur = { title: h[1].replace(/\s*\(.*?\)\s*$/, ""), rows: [] }; sections.push(cur); continue; }
    if (/^## PRODUCTION NOTES/.test(ln)) break;
    if (!started || !cur) continue;
    let m;
    if ((m = ln.match(/^\*\*(ME|GOD):\*\*\s*(.*)$/))) cur.rows.push({ voice: m[1], ...cleanSpoken(m[2]), raw: m[2] });
    else if ((m = ln.match(/^\*\(VO\)\*\s*(.*)$/))) cur.rows.push({ voice: "VO", ...cleanSpoken(m[1]), raw: m[1] });
    else if ((m = ln.match(/types: \*"(.+?)"\*/))) cur.rows.push({ voice: "ME", ...cleanSpoken(m[1]), raw: m[1], typed: true });
    else if ((m = ln.match(/old message bubble.*?: \*"(.+?)"\*/)) || (m = ln.match(/^`\[SCREEN\]` \*"(.+?)" \(\d{4}\)\*/))) cur.rows.push({ voice: "ME", ...cleanSpoken(m[1]), raw: m[1], reading: true });
    if (/⏸/.test(ln) || /holds for/.test(ln)) {
      const s = ln.match(/(\d+) (?:full )?seconds?/);
      const secs = s ? s[1] : (/holds … holds/.test(ln) ? "3" : "2");
      if (!/^\*\*(ME|GOD)/.test(ln)) {
        const prev = cur.rows[cur.rows.length - 1];
        if (prev && prev.pause) prev.pause = String(Math.max(+prev.pause, +secs));
        else cur.rows.push({ pause: secs });
      }
    }
  }
  return sections;
}

// ---------- docx helpers ----------
function runsFromMd(text, base = {}) {
  // **bold** and *italic*
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const s = m[0];
    if (s.startsWith("**")) out.push(new TextRun({ text: s.slice(2, -2), bold: true, ...base }));
    else out.push(new TextRun({ text: s.slice(1, -1), italics: true, ...base }));
    last = m.index + s.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(children, width, fill, opts = {}) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: "center", ...opts,
    children: Array.isArray(children) ? children : [children],
  });
}
const P = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], ...o });
const T = (text, o = {}) => new TextRun({ text, ...o });

function findDir(dirs, text) {
  // prefer the longest key contained in the line, so "I'm in the room. I never left it." beats "I'm in the room."
  let best = null, bestLen = 0;
  for (const [k, v] of Object.entries(dirs)) {
    const hit = text.includes(k) || (k.includes(text) && text.length > 3);
    if (hit && k.length > bestLen) { best = v; bestLen = k.length; }
  }
  return best;
}

function buildEpisode(ep) {
  const sections = parse(ep.md);
  let n = 0;
  const body = [];
  for (const sec of sections) {
    body.push(P(T(sec.title.toUpperCase()), { heading: HeadingLevel.HEADING_2 }));
    const rows = [new TableRow({
      tableHeader: true,
      children: [
        cell(P(T("#", { bold: true })), W.num, GOLD_FILL),
        cell(P(T("Voice", { bold: true })), W.voice, GOLD_FILL),
        cell(P(T("Line (read exactly this)", { bold: true })), W.line, GOLD_FILL),
        cell(P(T("Delivery", { bold: true })), W.dir, GOLD_FILL),
      ],
    })];
    for (const r of sec.rows) {
      if (r.pause) {
        rows.push(new TableRow({ children: [cell(P(T(`⏸  PAUSE ${r.pause} SEC  (silence; leave room tone)`, { italics: true, color: "666666", size: 20 }), { alignment: AlignmentType.CENTER }), W.num + W.voice + W.line + W.dir, PAUSE_FILL, { columnSpan: 4 })] }));
        continue;
      }
      let spoken = r.text;
      let dir = findDir(ep.dir, r.raw.replace(/\*\*/g, "")) || findDir(ep.dir, spoken.replace(/\*/g, "")) || DEFAULT_DIR[r.voice];
      if (!spoken.replace(/[*\s]/g, "")) {
        continue; // emoji-only bubble: nothing to record
      } else if (spoken === "?") {
        spoken = "Hm?";
        dir = "Curious.";
      } else if (r.emo.length && dir === DEFAULT_DIR[r.voice]) {
        dir += ` (${r.emo.join(", ")})`;
      }
      if (r.typed && !/typing/.test(dir)) dir = dir + " She's typing this as she says it.";
      if (r.reading && !/[Rr]ead/.test(dir)) dir = "Reading her old message aloud. " + dir;
      n++;
      const id = `${ep.code}-${String(n).padStart(3, "0")}`;
      const fill = r.voice === "GOD" ? GOD_FILL : undefined;
      const voiceLabel = { ME: "ME", GOD: "GOD", VO: "TO CAMERA" }[r.voice];
      const lineChildren = [P(runsFromMd(spoken, { size: 26, ...(r.voice === "GOD" ? { color: "3A2A08" } : {}) }), { spacing: { line: 300 } })];
      if (r.refs.length) lineChildren.push(P(T(`On screen only (don't read): ${r.refs.join("; ")}`, { size: 16, color: "8C8C8C", italics: true })));
      rows.push(new TableRow({
        cantSplit: true,
        children: [
          cell(P(T(id, { size: 18, color: "666666" })), W.num, fill),
          cell(P(T(voiceLabel, { bold: true, size: 20, color: r.voice === "GOD" ? GOLD : "222222" })), W.voice, fill),
          cell(lineChildren, W.line, fill),
          cell(P(T(dir, { italics: true, size: 18, color: "555555" })), W.dir, fill),
        ],
      }));
    }
    body.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [W.num, W.voice, W.line, W.dir], rows }));
  }

  const bullet = (runs) => new Paragraph({ numbering: { reference: "b", level: 0 }, children: Array.isArray(runs) ? runs : [runs] });
  const front = [
    P(T(`VOICE RECORDING SCRIPT`), { heading: HeadingLevel.TITLE }),
    P(T(`Texting God · ${ep.title}`, { size: 28, color: GOLD, bold: true }), { alignment: AlignmentType.CENTER, spacing: { after: 240 } }),
    P(T(`${n} lines · read by one narrator · about 9 minutes when edited with pauses`, { color: "666666" }), { alignment: AlignmentType.CENTER, spacing: { after: 360 } }),

    ...(ep.note ? [P([T("Before you record: ", { bold: true, color: GOLD }), T(ep.note, { italics: true })], { spacing: { after: 240 } })] : []),
    P(T("Your three voices"), { heading: HeadingLevel.HEADING_1 }),
    bullet([T("ME ", { bold: true }), T("is her own messages. Natural speed, honest, a little tired. Laugh where it's funny.")]),
    bullet([T("GOD ", { bold: true, color: GOLD }), T("is God's replies, read aloud as they arrive. About 20% slower, softer, slightly lower. Never preachy: read them as if the message just stopped you in your tracks. These rows are shaded gold.")]),
    bullet([T("TO CAMERA ", { bold: true }), T("is you talking straight to the viewer, warm, like a voice note to a friend.")]),

    P(T("How to record"), { heading: HeadingLevel.HEADING_1 }),
    bullet([T("Record every line ", {}), T("2–3 times", { bold: true }), T(". Say the line number before each one (\"E1-014, take 2\") so the editor can find it.")]),
    bullet([T("Leave about ", {}), T("1 second of silence", { bold: true }), T(" before and after every take. Don't record the PAUSE rows: those are for the editor. They show where silence goes on screen.")]),
    bullet([T("Scripture references in grey ("), T("On screen only", { italics: true }), T(") appear on screen. Don't read them aloud.")]),
    bullet([T("Bold words", { bold: true }), T(" in a line are the words to stress.")]),
    bullet(T("Pace: about 120–130 words a minute for ME and TO CAMERA, slower for GOD. If it feels too slow, it's probably right.")),
    bullet(T("Record 10 seconds of silent room tone at the start. The editor uses it to fill gaps.")),
    bullet(T("Quiet room, soft furnishings (a closet full of clothes works well), phone on airplane mode, water nearby, no ice.")),
    bullet(T("Smile on the light lines: listeners can hear it.")),
    bullet([T("Save files as "), T(`${ep.code}_L014_take2.wav`, { font: "Consolas" }), T(", or record in one long file and slate the numbers out loud.")]),

    P(T("Pronunciation"), { heading: HeadingLevel.HEADING_1 }),
    ...ep.pron.map(([w, p]) => bullet([T(w, { bold: true }), T(`: ${p}`)])),
    new Paragraph({ children: [], pageBreakBefore: true }),
  ];

  return new Document({
    creator: "Texting God",
    title: `Voice script – ${ep.title}`,
    styles: {
      default: { document: { run: { font: "Calibri", size: 22 } } },
      paragraphStyles: [
        { id: "Title", name: "Title", basedOn: "Normal", run: { font: "Georgia", size: 44, bold: true, color: GOLD }, paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 600, after: 120 }, keepNext: true } },
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: "Georgia", size: 30, bold: true, color: GOLD }, paragraph: { spacing: { before: 300, after: 120 }, keepNext: true, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: "Georgia", size: 26, bold: true, color: GOLD }, paragraph: { spacing: { before: 360, after: 120 }, keepNext: true, outlineLevel: 1 } },
      ],
    },
    numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
      footers: { default: new Footer({ children: [P([T(`Texting God · ${ep.code} voice script · page `, { size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" })], { alignment: AlignmentType.CENTER })] }) },
      children: [...front, ...body],
    }],
  });
}

(async () => {
  for (const ep of EPISODES) {
    const buf = await Packer.toBuffer(buildEpisode(ep));
    fs.writeFileSync(path.join(ROOT, ep.out), buf);
    console.log("wrote", ep.out);
  }
})();
