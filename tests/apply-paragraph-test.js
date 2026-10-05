// Test for the paragraph replacement logic of the CorelDRAW import, without CorelDRAW.
//
// applyParagraph() below is a line-by-line JavaScript port of ApplyParagraph in
// src/SupertextTranslation.bas. It runs against a mock story whose characters
// carry a style name, with InsertBeforeWide, CopyAttributes and Delete
// simulated by array operations. Keep the port in sync when you change the VBA.
//
// Run: node tests/apply-paragraph-test.js
const assert = (ok, m) => { if (!ok) { console.error("FAIL " + m); process.exitCode = 1; } };
function mkStory(parts) { const ch = []; for (const [t, st] of parts) for (const c of t) ch.push({ c, st }); return ch; }
const text = s => s.map(x => x.c).join("");
const isHard = c => { const n = c.charCodeAt(0); return (n < 32 && ![9, 10, 11, 13].includes(n)) || n === 0xFFFC; };

function applyParagraph(story, para, segs) {
  const txt = para.text, pStart = para.start;
  const hard = []; for (let i = 0; i < txt.length; i++) if (isHard(txt[i])) hard.push(pStart + i);
  const chunks = [[]]; let n = 0;
  for (const seg of segs) {
    if (seg.hard) { n++; if (seg.hard !== "o" + n) return false; chunks.push([]); }
    else if (seg.text.length) chunks[chunks.length - 1].push(seg);
  }
  if (n !== hard.length) return false;
  const loc = para.runs.length ? para.runs.map(r => r.start) : [-1];
  const bounds = [pStart]; for (const h of hard) bounds.push(h, h + 1); bounds.push(pStart + txt.length);
  // Pass 1: insert and format every chunk while all original characters exist.
  const ins = [];
  for (let c = chunks.length - 1; c >= 0; c--) {
    const cs = bounds[2 * c];
    const newText = chunks[c].map(s => s.text).join(""), L = newText.length;
    ins[c] = L;
    if (!L) continue;
    story.splice(cs, 0, ...[...newText].map(ch => ({ c: ch, st: "INSERTED" })));
    for (let k = 0; k < loc.length; k++) if (loc[k] >= cs) loc[k] += L;
    let off = 0;
    for (const seg of chunks[c]) {
      let st = seg.style; if (st < 0 || st >= loc.length) st = 0;
      let src = loc[st]; if (src < 0) src = loc.find(x => x >= 0) ?? -1;
      if (src >= 0) for (let i = 0; i < seg.text.length; i++) story[cs + off + i].st = story[src].st;
      off += seg.text.length;
    }
  }
  // Pass 2: delete the old text, last chunk first.
  let shift = ins.reduce((x, y) => x + y, 0);
  for (let c = chunks.length - 1; c >= 0; c--) {
    const cs = bounds[2 * c], oldLen = bounds[2 * c + 1] - cs;
    if (oldLen > 0) story.splice(cs + shift, oldLen);
    shift -= ins[c];
  }
  return true;
}
function runsOf(story, a, b) { const r = []; for (let i = a; i < b; i++) { const l = r[r.length - 1]; if (l && l.key === story[i].st) l.len++; else r.push({ start: i, len: 1, key: story[i].st }); } return r; }

// Paragraph 2 of a story: "ab " B:"fett" " x" OBJ " end" R:"rot", then paragraph 3
const story = mkStory([["Erster Absatz\r", "N"], ["Jetzt ", "N"], ["fett", "B"], [" und ￼ hier ", "N"], ["rot", "R"], ["\rLetzter", "N"]]);
const t = text(story); const pStart = t.indexOf("Jetzt"); const pEnd = t.indexOf("\rLetzter");
const para = { start: pStart, text: t.slice(pStart, pEnd), runs: runsOf(story, pStart, pEnd) };
// Translation moves the red word before the object and the bold word to the end.
const ok = applyParagraph(story, para, [
  { text: "Maintenant ", style: 0 }, { text: "rouge", style: 3 }, { text: " et ", style: 0 },
  { hard: "o1", text: "", style: 0 }, { text: " ici ", style: 0 }, { text: "gras", style: 1 }]);
const out = text(story);
assert(ok, "apply returned false");
assert(out === "Erster Absatz\rMaintenant rouge et ￼ ici gras\rLetzter", "text: " + JSON.stringify(out));
const st = w => story[out.indexOf(w)].st;
assert(st("rouge") === "R" && st("gras") === "B" && st("Maintenant") === "N" && st(" ici") === "N" && st(" et") === "N", "styles");
assert(story.every(x => x.st !== "INSERTED"), "uncopied formatting");
assert(story[out.indexOf("￼")].st === "N" && story[0].st === "N" && st("Letzter") === "N", "neighbours");
// style whose only source run was in a later, already replaced chunk
const s2 = mkStory([["A", "N"], ["B", "B"], ["￼", "N"], ["C", "R"]]);
applyParagraph(s2, { start: 0, text: text(s2), runs: runsOf(s2, 0, 4) },
  [{ text: "rr", style: 3 }, { text: "x", style: 0 }, { hard: "o1", text: "", style: 0 }, { text: "bb", style: 1 }]);
assert(text(s2) === "rrx￼bb" && s2[0].st === "R" && s2[2].st === "N" && s2[4].st === "B", "cross-chunk: " + JSON.stringify(s2));
// placeholder mismatch
const s3 = mkStory([["A￼B", "N"]]);
assert(applyParagraph(s3, { start: 0, text: "A￼B", runs: runsOf(s3, 0, 3) }, [{ text: "X", style: 0 }]) === false && text(s3) === "A￼B", "mismatch");
console.log(process.exitCode ? "FAILED" : "All checks passed");
