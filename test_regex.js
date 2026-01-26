
const BLOCK_ID_PREFIX = "mtd_";
const CHECKLIST_BLOCK_ID_PREFIX = "mtdc_";
const SYNC_MARKER_NAME = "MicrosoftToDoSync";

function buildSyncMarker(blockId) {
  return `<!-- ${SYNC_MARKER_NAME}:${blockId} -->`;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function parseMarkdownTasks(lines) {
  const tasks = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdCaretPattern = /\s+\^([a-z0-9_]+)\s*$/i;
  // The regex from source code
  const blockIdCommentPattern = /\s*<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*([a-z0-9_]+)\s*-->\s*$/i;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = taskPattern.exec(line);
    if (!match) continue;
    const indent = match[1] ?? "";
    const bullet = (match[2] ?? "-");
    const completed = (match[3] ?? " ").toLowerCase() === "x";
    const rest = (match[4] ?? "").trim();
    if (!rest) continue;

    const commentMatch = blockIdCommentPattern.exec(rest);
    const caretMatch = commentMatch ? null : blockIdCaretPattern.exec(rest);
    const markerMatch = commentMatch || caretMatch;
    const existingBlockId = markerMatch ? markerMatch[1] : "";
    const rawTitleWithTag = markerMatch ? rest.slice(0, markerMatch.index).trim() : rest;
    
    tasks.push({
      title: rawTitleWithTag,
      blockId: existingBlockId
    });
  }
  return tasks;
}

function sanitizeTitleForGraph(title) {
  const input = (title || "").trim();
  if (!input) return "";
  const withoutIds = input
    .replace(/\^mtdc?_[a-z0-9_]+/gi, " ")
    .replace(/<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*mtdc?_[a-z0-9_]+\s*-->/gi, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
  return withoutIds;
}

// Test cases
const lines = [
  "- [ ] Task 1 <!-- mtd:mtd_12345 -->",
  "- [ ] Task 2 <!-- MicrosoftToDoSync:mtd_67890 -->",
  "- [ ] Task 3 ^mtd_abcde",
  "- [ ] Task 4 <!-- MicrosoftToDoSync:mtd_mixed123 -->  ",
  "- [ ] Task 5 <!-- mtd:mtd_oldstyle -->",
  "- [ ] Task 6"
];

console.log("--- Testing Parse ---");
const parsed = parseMarkdownTasks(lines);
console.log(JSON.stringify(parsed, null, 2));

console.log("\n--- Testing Sanitize ---");
parsed.forEach(t => {
  const original = t.title + (t.blockId ? ` <!-- ...:${t.blockId} -->` : ""); // Rough reconstruction
  // Actually we should test sanitize on the full line content minus the checkbox
  // But sanitizeTitleForGraph is usually called on the TITLE part (which might include the marker if parse failed?)
  // No, sanitizeTitleForGraph is called on the raw input from Graph or local title.
  // Wait, if parseMarkdownTasks extracts the title correctly, `t.title` should NOT contain the marker.
  console.log(`Title: "${t.title}", Sanitized: "${sanitizeTitleForGraph(t.title)}"`);
  
  // Test sanitizing a string that HAS the marker (simulating dirty input)
  const dirty = `${t.title} <!-- MicrosoftToDoSync:${t.blockId || 'none'} -->`;
  console.log(`Dirty: "${dirty}", Cleaned: "${sanitizeTitleForGraph(dirty)}"`);
});

const legacyDirty = "Task with legacy <!-- mtd:mtd_legacy -->";
console.log(`Legacy Dirty: "${legacyDirty}", Cleaned: "${sanitizeTitleForGraph(legacyDirty)}"`);

