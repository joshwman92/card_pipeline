const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const contentPath = path.join(__dirname, "..", "cardladder-autocomp", "extension", "content.js");
const source = fs.readFileSync(contentPath, "utf8");
const backgroundPath = path.join(__dirname, "..", "cardladder-autocomp", "extension", "background.js");
const backgroundSource = fs.readFileSync(backgroundPath, "utf8");

function functionSource(name) {
  const start = source.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`Missing function ${name}`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    const char = source[index];
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(start, index + 1);
    }
  }
  throw new Error(`Could not extract function ${name}`);
}

function backgroundFunctionSource(name) {
  const start = backgroundSource.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`Missing background function ${name}`);
  const bodyStart = backgroundSource.indexOf("{", start);
  let depth = 0;
  for (let index = bodyStart; index < backgroundSource.length; index += 1) {
    const char = backgroundSource[index];
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) return backgroundSource.slice(start, index + 1);
    }
  }
  throw new Error(`Could not extract background function ${name}`);
}

const parserCode = `
const COMP_SOURCE_LABELS = ["eBay", "Fanatics", "Card Ladder"];
${functionSource("sourceLabelToPattern")}
const COMP_SOURCE_PATTERN_TEXT = COMP_SOURCE_LABELS
  .map(sourceLabelToPattern)
  .sort((a, b) => b.length - a.length)
  .join("|");
const COMP_SOURCE_PATTERN = new RegExp("\\\\b(" + COMP_SOURCE_PATTERN_TEXT + ")\\\\b", "i");
${functionSource("compDatePattern")}
${functionSource("sourceLineMatch")}
${functionSource("currentCompChunkOnly")}
${functionSource("cleanCompTitle")}
${functionSource("parseCompChunk")}
${functionSource("parseMoneyValue")}
${functionSource("escapeRegExp")}
${functionSource("visibleTextMatchesCert")}
${functionSource("cleanProfileTitle")}
${functionSource("extractProfileFromText")}
${functionSource("shouldAcceptStableResultAfterFreshRetry")}
${backgroundFunctionSource("isRecoverableResultWaitFailure")}
${backgroundFunctionSource("profileIdFromPageUrl")}
${backgroundFunctionSource("isBareProfileIdTitle")}
globalThis.parseCompChunk = parseCompChunk;
globalThis.sourceLineMatch = sourceLineMatch;
globalThis.parseMoneyValue = parseMoneyValue;
globalThis.visibleTextMatchesCert = visibleTextMatchesCert;
globalThis.extractProfileFromText = extractProfileFromText;
globalThis.shouldAcceptStableResultAfterFreshRetry = shouldAcceptStableResultAfterFreshRetry;
globalThis.isRecoverableResultWaitFailure = isRecoverableResultWaitFailure;
globalThis.profileIdFromPageUrl = profileIdFromPageUrl;
globalThis.isBareProfileIdTitle = isBareProfileIdTitle;
`;

const context = {};
vm.createContext(context);
vm.runInContext(parserCode, context);

const chunk = [
  "eBay 2014 Panini Flawless Greats Patches Autographs Gold #22 Joe Montana BGS 9",
  "Jun 1, 2026 Auction $25.00",
  "Fanatics 2022 Panini Donruss Chet Holmgren PSA 9",
  "Jun 2, 2026 Buy It Now $90.00",
].join(" ");

const comp = context.parseCompChunk(chunk, context.sourceLineMatch("eBay"));

assert(comp, "Expected first comp to parse");
assert.strictEqual(comp.source, "EBAY");
assert.strictEqual(comp.date_sold, "Jun 1, 2026");
assert.strictEqual(comp.price, "$25.00");
assert.match(comp.title, /Joe Montana/);
assert.doesNotMatch(comp.title, /Chet Holmgren/);
assert.strictEqual(context.parseMoneyValue("$20.27k"), 20270);
assert.strictEqual(context.parseMoneyValue("$20.27"), 20.27);
assert.strictEqual(
  context.visibleTextMatchesCert("Cert # 121465724 Grade: GEM MT 10 Profile: 2024 Topps Chrome", "121465724"),
  true,
);
assert.strictEqual(
  context.visibleTextMatchesCert("Cert 0007 434 213 Grade: BGS 9.5", "0007434213"),
  true,
);
assert.strictEqual(
  context.visibleTextMatchesCert("Cert # 121465725 Grade: GEM MT 10", "121465724"),
  false,
);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(context.extractProfileFromText(
    "254 results Grade: 10, Grader: PSA, Profile: 2016 Pokemon Japanese XY PokeKyun Collection 019 Full Art/Gardevoir Ex 1st Edition (Pop 2239) x CL Value $990.00",
  ))),
  {
    grade: "10",
    grader: "PSA",
    title: "2016 Pokemon Japanese XY PokeKyun Collection 019 Full Art/Gardevoir Ex 1st Edition",
  },
);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(context.extractProfileFromText(
    "2 results Grade: 9, Grader: PSA, Profile: 2025 Topps Chrome X Cactus Jack 15 Aaron Judge Gold Refractor (Pop 8) x CL Value $196.00",
  ))),
  {
    grade: "9",
    grader: "PSA",
    title: "2025 Topps Chrome X Cactus Jack 15 Aaron Judge Gold Refractor",
  },
);
assert.strictEqual(
  context.shouldAcceptStableResultAfterFreshRetry({ cardTitle: "", allowStableResultAfterFreshRetry: true }),
  true,
);
assert.strictEqual(
  context.shouldAcceptStableResultAfterFreshRetry({ cardTitle: "Mega Venusaur", allowStableResultAfterFreshRetry: true }),
  false,
);
assert.strictEqual(
  context.isRecoverableResultWaitFailure({
    status: "error",
    error: "Card Ladder did not load a new matching result after submit.",
    pageUrl: "https://app.cardladder.com/sales-history?filters=grader%3Apsa%7CprofileId%3Apsa-12168705",
  }),
  true,
);
assert.strictEqual(
  context.isRecoverableResultWaitFailure({
    status: "error",
    error: "Card Ladder did not load a new matching result after submit.",
    pageUrl: "https://app.cardladder.com/sales-history",
  }),
  false,
);
assert.strictEqual(
  context.profileIdFromPageUrl(
    "https://app.cardladder.com/sales-history?filters=grader%3Apsa%7Cgrade%3Ag8%7CprofileId%3Apsa-6029173",
  ),
  "psa-6029173",
);
assert.strictEqual(context.isBareProfileIdTitle("psa-6029173", "PSA-6029173"), true);
assert.strictEqual(
  context.isBareProfileIdTitle("2021 Pokemon Swsh Black Star Promo #136 Mimikyu-Holo", "psa-6029173"),
  false,
);

console.log("extension parser regression ok");
