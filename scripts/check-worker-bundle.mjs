// Sanity check for the Web Worker bundle - the artefact the website actually loads.
//
// Runs the built IIFE inside a fake worker global scope (self + addEventListener +
// postMessage) with vm.runInThisContext, which is the closest Node equivalent to how a
// browser evaluates a classic worker script. Proves the bundle boots, announces itself
// ready, analyses real bytes and posts a report back - before it ever reaches a browser.
//
//   npm run check:worker
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import vm from "node:vm";

const __dirname = dirname(fileURLToPath(import.meta.url));
const bundle = join(__dirname, "..", "dist", "spreadsheet-health-check.worker.js");
const fixture = join(__dirname, "..", "test", "fixtures", "key-person.xlsx");

const posted = [];
const listeners = [];
const fail = (msg) => { console.error(`FAIL: ${msg}`); process.exit(1); };

// The fake worker global scope.
globalThis.self = globalThis;
globalThis.addEventListener = (type, fn) => { if (type === "message") listeners.push(fn); };
globalThis.postMessage = (msg) => posted.push(msg);

vm.runInThisContext(readFileSync(bundle, "utf8"), { filename: bundle });

if (posted[0]?.type !== "ready") fail(`expected a ready message on boot, got ${JSON.stringify(posted[0])}`);
if (listeners.length !== 1) fail(`expected exactly one message listener, got ${listeners.length}`);

const bytes = new Uint8Array(readFileSync(fixture));
// Hand the worker a detached-style ArrayBuffer copy, as a real postMessage transfer would.
listeners[0]({ data: { id: 1, bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) } });

const reply = posted[1];
if (reply?.type !== "result") fail(`expected a result, got ${JSON.stringify(reply)}`);
if (reply.id !== 1) fail(`reply id ${reply.id} does not match the request`);

const { riskBand, score, findings } = reply.report;
console.log(`worker replied: ${riskBand} (${score}) - ${findings.map((f) => f.id).join(", ")}`);

if (!findings.some((f) => f.id === "key-person.opaque-references")) fail("expected the key-person fixture to report opaque references");
if (findings.some((f) => f.id === "key-person.single-author")) fail("the deleted single-author check is back in the bundle");
if (!findings.every((f) => typeof f.action === "string" && f.action.length > 20)) fail("a finding reached the browser with no action to take");

// An unreadable file must come back as a handled error, not an unhandled throw.
listeners[0]({ data: { id: 2, bytes: new Uint8Array([1, 2, 3, 4]).buffer } });
if (posted[2]?.type !== "error" || posted[2].id !== 2) fail(`expected a handled error for junk bytes, got ${JSON.stringify(posted[2])}`);

console.log("OK: worker bundle boots, analyses, reports actions and handles a bad file.");
