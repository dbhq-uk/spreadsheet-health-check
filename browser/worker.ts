// The Web Worker adapter.
//
// This lives outside src/ on purpose: src/ is the pure engine - no DOM, no fetch, no fs - and
// `self`/`postMessage` are browser APIs. This file is the only place the engine meets the
// browser, and it is deliberately thin enough to read in one sitting.
//
// It makes no network calls of any kind. The bytes arrive from the page, are analysed here,
// and the report goes back. Nothing leaves the machine - the no-upload guarantee holds inside
// the worker exactly as it did on the main thread, and you can confirm that by reading this
// file in the public repo or by watching the network tab while it runs.
//
// The bundle built from this entry point is self-contained (it inlines the engine), so it is a
// single artefact to hash, attest and pin - the site loads exactly this file and nothing else.
import { analyse } from "../src/index";

export interface AnalyseRequest {
  id: number;
  bytes: ArrayBuffer;
}

export type AnalyseResponse =
  | { type: "ready" }
  | { type: "result"; id: number; report: ReturnType<typeof analyse> }
  | { type: "error"; id: number; message: string };

const post = (msg: AnalyseResponse) => (self as unknown as Worker).postMessage(msg);

self.addEventListener("message", (event: MessageEvent<AnalyseRequest>) => {
  const { id, bytes } = event.data ?? ({} as AnalyseRequest);
  try {
    post({ type: "result", id, report: analyse(new Uint8Array(bytes)) });
  } catch (err) {
    // The engine throws on a file it cannot read - password-protected, corrupt, or not a
    // spreadsheet at all. The page turns this into a human sentence; it is not a crash.
    post({ type: "error", id, message: err instanceof Error ? err.message : String(err) });
  }
});

// Tell the page the engine is parsed and ready, so it can warm the worker on load rather than
// paying for the parse at the moment the user drops a file.
post({ type: "ready" });
