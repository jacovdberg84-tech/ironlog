// IRONLOG/api/utils/borrisQueue.js — Borris works on costing proposals one at a
// time in the background. A slow local model then never holds a web request open
// (proxies cut those at ~100 s) and never runs twice at once on a small server.
// The screen polls for the result.

/**
 * run(key) resolves to the result for that key (or throws). Results are kept for
 * keepMs so a reopened panel can pick them up.
 */
export function createBorrisQueue({ run, keepMs = 30 * 60 * 1000, maxWaiting = 5, now = () => Date.now() }) {
  const jobs = new Map();
  const waiting = [];
  let running = false;

  function prune() {
    const cutoff = now() - keepMs;
    for (const [key, job] of jobs) {
      if (job.finished_at && job.finished_at < cutoff) jobs.delete(key);
    }
  }

  async function drain() {
    if (running) return;
    running = true;
    try {
      while (waiting.length) {
        const key = waiting.shift();
        const job = jobs.get(key);
        if (!job || job.status !== "queued") continue;
        job.status = "working";
        job.started_at = now();
        try {
          job.result = await run(key);
          job.status = "done";
        } catch (err) {
          job.status = "failed";
          job.error = String(err?.message || err).slice(0, 300);
        }
        job.finished_at = now();
      }
    } finally {
      running = false;
    }
  }

  return {
    /** Queue a key unless it is already queued or running; refresh re-runs a finished one. */
    enqueue(key, { refresh = false } = {}) {
      prune();
      const existing = jobs.get(key);
      if (existing && ["queued", "working"].includes(existing.status)) return this.status(key);
      if (existing && existing.status === "done" && !refresh) return this.status(key);
      if (waiting.length >= maxWaiting) return { status: "busy", error: "Borris has too many requests waiting; try again shortly" };
      jobs.set(key, { status: "queued", queued_at: now(), result: null, error: null });
      waiting.push(key);
      void drain();
      return this.status(key);
    },
    status(key) {
      const job = jobs.get(key);
      if (!job) return { status: "none" };
      const ahead = job.status === "queued" ? waiting.indexOf(key) + (running ? 1 : 0) : 0;
      return {
        status: job.status,
        ahead,
        seconds: Math.round(((job.finished_at || now()) - (job.started_at || job.queued_at)) / 1000),
        result: job.result,
        error: job.error,
      };
    },
    /** Waits for the queue to go idle (tests). */
    async idle() {
      while (running || waiting.length) await new Promise((r) => setTimeout(r, 5));
    },
  };
}
