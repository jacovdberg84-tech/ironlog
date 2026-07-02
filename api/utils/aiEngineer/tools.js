import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { isAllowedRelativePath } from "./config.js";

function runCmd(command, args, cwd, timeoutMs = 60000) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    timeout: timeoutMs,
    shell: false,
    windowsHide: true,
  });
  return {
    command: `${command} ${args.join(" ")}`,
    cwd,
    exit_code: typeof result.status === "number" ? result.status : 1,
    stdout: String(result.stdout || "").trim(),
    stderr: String(result.stderr || "").trim(),
  };
}

function normalizeRel(relPath) {
  return String(relPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
}

function resolveSafe(baseRoot, relPath) {
  const rel = normalizeRel(relPath);
  if (!isAllowedRelativePath(rel)) {
    throw new Error(`Path not allowed: ${relPath}`);
  }
  const abs = path.resolve(baseRoot, rel);
  const root = path.resolve(baseRoot);
  if (!abs.startsWith(root)) {
    throw new Error(`Path escapes sandbox: ${relPath}`);
  }
  return { rel, abs };
}

function readFileSafe(absPath, maxChars = 16000) {
  try {
    const content = fs.readFileSync(absPath, "utf8");
    return content.length > maxChars ? `${content.slice(0, maxChars)}\n/* truncated */` : content;
  } catch (err) {
    return `(unreadable: ${String(err?.message || err)})`;
  }
}

function collectEditableFiles(absDir, relDir, out, max = 400) {
  if (out.length >= max) return;
  let entries = [];
  try {
    entries = fs.readdirSync(absDir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const ent of entries) {
    if (out.length >= max) break;
    if (ent.name === "node_modules" || ent.name === ".git") continue;
    const rel = relDir ? `${relDir}/${ent.name}` : ent.name;
    if (ent.isDirectory()) {
      collectEditableFiles(path.join(absDir, ent.name), rel, out, max);
    } else if (isAllowedRelativePath(rel) && /\.(js|html|css|json|md)$/i.test(ent.name)) {
      out.push(rel);
    }
  }
}

function walkFiles(repoRoot, max = 400) {
  const out = [];
  const roots = ["web", path.join("api", "routes"), path.join("api", "utils")];
  for (const root of roots) {
    const abs = path.join(repoRoot, root);
    if (!fs.existsSync(abs)) continue;
    collectEditableFiles(abs, root.replace(/\\/g, "/"), out, max);
  }
  const serverJs = path.join(repoRoot, "api", "server.js");
  if (fs.existsSync(serverJs)) out.push("api/server.js");
  return [...new Set(out)];
}

function searchInTree(baseRoot, pattern, globHint = "") {
  const needle = String(pattern || "").toLowerCase();
  if (!needle) return [];
  const files = walkFiles(baseRoot);
  const filtered = globHint
    ? files.filter((f) => f.includes(String(globHint).replace(/\\/g, "/").replace(/^\*\*\//, "")))
    : files;
  const hits = [];
  for (const rel of filtered) {
    if (hits.length >= 40) break;
    try {
      const { abs } = resolveSafe(baseRoot, rel);
      const body = fs.readFileSync(abs, "utf8");
      const lines = body.split(/\r?\n/);
      for (let i = 0; i < lines.length; i++) {
        if (lines[i].toLowerCase().includes(needle)) {
          hits.push({
            file: rel,
            line: i + 1,
            text: lines[i].trim().slice(0, 200),
          });
          if (hits.length >= 40) break;
        }
      }
    } catch {}
  }
  return hits;
}

export function createToolRegistry({ repoRoot, worktreePath = null }) {
  const readRoot = worktreePath || repoRoot;
  const state = { finalizedDiff: null, tool_trace: [] };

  function trace(name, input, output) {
    state.tool_trace.push({
      tool: name,
      input,
      output_preview: typeof output === "string" ? output.slice(0, 500) : output,
      at: new Date().toISOString(),
    });
  }

  const handlers = {
    list_files({ glob = "" } = {}) {
      const files = walkFiles(readRoot).filter((f) => !glob || f.includes(String(glob).replace(/\\/g, "/")));
      const out = files.slice(0, 120);
      trace("list_files", { glob }, out);
      return out;
    },

    read_file({ path: relPath } = {}) {
      const { rel, abs } = resolveSafe(readRoot, relPath);
      const content = readFileSafe(abs);
      trace("read_file", { path: rel }, content.slice(0, 200));
      return { path: rel, content };
    },

    search_repo({ pattern, glob = "" } = {}) {
      const rg = runCmd(
        "rg",
        ["-n", "--no-heading", "-m", "20", String(pattern || ""), ...(glob ? ["--glob", glob] : []), "web", "api/routes", "api/utils"],
        readRoot,
      );
      if (rg.exit_code === 0 && rg.stdout) {
        const lines = rg.stdout.split(/\r?\n/).filter(Boolean).slice(0, 40);
        trace("search_repo", { pattern, glob }, lines);
        return lines;
      }
      const hits = searchInTree(readRoot, pattern, glob);
      trace("search_repo", { pattern, glob }, hits);
      return hits;
    },

    git_apply_check({ diff } = {}) {
      if (!worktreePath) {
        return { ok: false, error: "git_apply_check requires a worktree" };
      }
      const patchDir = path.join(worktreePath, ".ai-engineer", "patches");
      fs.mkdirSync(patchDir, { recursive: true });
      const tmpRel = ".ai-engineer/patches/_agent-check.patch";
      const tmpAbs = path.join(worktreePath, tmpRel);
      const unified = String(diff || "").trim();
      if (!unified.startsWith("diff --git")) {
        return { ok: false, error: "Diff must start with diff --git" };
      }
      fs.writeFileSync(tmpAbs, `${unified}\n`, "utf8");
      const step = runCmd("git", ["apply", "--check", tmpRel], worktreePath);
      const ok = step.exit_code === 0;
      if (ok) state.finalizedDiff = unified;
      const out = {
        ok,
        stderr: step.stderr || null,
        stdout: step.stdout || null,
      };
      trace("git_apply_check", { bytes: unified.length }, out);
      return out;
    },
  };

  return {
    state,
    handlers,
    openAiToolDefs: [
      {
        type: "function",
        function: {
          name: "list_files",
          description: "List editable repo files under web/ and api/",
          parameters: {
            type: "object",
            properties: {
              glob: { type: "string", description: "Optional substring filter, e.g. maintenance" },
            },
          },
        },
      },
      {
        type: "function",
        function: {
          name: "read_file",
          description: "Read a file from the repo sandbox",
          parameters: {
            type: "object",
            properties: {
              path: { type: "string", description: "Relative path e.g. web/maintenance.js" },
            },
            required: ["path"],
          },
        },
      },
      {
        type: "function",
        function: {
          name: "search_repo",
          description: "Search code for a pattern in allowed folders",
          parameters: {
            type: "object",
            properties: {
              pattern: { type: "string" },
              glob: { type: "string" },
            },
            required: ["pattern"],
          },
        },
      },
      {
        type: "function",
        function: {
          name: "git_apply_check",
          description: "Validate a unified diff with git apply --check. Call when the patch is ready.",
          parameters: {
            type: "object",
            properties: {
              diff: { type: "string", description: "Full unified diff starting with diff --git" },
            },
            required: ["diff"],
          },
        },
      },
    ],
  };
}

export function summarizeRepoTree(repoRoot) {
  return walkFiles(repoRoot).slice(0, 100);
}
