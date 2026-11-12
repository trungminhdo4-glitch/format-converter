// SafeGit plugin for OpenCode - thin adapter, no policy of its own.
// Every tool call is classified by the SafeGit CLI before it runs.
// SafeGit decides; this file only transports the decision.
// Dual export: OpenCode 2.x calls setup(ctx); OpenCode 1.x calls server(input).
// Fail-closed: unknown project dir, spawn error, timeout or non-zero exit => refused.
import { spawnSync } from "node:child_process"
import path from "node:path"

const SAFEGIT_ARGV = process.env.SAFEGIT_CLI ? [process.env.SAFEGIT_CLI] : ["C:\\Users\\trung\\.agent-runtime\\safegit-venv\\Scripts\\python.exe", "-P", "-m", "safegit"]
const SAFEGIT_PINNED_ARGS = ["--policy", "C:\\Users\\trung\\Dev\\products\\format-converter\\.safegit\\safegit.policy.json", "--approval-store", "C:\\Users\\trung\\.safegit\\approvals.json", "--git", "C:\\Program Files\\Git\\mingw64\\libexec\\git-core\\git.exe"]

function gate(cwd, call) {
  if (!cwd) throw new Error("SafeGit: project directory unknown. Refusing the tool call.")
  const payload = JSON.stringify({ ...call, cwd, args: call.args ?? {} })
  const result = spawnSync(
    SAFEGIT_ARGV[0],
    [...SAFEGIT_ARGV.slice(1), "hook", "opencode", "--non-interactive", ...SAFEGIT_PINNED_ARGS],
    { input: payload, encoding: "utf8", timeout: 20000, windowsHide: true },
  )
  if (result.error) {
    throw new Error(`SafeGit could not run (${result.error.message}). Refusing the tool call.`)
  }
  if (result.status === 0) return
  const reason = (result.stderr || "").trim() || `SafeGit blocked this call (exit ${result.status})`
  throw new Error(reason)
}

function within(dir, root) {
  const norm = (p) => {
    const r = path.resolve(p)
    return process.platform === "win32" ? r.toLowerCase() : r
  }
  const rel = path.relative(norm(root), norm(dir))
  return rel === "" || (!rel.startsWith("..") && !path.isAbsolute(rel))
}

export default {
  id: "safegit",
  // OpenCode 2.x: the plugin runs in a background service whose process.cwd() is not the
  // project. ctx.location is where this plugin instance loaded, not necessarily where a
  // session runs, so each call is gated in its session's directory, and only if that lies
  // inside this instance's project (the pinned policy belongs to it). Throwing aborts the call.
  async setup(ctx) {
    const root = ctx?.location?.directory
    await ctx.tool.hook("execute.before", async (event) => {
      let session = null
      try {
        session = await ctx.session.get({ sessionID: event.sessionID })
      } catch (err) {
        throw new Error(`SafeGit: session lookup failed (${err?.message}). Refusing the tool call.`)
      }
      const cwd = session?.location?.directory
      if (!root || !cwd) throw new Error("SafeGit: project directory unknown. Refusing the tool call.")
      if (!within(cwd, root)) {
        throw new Error(`SafeGit: session directory ${cwd} is outside the plugin project ${root}. Refusing the tool call.`)
      }
      gate(cwd, { tool: event.tool, sessionID: event.sessionID, callID: event.id, args: event.input })
    })
  },
  // OpenCode 1.x
  async server({ directory }) {
    return {
      "tool.execute.before": async (input, output) => {
        gate(directory, { tool: input.tool, sessionID: input.sessionID, callID: input.callID, args: output.args })
      },
    }
  },
}
