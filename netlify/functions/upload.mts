import type { Config } from "@netlify/functions"
import { timingSafeEqual } from "node:crypto"
import { getStore } from "@netlify/blobs"

// One site-wide store holds the live data set. Strong consistency so a new upload shows immediately.
const store = () => getStore({ name: "council-dashboard", consistency: "strong" })
const LIVE_KEY = "live", PREVIOUS_KEY = "previous"

const MAX_BYTES = 4_000_000

function passcodeOk(req: Request): boolean {
  const expected = Netlify.env.get("UPLOAD_PASSCODE") ?? ""
  const given = req.headers.get("x-upload-passcode") ?? ""
  if (!expected) return false
  const a = Buffer.from(given), b = Buffer.from(expected)
  return a.length === b.length && timingSafeEqual(a, b)
}

// Text from spreadsheets is untrusted: strip markup characters from every string before storing.
function sanitize(v: unknown, depth = 0): unknown {
  if (depth > 12) throw new Error("Data is nested too deeply.")
  if (typeof v === "string") return v.replace(/[<>]/g, (c) => (c === "<" ? "‹" : "›")).slice(0, 2000)
  if (Array.isArray(v)) return v.map((x) => sanitize(x, depth + 1))
  if (v && typeof v === "object") {
    const out: Record<string, unknown> = {}
    for (const [k, x] of Object.entries(v)) out[k] = sanitize(x, depth + 1)
    return out
  }
  return v
}

// Shape check: the dashboard needs these to render. Detailed field checks happen in the browser before upload.
function shapeErrors(d: any): string[] {
  const e: string[] = []
  if (!d || typeof d !== "object") return ["The upload contains no data."]
  if (typeof d.asOf !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(d.asOf)) e.push("Snapshot date is missing.")
  if (!Array.isArray(d.pillars) || d.pillars.length !== 3) e.push("Exactly three pillars are required.")
  if (!Array.isArray(d.initiatives) || d.initiatives.length === 0) e.push("At least one initiative is required.")
  if (!Array.isArray(d.tasks)) e.push("The task list is missing.")
  if (!Array.isArray(d.people)) e.push("The team list is missing.")
  if (!d.info || typeof d.info !== "object") e.push("Metric definitions are missing.")
  if (d.sample !== false) e.push("Only live data can be published.")
  return e
}

export default async (req: Request) => {
  if (!passcodeOk(req)) return Response.json({ errors: ["The upload passcode is incorrect."] }, { status: 401 })
  const s = store()

  if (req.method === "DELETE") {  // remove live data (dashboard falls back to "no live data")
    const current = await s.get(LIVE_KEY, { type: "json" })
    if (current !== null) await s.setJSON(PREVIOUS_KEY, current)
    await s.delete(LIVE_KEY)
    return Response.json({ ok: true })
  }

  const text = await req.text()
  if (text.length > MAX_BYTES) return Response.json({ errors: ["The upload is too large."] }, { status: 413 })
  let body: any
  try { body = JSON.parse(text) } catch { return Response.json({ errors: ["The upload could not be read."] }, { status: 400 }) }
  const errors = shapeErrors(body?.data)
  if (errors.length) return Response.json({ errors }, { status: 422 })

  const record = {
    data: sanitize(body.data),
    meta: sanitize({
      uploadedAt: new Date().toISOString(),
      files: Array.isArray(body.meta?.files) ? body.meta.files.slice(0, 2) : [],
      uploadedBy: typeof body.meta?.uploadedBy === "string" ? body.meta.uploadedBy.slice(0, 80) : "",
      warnings: typeof body.meta?.warnings === "number" ? body.meta.warnings : 0,
    }),
  }
  const current = await s.get(LIVE_KEY, { type: "json" })
  if (current !== null) await s.setJSON(PREVIOUS_KEY, current)  // keep one previous version for rollback
  await s.setJSON(LIVE_KEY, record)
  return Response.json({ ok: true, uploadedAt: (record.meta as any).uploadedAt })
}

export const config: Config = {
  path: "/api/upload",
  method: ["POST", "DELETE"],
  // Slows down passcode guessing.
  rateLimit: { windowLimit: 10, windowSize: 60, aggregateBy: ["ip", "domain"] },
}
