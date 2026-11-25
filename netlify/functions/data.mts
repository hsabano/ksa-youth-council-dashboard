import type { Config } from "@netlify/functions"
import { getStore } from "@netlify/blobs"

// One site-wide store holds the live data set. Strong consistency so a new upload shows immediately.
const store = () => getStore({ name: "council-dashboard", consistency: "strong" })
const LIVE_KEY = "live"

// GET /api/data → the current live data set, or 404 when nothing has been published yet.
export default async () => {
  const live = await store().get(LIVE_KEY, { type: "json" })
  if (live === null) return Response.json({ error: "No live data has been published yet." }, { status: 404 })
  return Response.json(live, { headers: { "cache-control": "no-store" } })
}

export const config: Config = { path: "/api/data", method: "GET" }
