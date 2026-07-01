// ============================================================================
// notify-assign — ส่งอีเมล "ทันที" เมื่อถูกมอบหมายงาน
//   ทริกโดย Database Webhook: ตาราง public.notifications · INSERT
//   ส่งผ่าน Microsoft 365 (Microsoft Graph API · client credentials)
//   แล้ว set notifications.emailed = true (กัน cron ส่งซ้ำ)
//
// การติดตั้ง:
//   1) Secrets (Edge Function):
//        GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_SENDER,
//        NOTIFY_HOOK_SECRET
//        (SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY มีให้อัตโนมัติใน Edge Function)
//   2) Deploy:  supabase functions deploy notify-assign --no-verify-jwt
//   3) Database Webhook: Database → Webhooks → New
//        - Table: notifications · Events: Insert
//        - Type: HTTP Request → POST → URL ของฟังก์ชันนี้
//        - HTTP Header: x-hook-secret = <NOTIFY_HOOK_SECRET>
// ============================================================================
import { createClient } from "https://esm.sh/@supabase/supabase-js@2"

const TENANT = Deno.env.get("GRAPH_TENANT_ID")!
const CLIENT = Deno.env.get("GRAPH_CLIENT_ID")!
const SECRET = Deno.env.get("GRAPH_CLIENT_SECRET")!
const SENDER = Deno.env.get("GRAPH_SENDER")!
const HOOK_SECRET = Deno.env.get("NOTIFY_HOOK_SECRET") ?? ""
const APP_URL = "https://hrakararesources.github.io/akara-hr/"

const esc = (s: unknown) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")

async function graphToken(): Promise<string> {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: CLIENT, client_secret: SECRET,
      scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials",
    }),
  })
  if (!res.ok) throw new Error("token error: " + await res.text())
  return (await res.json()).access_token
}

function emailHtml(name: string, n: any): string {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
  <body style="font-family:'Sarabun','IBM Plex Sans Thai',Arial,sans-serif;background:#f2f4f8;margin:0;padding:20px">
    <div style="max-width:600px;margin:0 auto">
      <div style="background:linear-gradient(135deg,#1a3d7a,#2B5BAE);padding:28px 32px;border-radius:12px 12px 0 0;text-align:center">
        <div style="font-size:22px;font-weight:700;color:#fff;letter-spacing:.08em">AKARA RESOURCES</div>
        <div style="font-size:11px;color:rgba(255,255,255,.7);letter-spacing:.14em;margin-top:4px">HUMAN RESOURCES TASK MANAGEMENT</div>
        <div style="width:36px;height:3px;background:#F5C02B;border-radius:2px;margin:12px auto 0"></div>
      </div>
      <div style="background:#fff;padding:28px 32px">
        <div style="font-size:18px;font-weight:600;color:#1a1e2e;margin-bottom:6px">📨 คุณได้รับมอบหมายงานใหม่</div>
        <div style="font-size:14px;color:#6b7899;margin-bottom:16px">สวัสดีคุณ <strong style="color:#1a1e2e">${esc(name)}</strong></div>
        <div style="padding:14px 16px;background:#f2f4f8;border-radius:8px;border-left:4px solid #2B5BAE;font-size:14px;color:#1a2b4a;font-weight:600">${esc(n.body || n.title || "งานใหม่")}</div>
        <div style="margin-top:20px;text-align:center">
          <a href="${APP_URL}" style="display:inline-block;padding:12px 28px;background:#2B5BAE;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px">เปิดระบบ HR Task Board →</a>
        </div>
      </div>
      <div style="background:#f2f4f8;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center;font-size:11px;color:#6b7899">
        อีเมลนี้ส่งโดยอัตโนมัติจากระบบ Akara Resources HR Task Management<br>กรุณาอย่าตอบกลับอีเมลนี้
      </div>
    </div>
  </body></html>`
}

Deno.serve(async (req) => {
  try {
    if (HOOK_SECRET && req.headers.get("x-hook-secret") !== HOOK_SECRET) {
      return new Response("unauthorized", { status: 401 })
    }
    const body = await req.json()
    const n = body?.record ?? body
    if (!n || n.type !== "task_assigned" || !n.member_id) return new Response("skip", { status: 200 })
    if (n.emailed === true) return new Response("already emailed", { status: 200 })

    const sb = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!)
    const { data: m } = await sb.from("members").select("name,email").eq("id", n.member_id).single()
    if (!m?.email) {
      await sb.from("notifications").update({ emailed: true }).eq("id", n.id)
      return new Response("no email", { status: 200 })
    }

    const token = await graphToken()
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(SENDER)}/sendMail`,
      {
        method: "POST",
        headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          message: {
            subject: "📨 คุณได้รับมอบหมายงานใหม่ — Akara HR",
            body: { contentType: "HTML", content: emailHtml(m.name || "", n) },
            toRecipients: [{ emailAddress: { address: m.email } }],
          },
          saveToSentItems: false,
        }),
      },
    )
    if (res.status !== 202) {
      const err = await res.text()
      console.error("sendMail error:", res.status, err)
      return new Response(err, { status: 500 })
    }
    await sb.from("notifications").update({ emailed: true }).eq("id", n.id)
    return new Response("sent", { status: 200 })
  } catch (e) {
    console.error(e)
    return new Response(String(e), { status: 500 })
  }
})
