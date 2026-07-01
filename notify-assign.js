// notify-assign.js — Email when a task is assigned (via Microsoft Graph API)
// รันโดย GitHub Actions ทุก ~15 นาที: อ่านแจ้งเตือน type=task_assigned ที่ยังไม่ส่งอีเมล
// (notifications.emailed=false) → ส่งอีเมลถึงพนักงานผ่าน MS365 Graph → mark emailed=true
//
// Secrets ที่ต้องมี:
//   SUPABASE_URL, SUPABASE_KEY (service_role)
//   GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_SENDER
//   (ชุด GRAPH_* = Azure AD app เดียวกับ Health Benefit, สิทธิ์ Mail.Send;
//    GRAPH_SENDER = hr-notify@akararesources.com)

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const TENANT = process.env.GRAPH_TENANT_ID;
const CLIENT = process.env.GRAPH_CLIENT_ID;
const SECRET = process.env.GRAPH_CLIENT_SECRET;
const SENDER = process.env.GRAPH_SENDER;

// ── Supabase REST ──
async function sb(path, method='GET', body=null) {
  const res = await fetch(SUPABASE_URL + '/rest/v1/' + path, {
    method,
    headers: {
      apikey: SUPABASE_KEY, Authorization: 'Bearer ' + SUPABASE_KEY,
      'Content-Type': 'application/json', Prefer: 'return=minimal',
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const t = await res.text();
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${t}`);
  return t ? JSON.parse(t) : null;
}

// ── Microsoft Graph ──
async function graphToken() {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT, client_secret: SECRET,
      scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials',
    }),
  });
  if (!res.ok) throw new Error('token error: ' + await res.text());
  return (await res.json()).access_token;
}
async function sendMail(token, to, subject, html) {
  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(SENDER)}/sendMail`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: { subject, body: { contentType: 'HTML', content: html }, toRecipients: [{ emailAddress: { address: to } }] },
      saveToSentItems: false,
    }),
  });
  if (res.status !== 202) throw new Error(`sendMail ${res.status}: ${await res.text()}`);
}

function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function emailHtml(name, list) {
  const rows = list.map(n => `
    <tr>
      <td style="padding:10px 14px;border-bottom:1px solid #eee;font-weight:600">${esc(n.body||n.title||'งานใหม่')}</td>
      <td style="padding:10px 14px;border-bottom:1px solid #eee;color:#6b7899;white-space:nowrap">${new Date(n.created_at).toLocaleDateString('th-TH',{day:'numeric',month:'short'})}</td>
    </tr>`).join('');
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
      <div style="font-size:14px;color:#6b7899;margin-bottom:20px">สวัสดีคุณ <strong style="color:#1a1e2e">${esc(name)}</strong> — มี ${list.length} งานที่เพิ่งได้รับมอบหมายให้คุณ</div>
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr style="background:#f2f4f8">
          <th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">งาน</th>
          <th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">วันที่</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div style="margin-top:20px;text-align:center">
        <a href="https://hrakararesources.github.io/akara-hr/" style="display:inline-block;padding:12px 28px;background:#2B5BAE;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px">เปิดระบบ HR Task Board →</a>
      </div>
    </div>
    <div style="background:#f2f4f8;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center;font-size:11px;color:#6b7899">
      อีเมลนี้ส่งโดยอัตโนมัติจากระบบ Akara Resources HR Task Management<br>กรุณาอย่าตอบกลับอีเมลนี้
    </div>
  </div>
</body></html>`;
}

async function main() {
  console.log('🚀 Assign-notification email (Graph) run...');
  const notifs = await sb(`notifications?type=eq.task_assigned&emailed=eq.false&select=id,member_id,title,body,task_id,created_at&order=created_at`);
  if (!Array.isArray(notifs) || !notifs.length) { console.log('✅ No new assignment notifications to email.'); return; }
  console.log(`📋 Found ${notifs.length} unemailed assignment notification(s)`);

  const members = await sb('members?select=id,name,email');
  const memberMap = {};
  (members||[]).forEach(m => { memberMap[m.id] = m; });

  const byMember = {};
  notifs.forEach(n => { (byMember[n.member_id] = byMember[n.member_id] || []).push(n); });

  const token = await graphToken();
  const doneIds = [];
  let sent = 0, failed = 0;

  for (const [memberId, list] of Object.entries(byMember)) {
    const member = memberMap[memberId];
    if (!member || !member.email) { list.forEach(n => doneIds.push(n.id)); console.log(`⚠️  No email for member ${memberId}, skip (${list.length})`); continue; }
    try {
      await sendMail(token, member.email, `📨 คุณได้รับมอบหมายงานใหม่ (${list.length}) — Akara HR`, emailHtml(member.name, list));
      list.forEach(n => doneIds.push(n.id));
      console.log(`✅ Sent to ${member.name} <${member.email}> (${list.length})`);
      sent++;
    } catch (e) { console.error(`❌ Failed to ${member.email}:`, e.message); failed++; }
  }

  if (doneIds.length) {
    const idList = doneIds.map(id => `"${id}"`).join(',');
    await sb(`notifications?id=in.(${idList})`, 'PATCH', { emailed: true });
  }
  console.log(`\n📊 Summary: ${sent} sent, ${failed} failed, ${doneIds.length} marked emailed`);
}

main().catch(err => { console.error('❌ Fatal:', err); process.exit(1); });
