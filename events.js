// events.js — HR Event Reminder Email (แจ้งเตือนกิจกรรมล่วงหน้า 1 วัน)
// รันโดย GitHub Actions ทุกเช้า 08:00 น. — หากิจกรรมใน hr_events ที่จะถึงพรุ่งนี้ → ส่งอีเมลถึง HR ทุกคน
//
// Secrets ที่ต้องมี:
//   SUPABASE_URL, SUPABASE_KEY (service_role)
//   GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_SENDER

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const TENANT = process.env.GRAPH_TENANT_ID;
const CLIENT = process.env.GRAPH_CLIENT_ID;
const SECRET = process.env.GRAPH_CLIENT_SECRET;
const SENDER = process.env.GRAPH_SENDER;

// ── Supabase REST ──
async function sbGet(path) {
  const res = await fetch(SUPABASE_URL + '/rest/v1/' + path, {
    headers: { apikey: SUPABASE_KEY, Authorization: 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json' }
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  const t = await res.text();
  return t ? JSON.parse(t) : null;
}

// ── Microsoft Graph ──
async function graphToken() {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CLIENT, client_secret: SECRET, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  });
  if (!res.ok) throw new Error('token error: ' + await res.text());
  return (await res.json()).access_token;
}
async function sendMail(token, to, subject, html) {
  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(SENDER)}/sendMail`, {
    method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: { subject, body: { contentType: 'HTML', content: html }, toRecipients: [{ emailAddress: { address: to } }] }, saveToSentItems: false }),
  });
  if (res.status !== 202) throw new Error(`sendMail ${res.status}: ${await res.text()}`);
}

function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function formatDate(d) {
  if (!d) return '';
  return new Date(d).toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
}

const TYPE_LABEL = {
  general: 'ทั่วไป', training: '📚 อบรม', meeting: '👥 ประชุม', holiday: '🏖️ วันหยุด',
  company: '🎉 กิจกรรมบริษัท', deadline: '⏰ กำหนดส่ง', birthday: '🎂 วันเกิด', probation: '📋 ครบ Probation',
};

function emailHtml(name, events, dateStr) {
  const rows = events.map(e => `
    <tr>
      <td style="padding:10px 14px;border-bottom:1px solid #eee;font-weight:600">${esc(e.title)}</td>
      <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center">
        <span style="background:#e8eef8;color:#2B5BAE;padding:3px 10px;border-radius:12px;font-size:12px">${esc(TYPE_LABEL[e.type] || e.type)}</span>
      </td>
      <td style="padding:10px 14px;border-bottom:1px solid #eee;color:#6b7899">${esc(e.location || '—')}</td>
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
      <div style="font-size:18px;font-weight:600;color:#1a1e2e;margin-bottom:6px">🗓️ เตือนกิจกรรม HR พรุ่งนี้</div>
      <div style="font-size:14px;color:#6b7899;margin-bottom:20px">
        สวัสดีคุณ <strong style="color:#1a1e2e">${esc(name)}</strong> — พรุ่งนี้ (${dateStr}) มี ${events.length} กิจกรรมในปฏิทิน HR
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr style="background:#f2f4f8">
          <th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">กิจกรรม</th>
          <th style="padding:10px 14px;text-align:center;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">ประเภท</th>
          <th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">สถานที่</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div style="margin-top:20px;text-align:center">
        <a href="https://hrakararesources.github.io/akara-hr/" style="display:inline-block;padding:12px 28px;background:#2B5BAE;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px">เปิดปฏิทินกิจกรรม HR →</a>
      </div>
    </div>
    <div style="background:#f2f4f8;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center;font-size:11px;color:#6b7899">
      อีเมลนี้ส่งโดยอัตโนมัติจากระบบ Akara Resources HR Task Management<br>กรุณาอย่าตอบกลับอีเมลนี้
    </div>
  </div>
</body></html>`;
}

async function main() {
  console.log('🚀 Starting HR event reminder...');

  // วันพรุ่งนี้ตามเวลาไทย (runner เป็น UTC — บวก 7 ชม. ก่อนตัดวัน)
  const nowTh = new Date(Date.now() + 7 * 60 * 60 * 1000);
  nowTh.setUTCDate(nowTh.getUTCDate() + 1);
  const tomorrowStr = nowTh.toISOString().slice(0, 10);
  console.log(`📅 Checking HR events on: ${tomorrowStr}`);

  // กิจกรรมที่เริ่มพรุ่งนี้ หรือกิจกรรมหลายวันที่คร่อมวันพรุ่งนี้อยู่
  const starting = await sbGet(`hr_events?event_date=eq.${tomorrowStr}&select=*`);
  const spanning = await sbGet(`hr_events?event_date=lt.${tomorrowStr}&end_date=gte.${tomorrowStr}&select=*`);
  const events = [...(starting || []), ...(spanning || [])];

  if (!events.length) { console.log('✅ No HR events tomorrow. Nothing to send.'); return; }
  console.log(`📋 Found ${events.length} event(s) tomorrow`);

  const members = await sbGet('members?select=id,name,email');
  const recipients = (members || []).filter(m => m.email);
  if (!recipients.length) { console.log('⚠️  No members with email. Nothing to send.'); return; }

  const dateStr = formatDate(tomorrowStr);
  const token = await graphToken();
  let sent = 0, failed = 0;

  for (const member of recipients) {
    try {
      await sendMail(token, member.email, `🗓️ เตือนกิจกรรม HR พรุ่งนี้ (${events.length}) — Akara HR`, emailHtml(member.name, events, dateStr));
      console.log(`✅ Sent to ${member.name} <${member.email}> (${events.length} event(s))`);
      sent++;
    } catch (e) {
      console.error(`❌ Failed to send to ${member.email}:`, e.message);
      failed++;
    }
  }

  console.log(`\n📊 Summary: ${sent} sent, ${failed} failed`);
}

main().catch(err => { console.error('❌ Fatal error:', err); process.exit(1); });
