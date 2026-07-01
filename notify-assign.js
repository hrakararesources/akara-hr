// notify-assign.js — Email when a task is assigned
// รันโดย GitHub Actions ทุก ~15 นาที: อ่านแจ้งเตือน type=task_assigned ที่ยังไม่ส่งอีเมล
// (notifications.emailed=false) → ส่งอีเมลถึงพนักงาน แล้ว mark emailed=true กันส่งซ้ำ
// ใช้ SUPABASE_KEY = service_role (bypass RLS) เหมือน notify.js

const https = require('https');
const { createTransport } = require('nodemailer');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;

function sbReq(path, method='GET', body=null) {
  return new Promise((resolve, reject) => {
    const url = new URL(SUPABASE_URL + '/rest/v1/' + path);
    const data = body ? JSON.stringify(body) : null;
    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method,
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': 'Bearer ' + SUPABASE_KEY,
        'Content-Type': 'application/json',
        'Prefer': 'return=minimal',
      }
    };
    const req = https.request(options, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => {
        if (res.statusCode >= 300) return reject(new Error(d || ('HTTP '+res.statusCode)));
        try { resolve(d ? JSON.parse(d) : null); } catch(e) { resolve(null); }
      });
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}
const sbGet = p => sbReq(p, 'GET');

const SMTP_PORT = parseInt(process.env.SMTP_PORT) || 587;
const transporter = createTransport({
  host: process.env.SMTP_HOST,
  port: SMTP_PORT,
  secure: SMTP_PORT === 465,
  requireTLS: SMTP_PORT !== 465,   // MS365 = 587 STARTTLS
  auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
  connectionTimeout: 20000, greetingTimeout: 20000, socketTimeout: 20000,
  tls: { rejectUnauthorized: false }
});

function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

async function main() {
  console.log('🚀 Assign-notification email run...');

  // แจ้งเตือนงานที่เพิ่งมอบหมาย ที่ยังไม่ส่งอีเมล
  const notifs = await sbGet(
    `notifications?type=eq.task_assigned&emailed=eq.false&select=id,member_id,title,body,task_id,created_at&order=created_at`
  );
  if (!Array.isArray(notifs) || !notifs.length) {
    console.log('✅ No new assignment notifications to email.');
    return;
  }
  console.log(`📋 Found ${notifs.length} unemailed assignment notification(s)`);

  const members = await sbGet('members?select=id,name,email');
  const memberMap = {};
  (members||[]).forEach(m => { memberMap[m.id] = m; });

  // group ตามผู้รับ
  const byMember = {};
  notifs.forEach(n => { (byMember[n.member_id] = byMember[n.member_id] || []).push(n); });

  const doneIds = [];
  let sent = 0, failed = 0;

  for (const [memberId, list] of Object.entries(byMember)) {
    const member = memberMap[memberId];
    // ไม่มีอีเมล → mark ว่าประมวลผลแล้ว กันค้างวนซ้ำ
    if (!member || !member.email) {
      list.forEach(n => doneIds.push(n.id));
      console.log(`⚠️  No email for member ${memberId}, skip (${list.length})`);
      continue;
    }

    const rows = list.map(n => `
      <tr>
        <td style="padding:10px 14px;border-bottom:1px solid #eee;font-weight:600">${esc(n.body||n.title||'งานใหม่')}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #eee;color:#6b7899;white-space:nowrap">${new Date(n.created_at).toLocaleDateString('th-TH',{day:'numeric',month:'short'})}</td>
      </tr>`).join('');

    const html = `
<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:'Sarabun','IBM Plex Sans Thai',sans-serif;background:#f2f4f8;margin:0;padding:20px">
  <div style="max-width:600px;margin:0 auto">
    <div style="background:linear-gradient(135deg,#1a3d7a,#2B5BAE);padding:28px 32px;border-radius:12px 12px 0 0;text-align:center">
      <div style="font-size:22px;font-weight:700;color:#fff;letter-spacing:.08em">AKARA RESOURCES</div>
      <div style="font-size:11px;color:rgba(255,255,255,.7);letter-spacing:.14em;margin-top:4px">HUMAN RESOURCES TASK MANAGEMENT</div>
      <div style="width:36px;height:3px;background:#F5C02B;border-radius:2px;margin:12px auto 0"></div>
    </div>
    <div style="background:#fff;padding:28px 32px">
      <div style="font-size:18px;font-weight:600;color:#1a1e2e;margin-bottom:6px">📨 คุณได้รับมอบหมายงานใหม่</div>
      <div style="font-size:14px;color:#6b7899;margin-bottom:20px">
        สวัสดีคุณ <strong style="color:#1a1e2e">${esc(member.name)}</strong> — มี ${list.length} งานที่เพิ่งได้รับมอบหมายให้คุณ
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr style="background:#f2f4f8">
          <th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">งาน</th>
          <th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">วันที่</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div style="margin-top:24px;padding:16px;background:#e8eef8;border-radius:8px;border-left:4px solid #2B5BAE;font-size:13px;color:#6b7899">
        💡 เปิดระบบเพื่อดูรายละเอียดงานและอัปเดตความคืบหน้า
      </div>
      <div style="margin-top:20px;text-align:center">
        <a href="https://hrakararesources.github.io/akara-hr/" style="display:inline-block;padding:12px 28px;background:#2B5BAE;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px">เปิดระบบ HR Task Board →</a>
      </div>
    </div>
    <div style="background:#f2f4f8;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center;font-size:11px;color:#6b7899">
      อีเมลนี้ส่งโดยอัตโนมัติจากระบบ Akara Resources HR Task Management<br>กรุณาอย่าตอบกลับอีเมลนี้
    </div>
  </div>
</body></html>`;

    try {
      await transporter.sendMail({
        from: `"Akara HR System" <${process.env.SMTP_FROM}>`,
        to: member.email,
        subject: `📨 คุณได้รับมอบหมายงานใหม่ (${list.length}) — Akara HR`,
        html,
      });
      list.forEach(n => doneIds.push(n.id));
      console.log(`✅ Sent to ${member.name} <${member.email}> (${list.length})`);
      sent++;
    } catch (e) {
      console.error(`❌ Failed to ${member.email}:`, e.message);
      failed++;
    }
  }

  // mark emailed=true สำหรับที่ส่ง/ข้ามแล้ว (กันส่งซ้ำ)
  if (doneIds.length) {
    const idList = doneIds.map(id => `"${id}"`).join(',');
    await sbReq(`notifications?id=in.(${idList})`, 'PATCH', { emailed: true });
  }
  console.log(`\n📊 Summary: ${sent} email(s) sent, ${failed} failed, ${doneIds.length} marked emailed`);
}

main().catch(err => { console.error('❌ Fatal:', err); process.exit(1); });
