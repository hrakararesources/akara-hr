// notify.js — Task Due Date Email Notification
// รันโดย GitHub Actions ทุกเช้า 08:00 น.

const https = require('https');
const { createTransport } = require('nodemailer');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;

// ── FETCH จาก Supabase ──
function sbGet(path) {
  return new Promise((resolve, reject) => {
    const url = new URL(SUPABASE_URL + '/rest/v1/' + path);
    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: 'GET',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': 'Bearer ' + SUPABASE_KEY,
        'Content-Type': 'application/json',
      }
    };
    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch(e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// ── EMAIL TRANSPORTER ──
const transporter = createTransport({
  host: process.env.SMTP_HOST,
  port: parseInt(process.env.SMTP_PORT) || 587,
  secure: parseInt(process.env.SMTP_PORT) === 465,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
  tls: { rejectUnauthorized: false }
});

function formatDate(d) {
  if (!d) return '';
  return new Date(d).toLocaleDateString('th-TH', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  });
}

function priorityLabel(p) {
  return { 'เร่งด่วน': '🔴 เร่งด่วน', 'สูง': '🟠 สูง', 'ปานกลาง': '🔵 ปานกลาง', 'ต่ำ': '⚪ ต่ำ' }[p] || p;
}

// ── MAIN ──
async function main() {
  console.log('🚀 Starting task due date notification...');

  // วันพรุ่งนี้
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = tomorrow.toISOString().slice(0, 10);
  console.log(`📅 Checking tasks due on: ${tomorrowStr}`);

  // ดึง tasks ที่ครบกำหนดพรุ่งนี้ และยังไม่เสร็จ
  const tasks = await sbGet(
    `tasks?due=eq.${tomorrowStr}&status=neq.เสร็จสิ้น&select=*`
  );

  if (!tasks || !tasks.length) {
    console.log('✅ No tasks due tomorrow. Nothing to send.');
    return;
  }

  console.log(`📋 Found ${tasks.length} task(s) due tomorrow`);

  // ดึง members ทั้งหมด
  const members = await sbGet('members?select=*');
  const memberMap = {};
  members.forEach(m => { memberMap[m.id] = m; });

  // จัดกลุ่ม tasks ตาม member
  const tasksByMember = {};
  tasks.forEach(t => {
    if (!tasksByMember[t.member_id]) tasksByMember[t.member_id] = [];
    tasksByMember[t.member_id].push(t);
  });

  let sent = 0;
  let failed = 0;

  for (const [memberId, memberTasks] of Object.entries(tasksByMember)) {
    const member = memberMap[memberId];
    if (!member || !member.email) {
      console.log(`⚠️  No email for member ${memberId}, skipping`);
      continue;
    }

    const taskRows = memberTasks.map(t => `
      <tr>
        <td style="padding:10px 14px;border-bottom:1px solid #eee;font-weight:600">${t.title}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center">${priorityLabel(t.prio)}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center">
          <span style="background:#e8eef8;color:#2B5BAE;padding:3px 10px;border-radius:12px;font-size:12px">${t.status}</span>
        </td>
        <td style="padding:10px 14px;border-bottom:1px solid #eee;color:#e07b10;font-weight:600">${formatDate(t.due)}</td>
      </tr>
    `).join('');

    const html = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:'Sarabun','IBM Plex Sans Thai',sans-serif;background:#f2f4f8;margin:0;padding:20px">
  <div style="max-width:600px;margin:0 auto">
    <!-- Header -->
    <div style="background:linear-gradient(135deg,#1a3d7a,#2B5BAE);padding:28px 32px;border-radius:12px 12px 0 0;text-align:center">
      <div style="font-size:22px;font-weight:700;color:#fff;letter-spacing:.08em">AKARA RESOURCES</div>
      <div style="font-size:11px;color:rgba(255,255,255,.7);letter-spacing:.14em;margin-top:4px">HUMAN RESOURCES TASK MANAGEMENT</div>
      <div style="width:36px;height:3px;background:#F5C02B;border-radius:2px;margin:12px auto 0"></div>
    </div>

    <!-- Body -->
    <div style="background:#fff;padding:28px 32px">
      <div style="font-size:18px;font-weight:600;color:#1a1e2e;margin-bottom:6px">
        ⏰ แจ้งเตือน Task ครบกำหนดพรุ่งนี้
      </div>
      <div style="font-size:14px;color:#6b7899;margin-bottom:20px">
        สวัสดีคุณ <strong style="color:#1a1e2e">${member.name}</strong> 
        มี ${memberTasks.length} งานที่จะครบกำหนดในวันพรุ่งนี้ 
        <strong style="color:#e07b10">${formatDate(tomorrowStr)}</strong>
      </div>

      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:#f2f4f8">
            <th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">ชื่องาน</th>
            <th style="padding:10px 14px;text-align:center;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">ความสำคัญ</th>
            <th style="padding:10px 14px;text-align:center;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">สถานะ</th>
            <th style="padding:10px 14px;text-align:center;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#6b7899">กำหนดส่ง</th>
          </tr>
        </thead>
        <tbody>${taskRows}</tbody>
      </table>

      <div style="margin-top:24px;padding:16px;background:#fff8e0;border-radius:8px;border-left:4px solid #F5C02B;font-size:13px;color:#6b7899">
        💡 กรุณาอัปเดตสถานะงานในระบบให้ครบถ้วน หรือติดต่อผู้จัดการหากต้องการขยายเวลา
      </div>

      <div style="margin-top:20px;text-align:center">
        <a href="https://hrakararesources.github.io/akara-hr/" 
           style="display:inline-block;padding:12px 28px;background:#2B5BAE;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px">
          เปิดระบบ HR Task Board →
        </a>
      </div>
    </div>

    <!-- Footer -->
    <div style="background:#f2f4f8;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center;font-size:11px;color:#6b7899">
      อีเมลนี้ส่งโดยอัตโนมัติจากระบบ Akara Resources HR Task Management<br>
      กรุณาอย่าตอบกลับอีเมลนี้
    </div>
  </div>
</body>
</html>`;

    try {
      await transporter.sendMail({
        from: `"Akara HR System" <${process.env.SMTP_FROM}>`,
        to: member.email,
        subject: `⏰ แจ้งเตือน: มี ${memberTasks.length} งานครบกำหนดพรุ่งนี้ — ${formatDate(tomorrowStr)}`,
        html,
      });
      console.log(`✅ Sent to ${member.name} <${member.email}> (${memberTasks.length} task(s))`);
      sent++;
    } catch (e) {
      console.error(`❌ Failed to send to ${member.email}:`, e.message);
      failed++;
    }
  }

  console.log(`\n📊 Summary: ${sent} sent, ${failed} failed`);
}

main().catch(err => {
  console.error('❌ Fatal error:', err);
  process.exit(1);
});
