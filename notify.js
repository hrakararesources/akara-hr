// notify.js — Task Due Date Email Notification
// รันโดย GitHub Actions ทุกเช้า 08:00 น.

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const TENANT = process.env.GRAPH_TENANT_ID;
const CLIENT = process.env.GRAPH_CLIENT_ID;
const SECRET = process.env.GRAPH_CLIENT_SECRET;
const SENDER = process.env.GRAPH_SENDER;

// ── FETCH จาก Supabase ──
async function sbGet(path) {
  const res = await fetch(SUPABASE_URL + '/rest/v1/' + path, {
    headers: { apikey: SUPABASE_KEY, Authorization: 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json' }
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  const t = await res.text();
  return t ? JSON.parse(t) : null;
}

// ── Microsoft Graph (ส่งอีเมลผ่าน MS365 · client credentials) ──
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

  const todayStr = new Date().toISOString().slice(0,10);

  // งานครบกำหนดวันนี้ / พรุ่งนี้ / เกินกำหนด (ยังไม่เสร็จ)
  const tasksDueToday = await sbGet(
    `tasks?due=eq.${todayStr}&status=neq.เสร็จสิ้น&select=*`
  );
  const tasksDueTomorrow = await sbGet(
    `tasks?due=eq.${tomorrowStr}&status=neq.เสร็จสิ้น&select=*`
  );
  const tasksOverdue = await sbGet(
    `tasks?due=lt.${todayStr}&status=neq.เสร็จสิ้น&select=*`
  );

  const tasks = [...(tasksDueToday||[]), ...(tasksDueTomorrow||[]), ...(tasksOverdue||[])];

  if (!tasks.length) {
    console.log('✅ No tasks due today/tomorrow or overdue. Nothing to send.');
    return;
  }

  console.log(`📋 Found ${tasksDueToday?.length||0} due today, ${tasksDueTomorrow?.length||0} due tomorrow, ${tasksOverdue?.length||0} overdue`);

  // ดึง members ทั้งหมด
  const members = await sbGet('members?select=*');
  const memberMap = {};
  members.forEach(m => { memberMap[m.id] = m; });

  // ดึงผู้รับผิดชอบหลายคนต่องาน (junction) — ถ้าไม่มี/ยังไม่ได้สร้าง ใช้ member_id เดิม
  let assigneesByTask = {};
  try {
    const tas = await sbGet('task_assignees?select=task_id,member_id');
    (tas || []).forEach(r => {
      if (!assigneesByTask[r.task_id]) assigneesByTask[r.task_id] = [];
      assigneesByTask[r.task_id].push(r.member_id);
    });
  } catch (e) { assigneesByTask = {}; }

  // จัดกลุ่ม tasks ตาม member (ส่งถึงผู้รับผิดชอบทุกคนของงาน)
  const tasksByMember = {};
  tasks.forEach(t => {
    const ids = (assigneesByTask[t.id] && assigneesByTask[t.id].length) ? assigneesByTask[t.id] : [t.member_id];
    ids.forEach(mid => {
      if (!mid) return;
      if (!tasksByMember[mid]) tasksByMember[mid] = [];
      tasksByMember[mid].push(t);
    });
  });

  let sent = 0;
  let failed = 0;
  const token = await graphToken();

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
        ⏰ แจ้งเตือนงานใกล้ครบกำหนด / เกินกำหนด
      </div>
      <div style="font-size:14px;color:#6b7899;margin-bottom:20px">
        สวัสดีคุณ <strong style="color:#1a1e2e">${member.name}</strong>
        มี ${memberTasks.length} งานที่ต้องดูแล (ครบกำหนดวันนี้/พรุ่งนี้ หรือเกินกำหนด)
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
      await sendMail(token, member.email, `⏰ แจ้งเตือน Task HR — ${new Date().toLocaleDateString('th-TH',{day:'numeric',month:'long',year:'numeric'})}`, html);
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
