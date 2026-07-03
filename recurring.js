// recurring.js — สร้าง Task จากแม่แบบ "งานประจำทุกเดือน"
// รันโดย GitHub Actions ทุกวันเช้า (เช็ค last_period กันสร้างซ้ำ → สร้างครั้งเดียวต่อเดือน)
// การแจ้งเตือน/อีเมล: insert notifications แล้ว trigger (013) ยิง Edge Function ส่งเมลให้เอง

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;   // service_role

async function sb(path, opts = {}) {
  const res = await fetch(SUPABASE_URL + '/rest/v1/' + path, {
    method: opts.method || 'GET',
    headers: { apikey: SUPABASE_KEY, Authorization: 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json', Prefer: opts.prefer || '' },
    body: opts.body,
  });
  if (!res.ok) throw new Error(`${res.status}: ${await res.text()}`);
  const t = await res.text();
  return t ? JSON.parse(t) : null;
}

const uid = () => 'x' + Math.random().toString(36).slice(2) + Date.now().toString(36);
const pad = n => String(n).padStart(2, '0');

async function main() {
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth() + 1;
  const period = `${y}-${pad(m)}`;
  console.log('🔁 Recurring tasks — period', period);

  const templates = await sb('recurring_tasks?active=eq.true&select=*') || [];
  const due = templates.filter(t => t.last_period !== period);
  if (!due.length) { console.log('✅ Nothing to generate (all up to date).'); return; }

  let created = 0;
  for (const rt of due) {
    try {
      // ผู้รับผิดชอบ
      const asg = await sb(`recurring_task_assignees?recurring_id=eq.${rt.id}&select=member_id`) || [];
      let members = asg.map(a => a.member_id).filter(Boolean);
      if (!members.length && rt.created_by) members = [rt.created_by];
      if (!members.length) { console.log(`⚠️  ${rt.title}: no assignee, skip`); continue; }

      const taskId = uid();
      const dueDate = `${y}-${pad(m)}-${pad(rt.due_day || 1)}`;
      await sb('tasks', {
        method: 'POST', prefer: 'return=minimal',
        body: JSON.stringify({
          id: taskId, member_id: members[0], title: rt.title,
          description: rt.description || null, prio: rt.prio || 'ปานกลาง',
          status: 'รอดำเนินการ', due: dueDate, prog: 0,
          note: 'สร้างอัตโนมัติจากงานประจำเดือน', created_by: rt.created_by || members[0],
        }),
      });
      await sb('task_assignees', { method: 'POST', prefer: 'return=minimal',
        body: JSON.stringify(members.map(mid => ({ task_id: taskId, member_id: mid }))) });

      // แจ้งเตือน (insert → trigger ส่งอีเมลให้เอง)
      await sb('notifications', { method: 'POST', prefer: 'return=minimal',
        body: JSON.stringify(members.map(mid => ({
          member_id: mid, type: 'task_assigned', task_id: taskId,
          title: 'งานประจำเดือนใหม่', body: `${rt.title} · กำหนดส่ง ${dueDate}`,
          created_by: rt.created_by || mid,
        }))) });

      await sb(`recurring_tasks?id=eq.${rt.id}`, { method: 'PATCH', prefer: 'return=minimal',
        body: JSON.stringify({ last_period: period }) });

      console.log(`✅ Created: "${rt.title}" due ${dueDate} → ${members.length} assignee(s)`);
      created++;
    } catch (e) {
      console.error(`❌ ${rt.title}:`, e.message);
    }
  }
  console.log(`\n📊 Generated ${created} task(s) for ${period}`);
}

main().catch(e => { console.error('Fatal:', e); process.exit(1); });
