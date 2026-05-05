// ============================================
// ⚙️ CONFIGURATION
// ============================================

const SHEET_ID    = "TON_SHEET_ID_ICI";
const CALENDAR_ID = "primary";
const WEB_APP_URL = "TON_WEB_APP_URL_ICI";

// ============================================
// 🚀 MENU PERSONNALISÉ
// ============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Weekly Task Master")
    .addItem("Planifier ma semaine", "openPlanningPopup")
    .addSeparator()
    .addItem("Verifier les taches", "checkCompletedTasks")
    .addItem("Envoyer le recap hebdo", "sendWeeklyRecap")
    .addItem("Envoyer le briefing matin", "sendDailyMorningEmail")
    .addSeparator()
    .addItem("Archiver dans Historic", "archiveToHistoric")
    .addItem("Configurer les triggers", "setupTriggers")
    .addToUi();
}

// ============================================
// 🌐 WEB APP — Planification depuis l'email
// ============================================

function doGet(e) {
  return HtmlService.createHtmlOutput(getPlanningHTML())
    .setTitle("Weekly Task Master — Planification")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// 🎨 POP-UP INTERFACE DE PLANNING
// ============================================

function openPlanningPopup() {
  const html = HtmlService.createHtmlOutput(getPlanningHTML())
    .setWidth(820)
    .setHeight(620)
    .setTitle("Planifie ta Semaine");
  SpreadsheetApp.getUi().showModalDialog(html, "Weekly Task Master");
}

function getPlanningHTML() {
  const sheet       = SpreadsheetApp.openById(SHEET_ID);
  const taskLibrary = sheet.getSheetByName("Task Library");

  if (!taskLibrary) return "<p>Onglet 'Task Library' introuvable !</p>";

  const tasks = taskLibrary.getDataRange().getValues();
  let taskOptions = "";

  for (let i = 1; i < tasks.length; i++) {
    const nom       = tasks[i][0] || "";
    const duree     = tasks[i][1] || 0;
    const categorie = tasks[i][2] || "";
    const emoji     = tasks[i][3] || "";
    if (!nom) continue;

    taskOptions += `
      <div class="task-item">
        <input type="checkbox" id="task_${i}" value="${nom}|${duree}|${categorie}|${emoji}">
        <label for="task_${i}">
          <span class="emoji">${emoji}</span>
          <span class="name">${nom}</span>
          <span class="details">(${duree} min - ${categorie})</span>
        </label>
        <div class="schedule" id="schedule_${i}" style="display:none;">
          <label>Jours :</label>
          <div class="days">
            <label><input type="checkbox" value="lun"> Lun</label>
            <label><input type="checkbox" value="mar"> Mar</label>
            <label><input type="checkbox" value="mer"> Mer</label>
            <label><input type="checkbox" value="jeu"> Jeu</label>
            <label><input type="checkbox" value="ven"> Ven</label>
            <label><input type="checkbox" value="sam"> Sam</label>
            <label><input type="checkbox" value="dim"> Dim</label>
          </div>
          <label>Heure :</label>
          <input type="time" class="time-input" value="09:00">
        </div>
      </div>`;
  }

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <base target="_top">
      <style>
        * { box-sizing: border-box; }
        body { font-family: Arial, sans-serif; padding: 20px; background: linear-gradient(135deg, #667eea, #764ba2); margin: 0; }
        .container { background: white; border-radius: 10px; padding: 20px; max-height: 560px; overflow-y: auto; }
        h2 { color: #667eea; margin-top: 0; }
        h3 { color: #374151; font-size: 15px; margin: 0 0 12px; }
        .task-item { border: 2px solid #e0e0e0; border-radius: 8px; padding: 15px; margin-bottom: 12px; transition: border-color 0.3s; }
        .task-item:hover { border-color: #667eea; }
        .task-item > label { cursor: pointer; display: flex; align-items: center; gap: 10px; }
        .emoji { font-size: 22px; }
        .name { font-weight: bold; font-size: 15px; }
        .details { color: #666; font-size: 12px; }
        .schedule { margin-top: 12px; padding: 12px; background: #f5f5f5; border-radius: 5px; }
        .days { display: flex; gap: 8px; margin: 8px 0; flex-wrap: wrap; }
        .days label { padding: 4px 10px; background: white; border: 1px solid #ddd; border-radius: 5px; cursor: pointer; font-size: 13px; }
        .time-input { padding: 7px; border: 1px solid #ddd; border-radius: 5px; width: 120px; }
        .new-task-section { border: 2px dashed #667eea; border-radius: 10px; padding: 16px; margin-top: 20px; background: #f5f3ff; }
        .new-task-grid { display: grid; grid-template-columns: 2fr 1fr 1fr 0.5fr; gap: 8px; margin-bottom: 10px; }
        .new-task-grid input, .new-task-grid select { padding: 8px; border: 1px solid #ddd; border-radius: 6px; font-size: 13px; width: 100%; }
        .toggle-row { display: flex; gap: 16px; margin-bottom: 12px; align-items: center; flex-wrap: wrap; }
        .toggle-row label { display: flex; align-items: center; gap: 6px; cursor: pointer; font-size: 13px; font-weight: 600; }
        .tag-permanent { background: #ede9fe; color: #7c3aed; padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 700; }
        .tag-ponctuel  { background: #e0f2fe; color: #0369a1; padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 700; }
        .toggle-hint { font-size: 11px; font-style: italic; }
        .custom-schedule { background: #ede9fe; border-radius: 6px; padding: 10px 12px; margin-top: 8px; }
        .btn-add { background: #667eea; color: white; padding: 9px 20px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; width: 100%; margin-top: 8px; }
        #custom-tasks-list .task-item { border-color: #667eea; background: #faf5ff; }
        .buttons { margin-top: 20px; text-align: center; }
        .btn-save   { padding: 12px 30px; font-size: 15px; border: none; border-radius: 6px; cursor: pointer; margin: 0 8px; background: #10b981; color: white; }
        .btn-cancel { padding: 12px 30px; font-size: 15px; border: none; border-radius: 6px; cursor: pointer; margin: 0 8px; background: #6b7280; color: white; }
        .loading { display: none; text-align: center; padding: 16px; color: #667eea; font-weight: 600; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>Planifie ta Semaine</h2>
        <p style="color:#6b7280;font-size:13px;margin-top:-10px;margin-bottom:16px;">Coche les taches de ta bibliotheque, selectionne les jours et heures.</p>

        <div id="tasks">${taskOptions}</div>

        <div class="new-task-section">
          <h3>+ Ajouter une tache personnalisee</h3>
          <div class="toggle-row">
            <span style="font-size:12px;color:#6b7280;font-weight:600;">Type :</span>
            <label>
              <input type="radio" name="task-type" value="permanent" checked onchange="updateToggleStyle()">
              <span class="tag-permanent">Permanent</span>
            </label>
            <label>
              <input type="radio" name="task-type" value="ponctuel" onchange="updateToggleStyle()">
              <span class="tag-ponctuel">Ponctuel</span>
            </label>
            <span id="toggle-hint" class="toggle-hint" style="color:#7c3aed;">Sera sauvegardee dans la Task Library</span>
          </div>
          <div class="new-task-grid">
            <input type="text"   id="custom-name"     placeholder="Nom de la tache" />
            <input type="number" id="custom-duration"  placeholder="Duree (min)" value="60" min="5" />
            <select id="custom-category">
              <option value="Sport">Sport</option>
              <option value="Formation">Formation</option>
              <option value="Pro">Pro</option>
              <option value="Travail">Travail</option>
              <option value="Perso">Perso</option>
            </select>
            <input type="text" id="custom-emoji" placeholder="🎯" maxlength="2" style="text-align:center;font-size:18px;" />
          </div>
          <div class="custom-schedule">
            <div style="font-size:12px;font-weight:700;color:#374151;margin-bottom:6px;">Jours :</div>
            <div class="days" id="custom-days">
              <label><input type="checkbox" value="lun"> Lun</label>
              <label><input type="checkbox" value="mar"> Mar</label>
              <label><input type="checkbox" value="mer"> Mer</label>
              <label><input type="checkbox" value="jeu"> Jeu</label>
              <label><input type="checkbox" value="ven"> Ven</label>
              <label><input type="checkbox" value="sam"> Sam</label>
              <label><input type="checkbox" value="dim"> Dim</label>
            </div>
            <div style="margin-top:8px;">
              <span style="font-size:12px;font-weight:700;color:#374151;">Heure : </span>
              <input type="time" id="custom-time" value="09:00">
            </div>
          </div>
          <button class="btn-add" onclick="addCustomTask()">Ajouter cette tache</button>
        </div>

        <div id="custom-tasks-list"></div>

        <div class="buttons">
          <button class="btn-save"   onclick="savePlanning()">Creer les Evenements</button>
          <button class="btn-cancel" onclick="google.script.host.close()">Annuler</button>
        </div>
        <div class="loading" id="loading">Creation en cours...</div>
      </div>

      <script>
        var customTasks = [];

        function updateToggleStyle() {
          var type = document.querySelector('input[name="task-type"]:checked').value;
          var hint = document.getElementById('toggle-hint');
          if (type === 'permanent') {
            hint.style.color = '#7c3aed';
            hint.textContent = 'Sera sauvegardee dans la Task Library';
          } else {
            hint.style.color = '#0369a1';
            hint.textContent = 'Cette semaine uniquement — non sauvegardee';
          }
        }

        function addCustomTask() {
          var nom   = document.getElementById('custom-name').value.trim();
          var duree = parseInt(document.getElementById('custom-duration').value) || 60;
          var cat   = document.getElementById('custom-category').value;
          var emoji = document.getElementById('custom-emoji').value || '';
          var days  = Array.from(document.querySelectorAll('#custom-days input:checked')).map(function(d){ return d.value; });
          var time  = document.getElementById('custom-time').value;
          var type  = document.querySelector('input[name="task-type"]:checked').value;

          if (!nom)            { alert('Entre un nom pour la tache.');   return; }
          if (days.length===0) { alert('Selectionne au moins un jour.'); return; }

          customTasks.push({ nom: nom, duree: duree, categorie: cat, emoji: emoji, days: days, time: time, isPermanent: type === 'permanent' });

          var tag = type === 'permanent'
            ? '<span class="tag-permanent">Permanent</span>'
            : '<span class="tag-ponctuel">Ponctuel</span>';
          document.getElementById('custom-tasks-list').innerHTML +=
            '<div class="task-item" style="margin-top:10px;">'
            + '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">'
            + '<span style="font-size:20px;">' + emoji + '</span>'
            + '<span style="font-weight:700;">' + nom + '</span>'
            + '<span style="font-size:11px;color:#6b7280;">(' + duree + ' min - ' + cat + ')</span>'
            + tag
            + '<span style="font-size:11px;color:#6b7280;">' + days.join(', ') + ' a ' + time + '</span>'
            + '</div></div>';

          document.getElementById('custom-name').value     = '';
          document.getElementById('custom-duration').value = 60;
          document.getElementById('custom-emoji').value    = '';
          document.querySelectorAll('#custom-days input').forEach(function(cb){ cb.checked = false; });
        }

        document.querySelectorAll('#tasks .task-item input[type="checkbox"]').forEach(function(cb) {
          cb.addEventListener('change', function() {
            this.closest('.task-item').querySelector('.schedule').style.display = this.checked ? 'block' : 'none';
          });
        });

        function savePlanning() {
          var tasks = [];
          document.querySelectorAll('#tasks .task-item').forEach(function(item) {
            var cb = item.querySelector('input[type="checkbox"]');
            if (!cb || !cb.checked) return;
            var parts = cb.value.split('|');
            var days  = Array.from(item.querySelectorAll('.days input:checked')).map(function(d){ return d.value; });
            var time  = item.querySelector('.time-input').value;
            if (days.length > 0) {
              tasks.push({ nom: parts[0], duree: parseInt(parts[1]), categorie: parts[2], emoji: parts[3], days: days, time: time, isPermanent: false });
            }
          });
          customTasks.forEach(function(t){ tasks.push(t); });
          if (tasks.length === 0) { alert('Aucune tache selectionnee !'); return; }
          document.getElementById('loading').style.display = 'block';
          google.script.run
            .withSuccessHandler(function(r) {
              var msg = r.created + ' evenements crees !';
              if (r.saved > 0) msg += ' ' + r.saved + ' tache(s) ajoutee(s) a la Task Library.';
              alert(msg);
              google.script.host.close();
            })
            .withFailureHandler(function(e) {
              alert('Erreur: ' + e);
              document.getElementById('loading').style.display = 'none';
            })
            .savePlanningFromPopup(tasks);
        }
      </script>
    </body>
    </html>`;
}

function savePlanningFromPopup(tasks) {
  const sheet     = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = sheet.getSheetByName("Weekly Plan");
  const libSheet  = sheet.getSheetByName("Task Library");
  const calendar  = CalendarApp.getCalendarById(CALENDAR_ID);

  const lastRow = planSheet.getLastRow();
  if (lastRow > 1) planSheet.getRange(2, 1, lastRow - 1, 7).clearContent();

  let created = 0, saved = 0;
  const today  = new Date();
  const dayMap = { "dim": 0, "lun": 1, "mar": 2, "mer": 3, "jeu": 4, "ven": 5, "sam": 6 };

  const libData  = libSheet.getDataRange().getValues();
  const existing = new Set(libData.map(r => String(r[0]).toLowerCase()));

  tasks.forEach(task => {
    if (task.isPermanent && !existing.has(task.nom.toLowerCase())) {
      libSheet.appendRow([task.nom, task.duree, task.categorie, task.emoji || "📌"]);
      existing.add(task.nom.toLowerCase());
      saved++;
    }
    task.days.forEach(day => {
      let daysUntil = dayMap[day] - today.getDay();
      if (daysUntil <= 0) daysUntil += 7;
      const date = new Date(today);
      date.setDate(today.getDate() + daysUntil);
      planSheet.appendRow([task.nom, date, task.time, task.duree, task.categorie, task.emoji, ""]);
      const [h, m] = task.time.split(":");
      const start  = new Date(date);
      start.setHours(parseInt(h), parseInt(m), 0, 0);
      const end = new Date(start.getTime() + task.duree * 60000);
      calendar.createEvent(`${task.emoji || ""} ${task.nom}`, start, end, { description: `Categorie: ${task.categorie}` });
      created++;
    });
  });

  return { created, saved };
}

// ============================================
// 📅 CRÉER LES ÉVÉNEMENTS DEPUIS LE PLAN
// ============================================

function createWeeklyEvents() {
  const sheet     = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = sheet.getSheetByName("Weekly Plan");
  const calendar  = CalendarApp.getCalendarById(CALENDAR_ID);
  const data      = planSheet.getDataRange().getValues();
  let created = 0, errors = 0;

  for (let i = 1; i < data.length; i++) {
    const [taskName, date, time, duration, category, emoji, check] = data[i];
    if (!taskName || check === "✅") continue;
    try {
      const eventDate = new Date(date);
      let h, m;
      if (typeof time === "string" && time.includes(":")) { [h, m] = time.split(":"); }
      else if (time instanceof Date) { h = time.getHours(); m = time.getMinutes(); }
      else continue;
      eventDate.setHours(parseInt(h), parseInt(m), 0, 0);
      const endDate = new Date(eventDate.getTime() + parseDuration(duration) * 60000);
      calendar.createEvent(`${emoji || ""} ${taskName}`, eventDate, endDate, { description: `Categorie: ${category}` });
      created++;
    } catch (e) {
      Logger.log("Erreur " + taskName + ": " + e.message);
      errors++;
    }
  }
  Logger.log(created + " crees, " + errors + " erreurs");
}

// ============================================
// ✅ VÉRIFIER LES TÂCHES COMPLÉTÉES
// ============================================

function checkCompletedTasks() {
  const sheet     = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = sheet.getSheetByName("Weekly Plan");
  const calendar  = CalendarApp.getCalendarById(CALENDAR_ID);
  const data      = planSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const [taskName, date, time] = data[i];
    if (!taskName) continue;
    const eventDate = new Date(date);
    let h, m;
    if (typeof time === "string" && time.includes(":")) { [h, m] = time.split(":"); h = parseInt(h); m = parseInt(m); }
    else if (time instanceof Date) { h = time.getHours(); m = time.getMinutes(); }
    else continue;
    eventDate.setHours(h, m, 0, 0);
    const endDate = new Date(eventDate.getTime() + 60 * 60000);
    const events  = calendar.getEvents(eventDate, endDate);
    const found   = events.find(e => e.getTitle().includes(taskName));
    planSheet.getRange(i + 1, 7).setValue(found ? "✅" : "❌");
  }

  archiveToHistoric();
  Logger.log("Verification terminee !");
}

// ============================================
// 📧 EMAIL RÉCAP HEBDOMADAIRE
// ============================================

function sendWeeklyRecap() {
  const sheet     = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = sheet.getSheetByName("Weekly Plan");
  const data      = planSheet.getDataRange().getValues();

  const XP_MULTIPLIERS = { "Sport": 1.5, "Formation": 1.3, "Pro": 1.2, "Travail": 1.1, "Perso": 1.0 };
  const CAT_COLORS = {
    "Sport":     { text: "#059669", bg: "#f0fdf4", bar: "#10b981" },
    "Formation": { text: "#2563eb", bg: "#eff6ff", bar: "#3b82f6" },
    "Pro":       { text: "#7c3aed", bg: "#f5f3ff", bar: "#8b5cf6" },
    "Travail":   { text: "#d97706", bg: "#fffbeb", bar: "#f59e0b" },
    "Perso":     { text: "#db2777", bg: "#fdf2f8", bar: "#ec4899" },
  };

  let totalXP = 0;
  const taskMap = {}, catStats = {};

  for (let i = 1; i < data.length; i++) {
    const [taskName, date, time, duration, category, emoji, check] = data[i];
    if (!taskName) continue;
    const cat = category || "Autre", dureeMin = parseDuration(duration);
    if (!taskMap[taskName])  taskMap[taskName]  = { cat, ok: 0, ko: 0 };
    if (!catStats[cat]) catStats[cat] = { ok: 0, ko: 0, min: 0 };
    if (check === "✅") { taskMap[taskName].ok++; catStats[cat].ok++; catStats[cat].min += dureeMin; totalXP += Math.round(dureeMin * (XP_MULTIPLIERS[cat] || 1.0)); }
    else if (check === "❌") { taskMap[taskName].ko++; catStats[cat].ko++; }
  }

  let totalOk = 0, totalKo = 0;
  Object.values(taskMap).forEach(t => { totalOk += t.ok; totalKo += t.ko; });
  const total = totalOk + totalKo;
  const rate  = total > 0 ? Math.round((totalOk / total) * 100) : 0;
  let rang  = totalXP < 500 ? "Debutant" : totalXP < 2000 ? "Regulier" : totalXP < 5000 ? "Discipline" : "Elite";
  let badge = rate === 100 ? "PERFECT WEEK !" : rate >= 80 ? "On Fire !" : rate >= 60 ? "Bien joue !" : "Continue comme ca !";

  let tasksHTML = "";
  Object.entries(taskMap).sort(([,a],[,b]) => (b.ok/(b.ok+b.ko||1)) - (a.ok/(a.ok+a.ko||1))).forEach(([name, t]) => {
    const taskTotal = t.ok + t.ko, dots = "&#9679;".repeat(t.ok) + "&#9675;".repeat(t.ko);
    const isAllDone = t.ko === 0, isAllMiss = t.ok === 0;
    const border = isAllDone ? "#bbf7d0" : isAllMiss ? "#fecaca" : "#e2e8f0";
    const bg = isAllDone ? "#f0fdf4" : isAllMiss ? "#fef2f2" : "#f9fafb";
    const scoreColor = isAllDone ? "#10b981" : isAllMiss ? "#ef4444" : "#f59e0b";
    const c = CAT_COLORS[t.cat] || { text: "#6b7280" };
    tasksHTML += `<table width="100%" cellpadding="0" cellspacing="0" style="background:${bg};border:1px solid ${border};border-radius:8px;margin-bottom:8px;"><tr><td style="padding:12px 14px;"><div style="font-weight:700;color:#1f2937;font-size:14px;">${name}</div><div style="font-size:11px;color:${c.text};margin-top:2px;">${t.cat}</div></td><td style="padding:12px 14px;text-align:right;white-space:nowrap;"><div style="font-weight:900;color:${scoreColor};font-size:15px;">${t.ok}/${taskTotal}</div><div style="margin-top:4px;letter-spacing:3px;font-size:13px;">${dots}</div></td></tr></table>`;
  });

  let catHTML = "";
  Object.keys(catStats).sort().forEach(cat => {
    const s = catStats[cat], c = CAT_COLORS[cat] || { text: "#6b7280", bg: "#f9fafb", bar: "#9ca3af" };
    const catTot = s.ok + s.ko, pct = catTot > 0 ? Math.round((s.ok / catTot) * 100) : 0;
    catHTML += `<table width="100%" cellpadding="0" cellspacing="0" style="background:${c.bg};border-radius:10px;margin-bottom:8px;"><tr><td style="padding:11px 14px;"><div style="font-weight:700;color:${c.text};font-size:13px;">${cat}</div><div style="font-size:11px;color:#9ca3af;margin-top:2px;">${s.ok} faites &middot; ${s.ko} manquees</div></td><td style="padding:11px 14px;text-align:right;white-space:nowrap;"><div style="font-weight:900;color:${c.bar};font-size:18px;">${pct}%</div><div style="font-size:11px;color:#9ca3af;">${(s.min/60).toFixed(1)}h</div></td></tr></table>`;
  });

  const emailBody = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:20px;background:#eef2f7;font-family:Arial,sans-serif;"><table width="100%" cellpadding="0" cellspacing="0" style="max-width:540px;margin:0 auto;"><tr><td><table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#1e1b4b,#4c1d95,#7c3aed);border-radius:16px 16px 0 0;"><tr><td style="padding:30px 24px;text-align:center;"><div style="font-size:10px;letter-spacing:3px;text-transform:uppercase;color:rgba(255,255,255,0.7);margin-bottom:8px;">Rapport Hebdomadaire</div><div style="font-size:24px;font-weight:900;color:white;margin-bottom:4px;">Weekly Task Master</div></td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:white;"><tr><td style="padding:32px 24px 20px;text-align:center;"><div style="display:inline-block;width:110px;height:110px;border-radius:55px;background:linear-gradient(135deg,#667eea,#764ba2);line-height:110px;font-size:34px;font-weight:900;color:white;margin-bottom:12px;">${rate}%</div><div style="font-size:13px;color:#6b7280;margin-bottom:14px;">Taux de completion</div><div style="display:inline-block;background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:9px 28px;border-radius:30px;font-size:13px;font-weight:700;">${badge}</div></td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:white;"><tr><td style="padding:0 20px 20px;"><table width="100%" cellpadding="0" cellspacing="0"><tr><td width="48%" style="background:linear-gradient(135deg,#fffbeb,#fef3c7);border:2px solid #f59e0b;border-radius:12px;padding:16px 12px;text-align:center;"><div style="font-size:10px;color:#92400e;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">XP Gagnes</div><div style="font-size:32px;font-weight:900;color:#d97706;">+${totalXP}</div></td><td width="4%"></td><td width="48%" style="background:linear-gradient(135deg,#eef2ff,#e0e7ff);border:2px solid #818cf8;border-radius:12px;padding:16px 12px;text-align:center;"><div style="font-size:10px;color:#4338ca;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Ton Rang</div><div style="font-size:20px;font-weight:900;color:#4f46e5;">${rang}</div></td></tr></table></td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:white;"><tr><td style="padding:0 20px 20px;"><div style="font-size:10px;color:#9ca3af;text-transform:uppercase;letter-spacing:2px;font-weight:700;margin-bottom:10px;">Par Categorie</div>${catHTML}</td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:#f9fafb;"><tr><td style="padding:18px 20px 24px;"><div style="font-size:10px;color:#374151;font-weight:800;text-transform:uppercase;letter-spacing:1.5px;margin-bottom:12px;">Taches &#9679; faites &#9675; manquees</div>${tasksHTML}</td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:white;"><tr><td style="padding:24px;text-align:center;"><div style="font-size:13px;color:#6b7280;margin-bottom:14px;">Pret pour la semaine prochaine ?</div><a href="${WEB_APP_URL}" target="_blank" style="display:inline-block;background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:14px 32px;border-radius:30px;font-size:14px;font-weight:800;text-decoration:none;">Pret a relever ! &rarr;</a></td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#1e1b4b,#4c1d95,#7c3aed);border-radius:0 0 16px 16px;"><tr><td style="padding:24px 20px;text-align:center;"><div style="color:white;font-size:16px;font-weight:800;margin-bottom:4px;">Continue comme ca !</div><div style="color:rgba(255,255,255,0.5);font-size:11px;">Weekly Task Master &middot; Rapport automatique</div></td></tr></table></td></tr></table></body></html>`;

  GmailApp.sendEmail(Session.getActiveUser().getEmail(), `Recap Semaine - ${rate}% - ${rang} - +${totalXP} XP`, "", { htmlBody: emailBody, name: "Weekly Task Master" });
  Logger.log("Email envoye ! " + rate + "% | +" + totalXP + " XP | " + rang);
}

// ============================================
// 🌅 EMAIL MATINAL QUOTIDIEN
// ============================================

function sendDailyMorningEmail() {
  const sheet     = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = sheet.getSheetByName("Weekly Plan");
  const data      = planSheet.getDataRange().getValues();
  const today     = new Date(), yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const todayStr = today.toLocaleDateString("fr-FR"), yesterdayStr = yesterday.toLocaleDateString("fr-FR");
  const doneYesterday = [], plannedToday = [];

  for (let i = 1; i < data.length; i++) {
    const [taskName, date, time, duration, category, emoji, check] = data[i];
    if (!taskName) continue;
    const taskDate = new Date(date).toLocaleDateString("fr-FR");
    if (taskDate === yesterdayStr && check === "✅") doneYesterday.push({ nom: taskName, cat: category, duree: parseDuration(duration) });
    if (taskDate === todayStr) plannedToday.push({ nom: taskName, cat: category, heure: time, duree: parseDuration(duration) });
  }

  const tips = [
    "Chaque effort compte. Une brique par jour, et la maison se construit.",
    "La discipline, c'est faire ce qu'il faut meme quand tu n'en as pas envie.",
    "Progres > Perfection. Avance, meme lentement.",
    "Tes habitudes d'aujourd'hui dessinent ta vie de demain.",
    "Une seule tache bien faite vaut mieux que dix baclees.",
    "Le succes se construit dans les moments ou personne ne regarde.",
    "Commence. Le reste suit toujours.",
  ];
  const tip = tips[today.getDay() % tips.length];
  const CAT_COLORS = { "Sport": { text: "#059669", bg: "#f0fdf4" }, "Formation": { text: "#2563eb", bg: "#eff6ff" }, "Pro": { text: "#7c3aed", bg: "#f5f3ff" }, "Travail": { text: "#d97706", bg: "#fffbeb" }, "Perso": { text: "#db2777", bg: "#fdf2f8" } };

  let doneHTML = doneYesterday.length > 0
    ? doneYesterday.map(t => { const c = CAT_COLORS[t.cat] || { text: "#6b7280", bg: "#f9fafb" }; return `<table width="100%" cellpadding="0" cellspacing="0" style="background:${c.bg};border-radius:8px;margin-bottom:6px;"><tr><td style="padding:10px 14px;"><span style="font-weight:700;color:#1f2937;font-size:13px;">${t.nom}</span><span style="font-size:11px;color:${c.text};margin-left:8px;">${t.cat}</span></td><td style="padding:10px 14px;text-align:right;"><span style="font-weight:700;color:#10b981;font-size:12px;">${t.duree}min</span></td></tr></table>`; }).join("")
    : `<p style="color:#9ca3af;font-size:13px;font-style:italic;margin:4px 0;">Aucune tache completee hier.</p>`;

  let todayHTML = plannedToday.length > 0
    ? plannedToday.map(t => { const c = CAT_COLORS[t.cat] || { text: "#6b7280", bg: "#f9fafb" }; const heure = t.heure instanceof Date ? `${t.heure.getHours()}h${String(t.heure.getMinutes()).padStart(2,"0")}` : String(t.heure).substring(0, 5); return `<table width="100%" cellpadding="0" cellspacing="0" style="background:${c.bg};border-radius:8px;margin-bottom:6px;"><tr><td style="padding:10px 14px;"><span style="font-weight:700;color:#1f2937;font-size:13px;">${t.nom}</span><span style="font-size:11px;color:${c.text};margin-left:8px;">${t.cat}</span></td><td style="padding:10px 14px;text-align:right;"><span style="font-weight:600;color:#667eea;font-size:12px;">${heure} &middot; ${t.duree}min</span></td></tr></table>`; }).join("")
    : `<p style="color:#9ca3af;font-size:13px;font-style:italic;margin:4px 0;">Rien de prevu. Profites-en pour avancer !</p>`;

  const jours = ["Dimanche","Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi"];
  const emailBody = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:20px;background:#eef2f7;font-family:Arial,sans-serif;"><table width="100%" cellpadding="0" cellspacing="0" style="max-width:540px;margin:0 auto;"><tr><td><table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#0f172a,#1e3a8a,#2563eb);border-radius:16px 16px 0 0;"><tr><td style="padding:28px 24px;text-align:center;"><div style="font-size:10px;letter-spacing:3px;text-transform:uppercase;color:rgba(255,255,255,0.6);margin-bottom:8px;">Briefing du matin</div><div style="font-size:22px;font-weight:900;color:white;margin-bottom:4px;">${jours[today.getDay()]} ${todayStr}</div></td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#667eea,#764ba2);"><tr><td style="padding:16px 24px;text-align:center;"><div style="font-size:13px;color:white;font-style:italic;line-height:1.6;">&laquo; ${tip} &raquo;</div></td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:white;"><tr><td style="padding:20px 20px 8px;"><div style="font-size:10px;color:#9ca3af;text-transform:uppercase;letter-spacing:2px;font-weight:700;margin-bottom:10px;">Hier (${yesterdayStr})</div>${doneHTML}</td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:white;"><tr><td style="padding:16px 20px 24px;"><div style="font-size:10px;color:#9ca3af;text-transform:uppercase;letter-spacing:2px;font-weight:700;margin-bottom:10px;">Aujourd'hui (${todayStr})</div>${todayHTML}</td></tr></table><table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#0f172a,#1e3a8a);border-radius:0 0 16px 16px;"><tr><td style="padding:20px;text-align:center;"><div style="color:white;font-size:14px;font-weight:800;">Bonne journee !</div></td></tr></table></td></tr></table></body></html>`;

  GmailApp.sendEmail(Session.getActiveUser().getEmail(), `Briefing ${jours[today.getDay()]} ${todayStr} - ${plannedToday.length} tache(s) prevue(s)`, "", { htmlBody: emailBody, name: "Weekly Task Master" });
  Logger.log("Email matinal envoye !");
}

// ============================================
// 📂 HISTORIC — Archivage automatique
// ============================================

function createHistoricSheetIfNeeded() {
  const sheet = SpreadsheetApp.openById(SHEET_ID);
  let historic = sheet.getSheetByName("Historic");
  if (!historic) {
    historic = sheet.insertSheet("Historic");
    historic.getRange("A1:G1").setValues([["Tache","Categorie","Date","Duree (min)","Semaine","Mois","Annee"]]);
    historic.getRange("A1:G1").setFontWeight("bold").setBackground("#1e1b4b").setFontColor("white");
    historic.setFrozenRows(1);
    [150, 100, 100, 110, 80, 70, 70].forEach((w, i) => historic.setColumnWidth(i + 1, w));
  }
  return historic;
}

function archiveToHistoric() {
  const sheet     = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = sheet.getSheetByName("Weekly Plan");
  const historic  = createHistoricSheetIfNeeded();
  const planData  = planSheet.getDataRange().getValues();
  const histData  = historic.getDataRange().getValues();
  const existing  = new Set();
  for (let i = 1; i < histData.length; i++) existing.add(histData[i][0] + "_" + histData[i][2]);

  const newRows = [];
  for (let i = 1; i < planData.length; i++) {
    const [taskName, date, time, duration, category, emoji, check] = planData[i];
    if (!taskName || check !== "✅") continue;
    const d = new Date(date), dateStr = d.toLocaleDateString("fr-FR"), key = taskName + "_" + dateStr;
    if (existing.has(key)) continue;
    const startOfYear = new Date(d.getFullYear(), 0, 1);
    const weekNum = Math.ceil(((d - startOfYear) / 86400000 + startOfYear.getDay() + 1) / 7);
    newRows.push([taskName, category || "Autre", dateStr, parseDuration(duration), "S" + weekNum, d.getMonth() + 1, d.getFullYear()]);
  }

  if (newRows.length > 0) historic.getRange(historic.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  Logger.log(newRows.length + " taches archivees.");
}

// ============================================
// 🛠️ UTILITAIRES
// ============================================

function parseDuration(duration) {
  if (!duration && duration !== 0) return 0;
  if (duration instanceof Date) return duration.getHours() * 60 + duration.getMinutes();
  const n = Number(duration);
  if (!isNaN(n)) return n > 0 && n < 1 ? Math.round(n * 1440) : Math.round(n);
  const s = String(duration);
  if (s.includes(":")) { const parts = s.split(":").map(Number); return parts[0] * 60 + (parts[1] || 0); }
  return parseInt(s) || 0;
}

// ============================================
// ⏰ TRIGGERS AUTOMATIQUES (lancer 1 seule fois)
// ============================================

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Verification + archivage chaque soir a 23h
  ScriptApp.newTrigger("checkCompletedTasks")
    .timeBased().everyDays(1).atHour(23).create();

  // Recap hebdo chaque dimanche a 20h
  ScriptApp.newTrigger("sendWeeklyRecap")
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(20).create();

  // Email matinal chaque jour a 8h
  ScriptApp.newTrigger("sendDailyMorningEmail")
    .timeBased().everyDays(1).atHour(8).create();

  Logger.log("3 triggers crees : 8h matin + 23h verif + dimanche 20h recap");
}

// ============================================
// 🧪 TEST COMPLET
// ============================================

function testAll() {
  Logger.log("TEST COMPLET");
  try {
    Logger.log("1 - Creation des evenements...");
    createWeeklyEvents();
    Utilities.sleep(2000);
    Logger.log("2 - Verification des taches...");
    checkCompletedTasks();
    Utilities.sleep(2000);
    Logger.log("3 - Email recap...");
    sendWeeklyRecap();
    Utilities.sleep(1000);
    Logger.log("4 - Email matinal...");
    sendDailyMorningEmail();
    Logger.log("TEST TERMINE !");
  } catch (e) {
    Logger.log("ERREUR : " + e.message);
  }
}
