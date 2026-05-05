# Weekly Task Master 🗓️

> Automatisation complète de la planification hebdomadaire via Google Apps Script — zéro app tierce, 100% Google Workspace.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-JavaScript-4285F4?style=flat&logo=google)
![Status](https://img.shields.io/badge/Status-Fonctionnel-10b981?style=flat)
![Trigger](https://img.shields.io/badge/Triggers-3%20actifs-8b5cf6?style=flat)

---

## 🎯 Le Problème

Planifier sa semaine manuellement est chronophage et on abandonne rapidement. Les apps de productivité existantes sont trop complexes ou déconnectées des outils du quotidien. Pourtant, Google Sheets, Agenda et Gmail sont déjà utilisés chaque jour.

---

## 💡 La Solution

Un système entièrement automatisé, intégré à Google Workspace, qui :
- **Planifie la semaine en 1 clic** depuis un pop-up dans Google Sheets
- **Crée les événements** dans Google Agenda automatiquement
- **Vérifie et marque** chaque tâche ✅ / ❌ chaque soir à 23h
- **Envoie un email récap** chaque dimanche à 20h (XP, rang, stats par catégorie)
- **Envoie un briefing matinal** chaque jour à 8h (hier / aujourd'hui / conseil)
- **Archive** toutes les tâches complétées dans une feuille Historic

---

## 🏗️ Architecture

```
Task Library (Sheets)
       ↓
openPlanningPopup()  ←  Menu Google Sheets
       ↓
savePlanningFromPopup()
  ├── Weekly Plan (Sheets)
  └── Google Agenda (événements)
       ↓
checkCompletedTasks()  ←  Trigger 23h / jour
  ├── Marque ✅ / ❌ dans Weekly Plan
  └── archiveToHistoric() → Historic (Sheets)
       ↓
sendWeeklyRecap()  ←  Trigger dimanche 20h
  └── Email HTML (XP · Rang · Stats · Bouton)
       ↓
sendDailyMorningEmail()  ←  Trigger 8h / jour
  └── Email briefing (hier · aujourd'hui · conseil)
```

---

## ⚙️ Fonctions principales

| Fonction | Rôle | Déclenchement |
|----------|------|---------------|
| `openPlanningPopup()` | Génère le formulaire HTML depuis la Task Library | Manuel (menu) |
| `savePlanningFromPopup(tasks)` | Écrit dans Weekly Plan + crée les événements Agenda | Après validation pop-up |
| `checkCompletedTasks()` | Compare les événements Agenda → marque ✅/❌ | Trigger 23h |
| `archiveToHistoric()` | Archive les tâches ✅ dans Historic (anti-doublon) | Après checkCompleted |
| `sendWeeklyRecap()` | Calcule XP, génère le HTML, envoie l'email récap | Trigger dim. 20h |
| `sendDailyMorningEmail()` | Briefing quotidien avec conseil motivant | Trigger 8h |
| `setupTriggers()` | Configure les 3 triggers automatiques | 1 seule fois |

---

## 🗂️ Structure Google Sheets

| Onglet | Colonnes | Rôle |
|--------|----------|------|
| **Task Library** | Nom · Durée · Catégorie · Emoji | Source de vérité des tâches |
| **Weekly Plan** | Tâche · Date · Heure · Durée · Catégorie · Emoji · ✅/❌ | Planning actif de la semaine |
| **Historic** | Tâche · Catégorie · Date · Durée · Semaine · Mois · Année | Archive filtrables |

---

## 🧠 Concepts techniques utilisés

- **Time-driven Triggers** — automatisation sans intervention
- **HtmlService** — pop-up HTML dynamique injecté dans Sheets
- **google.script.run** — communication asynchrone HTML ↔ serveur
- **APIs Google Workspace** — CalendarApp, GmailApp, SpreadsheetApp
- **Web App déployée** — bouton dans l'email qui relance le planning
- **Déduplication via Set()** — évite les doublons dans Historic
- **Parsing robuste** — gestion des formats de durée natifs de Sheets

---

## 🚀 Installation

```bash
# 1. Crée un nouveau Google Sheet
# 2. Extensions → Apps Script → colle le contenu de Code.gs
# 3. Configure les constantes
```

```javascript
const SHEET_ID    = "TON_SHEET_ID";
const CALENDAR_ID = "primary";
const WEB_APP_URL = "TON_URL_WEB_APP";
```

```bash
# 4. Lance setupTriggers() une seule fois
# 5. Déploie comme Web App → copie l'URL dans WEB_APP_URL
```

---

## 🛣️ Roadmap

- [ ] Gamification avancée (streaks, badges, dashboard XP visuel)
- [ ] Historic analysable (graphiques automatiques, heatmap)
- [ ] Conseils IA personnalisés via Gemini API
- [ ] Alertes intelligentes si une catégorie chute sous 50%
- [ ] Bouton "Prêt à relever !" amélioré (Web App complète)

---

## 👤 Auteur

**Alec DIAMIDIA** — 2026  
Projet réalisé dans le cadre d'un cours sur Google Apps Script
