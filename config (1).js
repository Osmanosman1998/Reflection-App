// ============================================================
//  SETTINGS
//  Intern details, targets, rotations and extra questions all come
//  from the Excel workbook through Power Automate. Nothing about
//  interns is stored here.
// ============================================================

const CONFIG = {
  version: "1.4",

  // Paste the two Power Automate links between the quotes (see the set-up guide).
  // Flow 1, "Reflection app: get intern"
  profileUrl: "https://defaulte5b166d191cc45aaaaebd5b9deb798.a3.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/22/workflows/a277e7f47a4642caab35717d31a93b6b/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BzAb2gq2vb9MHq4H1sWwC8KGNTX3pqI9356ImOLTLt8",
  // Flow 2, "Reflection app: save reflection"
  submitUrl: "https://defaulte5b166d191cc45aaaaebd5b9deb798.a3.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/05/workflows/6efb98f83981408f8062bc12be9fd461/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=odH3o2jfAtnehmcn03yPgu1YYM0AU176i62sY6BhK2c",

  // How long to wait for Power Automate before saving on the phone instead (milliseconds)
  profileTimeoutMs: 30000,
  saveTimeoutMs: 45000,

  // Preparing for Adulthood areas: one target each, in this order.
  // "column" must match the heading on the Interns tab in Excel.
  areas: [
    { name: "Learning, Education & Employment", icon: "💼", column: "Learning, Education & Employment target" },
    { name: "Independence",                     icon: "🏠", column: "Independence target" },
    { name: "Health",                           icon: "❤️", column: "Health target" },
    { name: "Friendships & Relationships",      icon: "🤝", column: "Friendships & Relationships target" }
  ],

  // Half-term check-in: how many days before the end of each half term it appears.
  // The half-term dates come from the Settings tab in Excel.
  checkInDays: 7,

  // Phoenix independence key. Replace the labels when the key is available.
  scale: [
    { n: 1, label: "Full support",  says: "Someone did it with me" },
    { n: 2, label: "Lots of help",  says: "I needed lots of help" },
    { n: 3, label: "Some help",     says: "I needed some help" },
    { n: 4, label: "A little help", says: "I needed a little help" },
    { n: 5, label: "On my own",     says: "I did it on my own" }
  ],

  // Which version of the app each support level sees. Interns never see their level.
  versionForLevel: { 1: "supported", 2: "middle", 3: "independent" },

  // Word goals. Interns can carry on below the goal after one nudge.
  wordGoals: {
    middle:      { follow: 8, target: 8 },
    independent: { well: 40, hard: 40, next: 30, learned: 40, target: 20 }
  },

  // Word banks for the supported version (and the first step of the middle version)
  banks: {
    well: ["I was on time", "I finished my tasks", "I worked on my own", "I asked for help when I needed it",
           "I talked to my team", "I stayed calm", "I followed the instructions", "I checked my work",
           "I used my checklist", "I helped someone"],
    hard: ["It was busy", "It was noisy", "The plan changed", "I did not understand the instructions",
           "I felt tired", "I felt worried", "I did not know who to ask", "The task was new"],
    next: ["check my list", "ask if I am not sure", "take my break on time", "say hello to my team",
           "work a bit faster", "check my work", "use my headphones on breaks", "write down the instructions"]
  }
};
