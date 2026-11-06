// ============================================================
//  SETTINGS
//  Intern details, targets, rotations and extra questions all come
//  from the Excel workbook through Power Automate. Nothing about
//  interns is stored here.
// ============================================================

const CONFIG = {
  version: "2.0",

  // Paste the two Power Automate links between the quotes (see the set-up guide).
  // Flow 1, "Reflection app: get intern"
  profileUrl: "https://defaulte5b166d191cc45aaaaebd5b9deb798.a3.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/24/workflows/716a3a4fbaa343b99b7d7d53e3d7acb7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=xPU4ahTwZfsavvXqtkCMSwBcsCpC075-SHxtSZkHhho",
  // Flow 2, "Reflection app: save reflection"
  submitUrl: "https://defaulte5b166d191cc45aaaaebd5b9deb798.a3.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/03/workflows/1eaf625c62344164a8f4b3c6d00c81e3/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=8ususLYokw3PMdE6LUpt3B4d230r53EYhxosQ40GDpA",
  // Team check page: "Reflection app: team list" and "Reflection app: save check"
  teamUrl: "",
  saveCheckUrl: "",

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


  // Shown under every text box (Levels 2 and 3)
  writingTip: "Write like you are talking to your manager: full sentences, a capital letter at the start and a full stop at the end.",

  // Text slang the writing check picks up, with the professional word to use instead
  slang: { u: "you", ur: "your", r: "are", gonna: "going to", wanna: "want to", gotta: "have to", idk: "I don't know",
           cuz: "because", cos: "because", coz: "because", bc: "because", pls: "please", plz: "please", thx: "thank you",
           tbh: "to be honest", ya: "you", yeah: "yes", dunno: "don't know", lol: "a full sentence", omg: "a full sentence" },

  // Word banks for the supported version (and the first step of the middle version)
  banks: {
    well: ["I was on time", "I finished my tasks", "I worked on my own", "I asked for help when I needed it",
           "I talked to my team", "I stayed calm", "I followed the instructions", "I checked my work",
           "I used my checklist", "I helped someone"],
    hard: ["It was busy", "It was noisy", "The plan changed", "I did not understand the instructions",
           "I felt tired", "I felt worried", "I did not know who to ask", "The task was new"],
    next: ["check my list", "ask if I am not sure", "take my break on time", "say hello to my team",
           "work a bit faster", "check my work", "use my headphones on breaks", "write down the instructions"],
    // "What did I do?" for each target, used when a rotation has no word bank yet (in PfA order)
    target: [
      ["I asked for help", "I followed the instructions", "I finished my task", "I checked my work", "I tried something new", "I worked on my own"],
      ["I used my checklist", "I got ready on my own", "I travelled on my own", "I managed my time", "I remembered what to do next", "I sorted a problem"],
      ["I took my break on time", "I drank water", "I ate my lunch", "I told someone how I felt", "I used my calm strategy", "I wore the right clothes"],
      ["I said hello to my team", "I spoke to my supervisor", "I worked with someone", "I said thank you", "I asked a question", "I helped someone"]
    ]
  }
};
