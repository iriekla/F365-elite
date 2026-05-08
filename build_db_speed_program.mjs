import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "D:/tri/outputs/db_speed_program";
const outputPath = `${outputDir}/14_week_defensive_back_speed_program.xlsx`;

await fs.mkdir(outputDir, { recursive: true });

const wb = Workbook.create();

const colors = {
  navy: "#1D3557",
  teal: "#2A9D8F",
  gold: "#E9C46A",
  coral: "#E76F51",
  light: "#F5F7FA",
  gray: "#E5E7EB",
  dark: "#111827",
  white: "#FFFFFF",
};

function title(sheet, range, text, subtitle = "") {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white, size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "middle",
  };
  if (subtitle) {
    const row = Number(range.match(/\d+/)?.[0] ?? 1) + 1;
    sheet.getRange(`A${row}:H${row}`).merge();
    sheet.getRange(`A${row}:H${row}`).values = [[subtitle]];
    sheet.getRange(`A${row}:H${row}`).format = {
      fill: colors.light,
      font: { italic: true, color: colors.dark },
      wrapText: true,
    };
  }
}

function styleHeader(range, fill = colors.teal) {
  range.format = {
    fill,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "middle",
    wrapText: true,
  };
}

function styleBody(range) {
  range.format = {
    wrapText: true,
    verticalAlignment: "top",
    borders: {
      insideHorizontal: { style: "Continuous", color: colors.gray },
      insideVertical: { style: "Continuous", color: colors.gray },
      edgeBottom: { style: "Continuous", color: colors.gray },
      edgeTop: { style: "Continuous", color: colors.gray },
      edgeLeft: { style: "Continuous", color: colors.gray },
      edgeRight: { style: "Continuous", color: colors.gray },
    },
  };
}

function setWidths(sheet, widths) {
  widths.forEach((w, i) => {
    sheet.getRangeByIndexes(0, i, 1, 1).format.columnWidthPx = w;
  });
}

const overview = wb.worksheets.add("Overview");
overview.showGridLines = false;
title(
  overview,
  "A1:H1",
  "14-Week Defensive Back Acceleration & Lateral Explosiveness Program",
  "Built for a high school senior defensive back. Train speed fresh, keep sprint reps high quality, and stop when mechanics or output drops."
);
setWidths(overview, [120, 135, 135, 135, 135, 135, 135, 135]);

overview.getRange("A4:H4").values = [[
  "Phase",
  "Weeks",
  "Primary Focus",
  "Speed Intensity",
  "Sprint Volume / Session",
  "Plyo Contacts",
  "COD / DB Reps",
  "Recovery Emphasis",
]];
styleHeader(overview.getRange("A4:H4"));
overview.getRange("A5:H8").values = [
  ["Phase 1", "1-4", "Mechanics, deceleration, lateral control, tissue prep", "75-90%", "120-200 yd", "40-70", "12-20", "Build base, full rest, correct landings"],
  ["Phase 2", "5-8", "Power, first-step force, lateral push-off", "90-97%", "180-260 yd", "50-80", "16-24", "48 hr between hard lower-body speed days"],
  ["Phase 3", "9-12", "DB-specific reactive speed and coverage transitions", "92-100%", "160-240 yd", "35-60", "20-35", "Monitor hamstring/groin; reduce volume if flat"],
  ["Phase 4", "13-14", "Peak, sharpen, test, stay fresh", "95-100%", "80-160 yd", "20-40", "6-16", "Low volume, high readiness, no hard conditioning"],
];
styleBody(overview.getRange("A5:H8"));

overview.getRange("A11:D11").values = [["Weekly Day", "Training Theme", "Primary Work", "Notes"]];
styleHeader(overview.getRange("A11:D11"), colors.coral);
overview.getRange("A12:D18").values = [
  ["Day 1", "Acceleration + lower-body strength", "Starts, 10-20 yd sprints, sled/hill work, main lift", "Highest priority day; train fresh"],
  ["Day 2", "Lateral explosiveness + DB movement", "Bounds, shuffle-to-sprint, backpedal breaks, hip flips", "Focus on clean plant angles"],
  ["Day 3", "Tempo + mobility", "65-70% tempo, hips, ankles, adductors", "Recovery and aerobic support"],
  ["Day 4", "Max velocity + plyometrics", "Build-ups, flying 10/20s, broad jumps, pogos", "Full rest; no fatigue chasing"],
  ["Day 5", "Change of direction + power strength", "5-10-5, T/L drill, reactive breaks, cleans/jump shrugs", "Game-speed cuts with full control"],
  ["Day 6", "Optional light skill", "Walk-through DB work, ball skills, easy tempo", "Keep it easy"],
  ["Day 7", "Off", "No structured work", "Sleep, food, hydration"],
];
styleBody(overview.getRange("A12:D18"));

overview.getRange("F11:H11").values = [["Metric", "Start", "Week 14 Test"]];
styleHeader(overview.getRange("F11:H11"), colors.coral);
overview.getRange("F12:H18").values = [
  ["10-yard split", "", ""],
  ["20-yard sprint", "", ""],
  ["40-yard dash", "", ""],
  ["5-10-5 shuttle", "", ""],
  ["Broad jump", "", ""],
  ["Vertical jump", "", ""],
  ["Body weight", "", ""],
];
styleBody(overview.getRange("F12:H18"));

overview.getRange("J4:M8").values = [
  ["Phase", "Low Volume", "High Volume", "Speed Intensity High"],
  ["Phase 1", 120, 200, 90],
  ["Phase 2", 180, 260, 97],
  ["Phase 3", 160, 240, 100],
  ["Phase 4", 80, 160, 100],
];
const chart = overview.charts.add("bar", overview.getRange("J4:M8"));
chart.title = "Sprint Volume and Intensity by Phase";
chart.hasLegend = true;
chart.setPosition("J10", "Q25");

const phase = wb.worksheets.add("Phase Details");
phase.showGridLines = false;
title(phase, "A1:H1", "Phase Details");
setWidths(phase, [95, 90, 210, 125, 145, 145, 190, 220]);
phase.getRange("A3:H3").values = [[
  "Phase",
  "Weeks",
  "Focus",
  "Intensity",
  "Volume",
  "Speed Specifics",
  "Energy System Targets",
  "Recovery Protocol",
]];
styleHeader(phase.getRange("A3:H3"));
phase.getRange("A4:H7").values = [
  ["1", "1-4", "Sprint posture, first-step mechanics, deceleration strength, lateral landing control, ankle/hip/adductor durability.", "Speed 75-90%; lift 65-80%; plyos controlled.", "120-200 yd acceleration; 120-180 yd max velocity; 40-70 plyo contacts; 12-20 COD reps.", "Wall drives, falling starts, 2-point starts, hill/sled 10 yd, skater sticks, W-drill, flying 10s.", "ATP-PC mechanics for 1-6 sec bursts; aerobic support through 65-70% tempo.", "Full rest between quality reps; 1 off day; mobility daily; stop if sprint quality drops 10%."],
  ["2", "5-8", "Force production, faster 0-10 yd, violent lateral push-off, sharper plant and drive, hip flip power.", "Speed 90-97%; lift 75-88%; plyos moderate-high.", "180-260 yd acceleration; 150-220 yd max velocity; 50-80 contacts; 16-24 COD reps.", "3-point starts, sled-to-free sprints, Heiden jump to sprint, mirror drill, flying 20s, sprint-float-sprint.", "ATP-PC power and alactic repeatability: short explosive sets with enough rest to stay fast.", "48 hr between hard lower-body speed days; reduce volume 20-30% if legs feel flat."],
  ["3", "9-12", "Transfer speed into DB movement: backpedal break, open hip, speed turn, route reaction, repeated coverage reps.", "Speed 92-100%; lift 70-85%; game-speed COD.", "160-240 yd acceleration; 120-200 yd max velocity; 20-35 DB reps; 35-60 contacts.", "Reaction starts, coach-point drills, backpedal to 45/90, speed turn, comeback break, route reaction.", "ATP-PC plus repeat burst tolerance for 4-7 sec football plays with 20-40 sec rest exposure.", "Monitor groin/hamstrings; no max sprinting sore; low-intensity tempo only for conditioning."],
  ["4", "13-14", "Peak and testing: stay explosive, reduce fatigue, sharpen starts and cuts.", "Speed 95-100%; lift 65-80%; low volume.", "80-160 yd sprint volume; 20-40 plyo contacts; 6-16 COD/DB reps.", "Short starts, flying 10s, light DB breaks, timed 5-10-5, final tests.", "Full ATP-PC recovery and nervous system freshness; no hard conditioning.", "48 hr from heavy lower lift to testing; full rest between tests; prioritize sleep and hydration."],
];
styleBody(phase.getRange("A4:H7"));

const weekly = wb.worksheets.add("Weekly Plan");
weekly.showGridLines = false;
title(weekly, "A1:I1", "14-Week Weekly Plan");
setWidths(weekly, [70, 95, 170, 170, 170, 170, 170, 170, 170]);
weekly.getRange("A3:I3").values = [["Week", "Phase", "Day 1 Accel", "Day 2 Lateral/DB", "Day 3 Recovery", "Day 4 Max Speed", "Day 5 COD/Strength", "Day 6 Optional", "Notes"]];
styleHeader(weekly.getRange("A3:I3"));
const weeklyRows = [];
for (let w = 1; w <= 14; w++) {
  const ph = w <= 4 ? "Phase 1" : w <= 8 ? "Phase 2" : w <= 12 ? "Phase 3" : "Phase 4";
  weeklyRows.push([
    w,
    ph,
    w <= 4 ? "Wall/falling starts; 6x10 + 5x15; hill/sled 5x10; strength 4x5" :
    w <= 8 ? "3-point starts 6x10; 5x20; sled 5x15; free 3x20; strength 5x3" :
    w <= 12 ? "Reaction starts; 8x10 + 6x15; sled 4x10; free 4x20; strength 3x3" :
    "Sharpen: 4-6x10; 3-4x20; light strength 2-3 sets",
    w <= 4 ? "Skater sticks, lateral bounds, shuffle-sprint, backpedal 45, W-drill" :
    w <= 8 ? "Continuous bounds, Heiden-to-sprint, mirror, shuffle-shuffle-sprint" :
    w <= 12 ? "Mirror, backpedal 45/90, speed turn, comeback break, route reaction" :
    "Low-volume shuffle-sprint and 3-4 clean DB breaks",
    "Tempo 8-12x60-80 yd at 65-70%; mobility; soft tissue",
    w <= 4 ? "Build-ups 5x30; flying 10s x4; broad jump 5x2" :
    w <= 8 ? "Flying 10s x5; flying 20s x3; sprint-float-sprint x4" :
    w <= 12 ? "Flying 10s x4; flying 20s x3; 30 yd sprint x4" :
    "3 flying 10s; keep full rest",
    w <= 4 ? "Decel 6 reps; 5-10-5 tech x5; T-drill x4; front squat/RDL" :
    w <= 8 ? "Timed 5-10-5 x6; L-drill x4; reactive cone breaks; clean/jump shrug" :
    w <= 12 ? "Coach-point x6-8; route reaction x6; 5-10-5 x4; maintain strength" :
    "Week 13: light timed reps. Week 14: test after 48 hr fresh window",
    "Ball skills, walk-through technique, easy mobility",
    w === 4 || w === 8 || w === 12 ? "Slight deload if times/jumps are down." : w === 14 ? "Test week. No hard conditioning." : "Track quality and recovery."
  ]);
}
weekly.getRange(`A4:I${weeklyRows.length + 3}`).values = weeklyRows;
styleBody(weekly.getRange(`A4:I${weeklyRows.length + 3}`));

const speed = wb.worksheets.add("Speed & COD Details");
speed.showGridLines = false;
title(speed, "A1:H1", "Speed Training Specifics");
setWidths(speed, [110, 130, 190, 120, 115, 115, 140, 210]);
speed.getRange("A3:H3").values = [["Category", "Phase", "Drill", "Sets/Reps", "Distance", "Intensity", "Rest", "Coaching Target"]];
styleHeader(speed.getRange("A3:H3"));
speed.getRange("A4:H27").values = [
  ["Acceleration", "1", "Wall drives", "3x10 each", "N/A", "Control", "30-45 sec", "Positive shin angle; push ground back"],
  ["Acceleration", "1", "Falling starts", "6 reps", "10 yd", "85-90%", "60-90 sec", "Clean first 3 steps"],
  ["Acceleration", "1", "2-point starts", "5 reps", "15 yd", "85-90%", "90 sec", "Stay low, rise naturally"],
  ["Lateral", "1", "Skater jump and stick", "3x5 each", "N/A", "Moderate", "45-60 sec", "Own the landing before next rep"],
  ["Max Velocity", "1", "Flying 10", "4 reps", "20 build + 10 fly", "90%", "2-3 min", "Tall posture, fast front-side mechanics"],
  ["COD", "1", "5-10-5 technique", "5 reps", "20 yd total", "75-85%", "90 sec", "No false step; low hips on plant"],
  ["Acceleration", "2", "3-point starts", "6 reps", "10 yd", "92-97%", "60-90 sec", "Explosive projection"],
  ["Acceleration", "2", "Sled sprint", "5 reps", "10-15 yd", "Heavy enough to push, not grind", "2 min", "Powerful drive phase"],
  ["Acceleration", "2", "Free sprint after sled", "3 reps", "20 yd", "95-97%", "2 min", "Transfer resisted force to free sprint"],
  ["Lateral", "2", "Heiden jump to sprint", "5 each", "5-10 yd sprint", "High", "60-90 sec", "Push laterally, snap into sprint"],
  ["Max Velocity", "2", "Flying 20", "3 reps", "20 build + 20 fly", "95-97%", "3-4 min", "Relaxed top speed"],
  ["COD", "2", "Timed 5-10-5", "6 reps", "20 yd total", "90-97%", "3-5 min", "Quality over fatigue"],
  ["Acceleration", "3", "Reaction starts", "6 reps", "15 yd", "95-100%", "90-120 sec", "React to visual cue, no guessing"],
  ["DB", "3", "Backpedal to 45 break", "5 each", "5-15 yd", "Game speed", "60-90 sec", "Plant, drive, finish downhill"],
  ["DB", "3", "Speed turn", "5 each", "10-20 yd", "Game speed", "90 sec", "Flip hips without rising"],
  ["DB", "3", "Route reaction", "6-8 reps", "Variable", "Game speed", "60-120 sec", "Eyes controlled, finish on ball"],
  ["Max Velocity", "3", "30-yard sprint", "4 reps", "30 yd", "95-100%", "3 min", "Smooth acceleration into speed"],
  ["COD", "3", "Coach-point drill", "6-8 reps", "Variable", "Game speed", "60-120 sec", "Reactive plant and redirect"],
  ["Acceleration", "4", "10-yard starts", "4-6 reps", "10 yd", "95-100%", "90 sec", "Feel sharp and fresh"],
  ["Max Velocity", "4", "Flying 10", "3 reps", "20 build + 10 fly", "95-100%", "3-4 min", "Fast but relaxed"],
  ["COD", "4", "Timed 5-10-5", "3-4 reps", "20 yd total", "95-100%", "3-5 min", "Peak performance, full rest"],
  ["Testing", "4", "Final tests", "1-3 attempts", "Metric-specific", "100%", "Full recovery", "Best clean effort"],
  ["Recovery Speed", "All", "Tempo runs", "8-12 reps", "60-80 yd", "65-70%", "Walk/45-60 sec", "Aerobic support, not fatigue"],
  ["Repeat Burst", "3", "Pedal-break-sprint", "2-3 sets x5", "5 pedal + 10 sprint", "90-95%", "30-40 sec reps; 3 min sets", "Alactic repeatability"],
];
styleBody(speed.getRange("A4:H27"));

const energy = wb.worksheets.add("Energy Systems");
energy.showGridLines = false;
title(energy, "A1:G1", "Energy System Development");
setWidths(energy, [145, 210, 190, 170, 145, 145, 220]);
energy.getRange("A3:G3").values = [["System", "Why It Matters for DB", "Program Target", "Best Methods", "Work Duration", "Rest", "Do / Avoid"]];
styleHeader(energy.getRange("A3:G3"));
energy.getRange("A4:G6").values = [
  ["ATP-PC", "First step, break on ball, short sprint, jump, hip flip.", "Max power for 1-6 sec efforts.", "5-30 yd sprints, jumps, med ball throws, full-rest COD.", "1-6 sec", "60 sec to 5 min", "Do stay explosive. Avoid conditioning fatigue during speed work."],
  ["Alactic repeatability", "Repeat explosive plays late in a drive without losing speed.", "High-quality bursts with incomplete but adequate rest.", "2-4 sets of 4-6 short DB/sprint reps.", "4-8 sec", "20-40 sec reps; 3 min sets", "Do stop before mechanics collapse. Avoid sloppy grind reps."],
  ["Aerobic support", "Helps recover between plays, sessions, and practices.", "Low-intensity recovery capacity.", "Tempo runs, bike, mobility circuits.", "10-20 min total", "Easy walk/rest", "Do keep 60-70%. Avoid long slow distance replacing speed."],
];
styleBody(energy.getRange("A4:G6"));

const recovery = wb.worksheets.add("Recovery Protocols");
recovery.showGridLines = false;
title(recovery, "A1:H1", "Recovery Protocols and Readiness Rules");
setWidths(recovery, [145, 190, 190, 190, 170, 170, 170, 180]);
recovery.getRange("A3:H3").values = [["Area", "Rule", "Phase 1", "Phase 2", "Phase 3", "Phase 4", "Red Flag", "Action"]];
styleHeader(recovery.getRange("A3:H3"));
recovery.getRange("A4:H12").values = [
  ["Between reps", "Rest enough to keep speed high.", "60-180 sec", "60 sec-4 min", "60 sec-4 min", "90 sec-5 min", "Times slow or steps get sloppy.", "Extend rest or stop."],
  ["Between days", "Protect hard lower-body speed days.", "48 hr if sore", "48 hr required", "48 hr required", "Keep legs fresh", "Heavy legs entering speed day.", "Reduce volume 20-30%."],
  ["Sleep", "Main recovery driver.", "8+ hr", "8+ hr", "8-9 hr", "8-9 hr", "Poor sleep 2 nights straight.", "Cut high-speed volume."],
  ["Mobility", "Hips, ankles, calves, groin daily.", "10 min/day", "10 min/day", "10-15 min/day", "10 min/day", "Stiffness changes mechanics.", "Do recovery session only."],
  ["Hamstrings", "Strength and monitor tightness.", "Nordic/isometric 2-3x/wk", "2x/wk", "1-2x/wk", "1x/wk light", "Tightness increases while sprinting.", "Stop sprinting that day."],
  ["Groin/adductors", "Important for lateral cuts.", "Copenhagen 2x/wk", "2x/wk", "1-2x/wk", "Light only", "Pain on lateral push.", "Stop COD; switch to mobility."],
  ["Conditioning", "Support recovery, do not steal speed.", "Tempo 65-70%", "Tempo after speed days only if fresh", "Repeat burst specific", "None hard", "Breathing workout ruins speed.", "Move conditioning to recovery day."],
  ["Testing readiness", "Fresh nervous system beats more volume.", "Not primary", "Track jumps/times", "Track jumps/times", "Peak", "Broad jump or 10 yd split down.", "Delay testing or lower volume."],
  ["Nutrition/hydration", "Fuel high-output training.", "Protein each meal", "Protein + carbs around training", "Same", "Same", "Cramps, low energy.", "Hydrate, add carbs, reduce volume."],
];
styleBody(recovery.getRange("A4:H12"));

const testing = wb.worksheets.add("Testing Tracker");
testing.showGridLines = false;
title(testing, "A1:H1", "Testing Tracker");
setWidths(testing, [145, 115, 115, 115, 115, 120, 150, 220]);
testing.getRange("A3:H3").values = [["Metric", "Week 0", "Week 4", "Week 8", "Week 12", "Week 14", "Change", "Notes"]];
styleHeader(testing.getRange("A3:H3"));
testing.getRange("A4:H10").values = [
  ["10-yard split", "", "", "", "", "", "", "Lower is better"],
  ["20-yard sprint", "", "", "", "", "", "", "Lower is better"],
  ["40-yard dash", "", "", "", "", "", "", "Lower is better"],
  ["5-10-5 shuttle", "", "", "", "", "", "", "Lower is better"],
  ["Broad jump", "", "", "", "", "", "", "Higher is better"],
  ["Vertical jump", "", "", "", "", "", "", "Higher is better"],
  ["Body weight", "", "", "", "", "", "", "Track context for speed changes"],
];
testing.getRange("G4").formulas = [["=IF(OR(B4=\"\",F4=\"\"),\"\",F4-B4)"]];
testing.getRange("G4:G10").fillDown();
styleBody(testing.getRange("A4:H10"));
testing.getRange("B4:G10").format.numberFormat = "0.00";

for (const sheet of wb.worksheets.items) {
  try {
    sheet.freezePanes.freezeRows(3);
  } catch {}
}

const inspect = await wb.inspect({
  kind: "sheet,table",
  maxChars: 5000,
  tableMaxRows: 6,
  tableMaxCols: 8,
});
console.log(inspect.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

for (const name of ["Overview", "Phase Details", "Weekly Plan", "Speed & COD Details", "Energy Systems", "Recovery Protocols", "Testing Tracker"]) {
  const preview = await wb.render({ sheetName: name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${name.replace(/[^A-Za-z0-9]/g, "_")}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(`saved:${outputPath}`);
