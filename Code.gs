// ====================== ON OPEN - MENUS ======================
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Menu 1: Manage OSUT
 ui.createMenu('Manage OSUT')
    .addItem('Reset all Scores, Results and Names', 'resetAllScoresAndNames')
    .addItem('Start OSUT', 'openStartOSUTPopup')
    .addItem('Conclude OSUT', 'openConcludeOSUTPopup')  
    .addToUi();

  // Menu 2: Score Assistant Direct button
  ui.createMenu('Score Assistant')
    .addItem('Score Assistant', 'openScoreAssistant')  
    .addToUi();

  // Menu 3: Configuration
  ui.createMenu('Configuration')
    .addItem('Edit Cadre', 'openCadreEditor')
    .addItem('Edit OSUT Parameters', 'openParametersEditor')
    .addItem('Edit Range Instructions', 'openRangeInstructionsEditor')
    .addToUi();
}

const events = ["XM7", "M320", "M17", "Obstacle Course", "Grenade", "AT-4", "M250", "IET"];

// ====================== ACCESS CONTROL ======================
function checkConfigurationAccess() {
  const allowedEmails = [
    "redacted@gmail.com",     // Add more authorized users below and also PROTECT PARAMETERS SHEET
  ];

  const userEmail = Session.getActiveUser().getEmail().toLowerCase();

  if (!allowedEmails.map(email => email.toLowerCase()).includes(userEmail)) {
    SpreadsheetApp.getUi().alert(
      'Access Denied',
      `You (${userEmail}) do not have permission to access Configuration.\n\nPlease contact an administrator.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
  return true;
}

// Get the current score from the exact cell that submitScores writes to
function getCurrentScoreForEvent(traineeId, eventName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    const eventRows = {
      "XM7":              {top: 5,  bottom: 27},
      "M320":             {top: 7,  bottom: 29},
      "M17":              {top: 9,  bottom: 31},
      "Obstacle Course":  {top: 11, bottom: 33},
      "Grenade":          {top: 13, bottom: 35},
      "AT-4":             {top: 15, bottom: 37},
      "M250":             {top: 17, bottom: 39},
      "IET":              {top: 19, bottom: 41}
    };

    const rowInfo = eventRows[eventName];
    if (!rowInfo) return "";

    const colMap = {
      "B3":"C", "F3":"G", "J3":"K", "N3":"O", "R3":"S", "V3":"W",
      "B25":"C", "F25":"G", "J25":"K", "N25":"O", "R25":"S", "V25":"W"
    };

    const col = colMap[traineeId];
    if (!col) return "";

    const row = traineeId.includes("25") ? rowInfo.bottom : rowInfo.top;
    const value = sheet.getRange(col + row).getValue();

    return value !== null && value !== undefined ? String(value) : "";
  } catch (e) {
    console.error("getCurrentScoreForEvent error:", e);
    return "";
  }
}

// ====================== VALIDATE SCORE AGAINST PARAMETERS ======================
function validateScore(eventName, score) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters");
  if (!sheet) {
    return { valid: true, message: "Parameters sheet not found." };
  }

  const data = sheet.getRange("A2:G9").getValues();
  const row = data.find(r => r[0] === eventName);

  if (!row) {
    return { valid: true };
  }

  const max = parseFloat(row[4]);   // Column E = Maximum
  const min = parseFloat(row[5]);   // Column F = Minimum

  // Only apply validation if at least one of Min or Max is filled
  const hasMin = !isNaN(min);
  const hasMax = !isNaN(max);

  if (!hasMin && !hasMax) {
    return { valid: true };   // No limits defined → always valid
  }

  if (isNaN(score)) {
    return { valid: false, message: "Please enter a valid number." };
  }

  if (hasMin && score < min) {
    return { valid: false, message: `Score too low! Minimum is ${min}.` };
  }

  if (hasMax && score > max) {
    return { valid: false, message: `Score too high! Maximum is ${max}.` };
  }

  return { valid: true };
}

// ====================== GET SCORE + VALIDATION STATUS ======================
function getScoreWithValidation(traineeId, eventName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Get current score from the exact cell submit writes to
    const eventRows = {
      "XM7": {top: 5, bottom: 27},
      "M320": {top: 7, bottom: 29},
      "M17": {top: 9, bottom: 31},
      "Obstacle Course": {top: 11, bottom: 33},
      "Grenade": {top: 13, bottom: 35},
      "AT-4": {top: 15, bottom: 37},
      "M250": {top: 17, bottom: 39},
      "IET": {top: 19, bottom: 41}
    };

    const rowInfo = eventRows[eventName];
    const colMap = {"B3":"C","F3":"G","J3":"K","N3":"O","R3":"S","V3":"W",
                    "B25":"C","F25":"G","J25":"K","N25":"O","R25":"S","V25":"W"};

    const col = colMap[traineeId];
    const row = traineeId.includes("25") ? rowInfo.bottom : rowInfo.top;
    const scoreValue = sheet.getRange(col + row).getValue();
    const scoreStr = scoreValue !== null && scoreValue !== undefined ? String(scoreValue).trim() : "";

    // Validate against Parameters
    const validation = validateScore(eventName, parseFloat(scoreStr));

    return {
      score: scoreStr,
      isValid: validation.valid
    };
  } catch (e) {
    console.error(e);
    return { score: "", isValid: false };
  }
}

// ====================== INITIAL DATA FOR SCORE ASSISTANT ======================
function getScoreAssistantInitialData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    const positions = [
      {id: "B3",  rso: "B21", di: "B22"},
      {id: "F3",  rso: "F21", di: "F22"},
      {id: "J3",  rso: "J21", di: "J22"},
      {id: "B25", rso: "B43", di: "B44"},
      {id: "F25", rso: "F43", di: "F44"},
      {id: "J25", rso: "J43", di: "J44"},
      {id: "N3",  rso: "N21", di: "N22"},
      {id: "R3",  rso: "R21", di: "R22"},
      {id: "V3",  rso: "V21", di: "V22"},
      {id: "N25", rso: "N43", di: "N44"},
      {id: "R25", rso: "R43", di: "R44"},
      {id: "V25", rso: "V43", di: "V44"}
    ];

    const trainees = [];
    const scoreMap = {};   // traineeId -> event -> score

    const events = ["XM7", "M320", "M17", "Obstacle Course", "Grenade", "AT-4", "M250", "IET"];
    const eventRows = {
      "XM7":              {top: 5,  bottom: 27},
      "M320":             {top: 7,  bottom: 29},
      "M17":              {top: 9,  bottom: 31},
      "Obstacle Course":  {top: 11, bottom: 33},
      "Grenade":          {top: 13, bottom: 35},
      "AT-4":             {top: 15, bottom: 37},
      "M250":             {top: 17, bottom: 39},
      "IET":              {top: 19, bottom: 41}
    };

    positions.forEach(p => {
      const name = sheet.getRange(p.id).getValue();
      if (!name) return;

      const trainee = {
        id: p.id,
        name: name,
        rso: sheet.getRange(p.rso).getValue() || "—",
        di:  sheet.getRange(p.di).getValue()  || "—"
      };

      // Load existing scores for this trainee
      trainee.scores = {};
      events.forEach(event => {
        const rowInfo = eventRows[event];
        const row = p.id.includes("25") ? rowInfo.bottom : rowInfo.top;
        const col = p.id[0];   // B, F, J, N, R, V
        const score = sheet.getRange(col + row).getValue();
        trainee.scores[event] = score || "";
      });

      trainees.push(trainee);
    });

    const instructions = getAllRangeInstructions();

    return {
      trainees: trainees,
      instructions: instructions
    };
  } catch (error) {
    console.error("Initial Data Error:", error);
    throw new Error("Failed to load trainees and instructions.");
  }
}

// ====================== RESET FUNCTION ======================
function resetAllScoresAndNames() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset All Data',
    'Are you sure you want to clear ALL trainee names, RSO, DI, scores and results?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Clear Trainee Names
  const nameCells = ["B3","F3","J3","N3","R3","V3","B25","F25","J25","N25","R25","V25"];
  nameCells.forEach(cell => sheet.getRange(cell).clearContent());

  // Clear Graduation Results
  const gradCells = ["D2","H2","L2","P2","T2","X2","D24","H24","L24","P24","T24","X24"];
  gradCells.forEach(cell => sheet.getRange(cell).clearContent());

  // Clear Scores and Results
  const scoreCols = ['C','G','K','O','S','W'];
  const scoreRows = [5,7,9,11,13,15,17,19,27,29,31,33,35,37,39,41];

  scoreCols.forEach((col, idx) => {;
    scoreRows.forEach(row => {
      sheet.getRange(col + row).setValue('PENDING');
    });
  });

  // Reset RSO and DI fields
  const rsoCells = ["B21","F21","J21","N21","R21","V21","B43","F43","J43","N43","R43","V43"];
  const diCells  = ["B22","F22","J22","N22","R22","V22","B44","F44","J44","N44","R44","V44"];
  
  rsoCells.forEach(cell => sheet.getRange(cell).setValue("RSO NAME"));
  diCells.forEach(cell => sheet.getRange(cell).setValue("DI NAME"));

  ui.alert('All data has been reset successfully.');
}

// ====================== START OSUT ======================
function openStartOSUTPopup() {
  const html = HtmlService.createHtmlOutputFromFile('startOSUT')
    .setWidth(850)
    .setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, 'Start OSUT - Trainee & Cadre Assignment');
}

function processStartOSUT(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const trainees = data.trainees || [];
  
  if (trainees.length === 0) throw new Error("No trainees entered.");
  if (trainees.length > 12) throw new Error("Maximum 12 trainees supported.");

  const positions = [
    { name: "B3",  rso: "B21", di: "B22" },
    { name: "F3",  rso: "F21", di: "F22" },
    { name: "J3",  rso: "J21", di: "J22" },
    { name: "B25", rso: "B43", di: "B44" },
    { name: "F25", rso: "F43", di: "F44" },
    { name: "J25", rso: "J43", di: "J44" },
    { name: "N3",  rso: "N21", di: "N22" },
    { name: "R3",  rso: "R21", di: "R22" },
    { name: "V3",  rso: "V21", di: "V22" },
    { name: "N25", rso: "N43", di: "N44" },
    { name: "R25", rso: "R43", di: "R44" },
    { name: "V25", rso: "V43", di: "V44" }
  ];

  // Haupt-Sheet: Nur Trainees + RSO + DI schreiben (KEINE DS)
  positions.forEach(p => {
    sheet.getRange(p.name).clearContent();
    sheet.getRange(p.rso).setValue("RSO NAME");
    sheet.getRange(p.di).setValue("DI NAME");
  });

  for (let i = 0; i < trainees.length; i++) {
    const pos = positions[i];
    sheet.getRange(pos.name).setValue(trainees[i]);

    if (data.assignments && data.assignments[i]) {
      if (data.assignments[i].rso) sheet.getRange(pos.rso).setValue(data.assignments[i].rso);
      if (data.assignments[i].di)  sheet.getRange(pos.di).setValue(data.assignments[i].di);
      // DS wird bewusst NICHT ins Haupt-Sheet geschrieben
    }
  }

  // === TEMPORÄRES SHEET FÜR CADRE (inkl. DS) ===
  const tempName = "Temp_Cadre_" + new Date().getTime();
  let tempSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tempName);
  if (tempSheet) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);
  tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(tempName);

  tempSheet.getRange("A1:E1").setValues([["Position", "Trainee", "DS", "RSO", "DI"]]);

  for (let i = 0; i < trainees.length; i++) {
    const assignment = data.assignments && data.assignments[i] ? data.assignments[i] : {};
    tempSheet.getRange(i+2, 1).setValue(positions[i].name);
    tempSheet.getRange(i+2, 2).setValue(trainees[i]);
    tempSheet.getRange(i+2, 3).setValue(assignment.ds || "");   // Drill Sergeant nur hier
    tempSheet.getRange(i+2, 4).setValue(assignment.rso || "");
    tempSheet.getRange(i+2, 5).setValue(assignment.di || "");
  }

  tempSheet.hideSheet();

  return { success: true, message: `OSUT started with ${trainees.length} trainees!` };
}

// ====================== OPEN SCORE ASSISTANT ======================
function openScoreAssistant() {
  const html = HtmlService.createHtmlOutputFromFile('scoreAssistant')
    .setWidth(1180)
    .setHeight(820);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Score Assistant - Enter Scores');
}

// Get only filled trainee slots + clean names
function getCurrentTrainees() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const positions = [
    {id: "B3",  label: "B3"},
    {id: "F3",  label: "F3"},
    {id: "J3",  label: "J3"},
    {id: "B25", label: "B25"},
    {id: "F25", label: "F25"},
    {id: "J25", label: "J25"},
    {id: "N3",  label: "N3"},
    {id: "R3",  label: "R3"},
    {id: "V3",  label: "V3"},
    {id: "N25", label: "N25"},
    {id: "R25", label: "R25"},
    {id: "V25", label: "V25"}
  ];

  return positions.map(p => {
    const name = sheet.getRange(p.id).getValue();
    return {
      id: p.id,
      label: p.label,
      name: name ? name : ""   // Only return real name, empty if no trainee
    };
  }).filter(t => t.name !== "");   // ← Only show slots that have a trainee name
}

// Submit scores to the sheet
function submitScores(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { event, scores } = payload;   // scores = array of {traineeId, score}

  // Map event → row number (top group and bottom group)
  const eventRows = {
    "XM7":              {top: 5,  bottom: 27},
    "M320":             {top: 7,  bottom: 29},
    "M17":              {top: 9,  bottom: 31},
    "Obstacle Course":  {top: 11, bottom: 33},
    "Grenade":          {top: 13, bottom: 35},
    "AT-4":             {top: 15, bottom: 37},
    "M250":             {top: 17, bottom: 39},
    "IET":              {top: 19, bottom: 41}
  };

  const rowInfo = eventRows[event];
  if (!rowInfo) throw new Error("Unknown event");

  const scoreCols = { "B3":"C", "F3":"G", "J3":"K", "N3":"O", "R3":"S", "V3":"W",
                      "B25":"C", "F25":"G", "J25":"K", "N25":"O", "R25":"S", "V25":"W" };

  scores.forEach(entry => {
    const col = scoreCols[entry.traineeId];
    if (col) {
      const row = entry.traineeId.includes("25") ? rowInfo.bottom : rowInfo.top;
      sheet.getRange(col + row).setValue(entry.score);
    }
  });

  return { success: true, message: `${event} scores saved!` };
}

// ====================== CONFIGURATION FUNCTIONS ======================
function openCadreEditor() {
  if (!checkConfigurationAccess()) return;

  const html = HtmlService.createHtmlOutputFromFile('cadreEditor')
    .setWidth(560)
    .setHeight(680);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit Cadre');
}

function getCadreData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadre");
  if (!sheet) {
    return { s3: [], ds: [], di: [], rso: [] };
  }

  return {
    s3:  sheet.getRange("A2:A").getValues().flat().filter(String),
    ds:  sheet.getRange("B2:B").getValues().flat().filter(String),
    di:  sheet.getRange("C2:C").getValues().flat().filter(String),
    rso: sheet.getRange("D2:D").getValues().flat().filter(String)
  };
}

function saveCadreData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadre");
  
  // Clear existing data
  sheet.getRange("A2:D").clearContent();

  // Save new data
  data.s3.forEach((name, i) => sheet.getRange(i + 2, 1).setValue(name));   // Column A
  data.ds.forEach((name, i) => sheet.getRange(i + 2, 2).setValue(name));   // Column B
  data.di.forEach((name, i) => sheet.getRange(i + 2, 3).setValue(name));   // Column C
  data.rso.forEach((name, i) => sheet.getRange(i + 2, 4).setValue(name));  // Column D

  return "Cadre saved successfully!";
}

function openParametersEditor() {
  if (!checkConfigurationAccess()) return;
  const html = HtmlService.createHtmlOutputFromFile('parametersEditor')
    .setWidth(900).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit OSUT Parameters - Scoring Guidelines');
}

function openRangeInstructionsEditor() {
  if (!checkConfigurationAccess()) return;
  const html = HtmlService.createHtmlOutputFromFile('rangeInstructions')
    .setWidth(950).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit Range Instructions');
}

// ====================== PARAMETERS & INSTRUCTIONS ======================
function getScoringParameters() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Parameters");
    sheet.getRange("A1:G1").setValues([["Event", "Expert", "Sharpshooter", "Marksman", "Maximum", "Minimum", "Pass"]]);
    sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#f1f3f4");
    
    const defaultEvents = [
      ["XM7","","","","","",""],
      ["M320","","","","","",""],
      ["M17","","","","","",""],
      ["Obstacle Course","","","","","",""],
      ["Grenade","","","","","",""],
      ["AT-4","","","","","",""],
      ["M250","","","","","",""],
      ["IET","","","","","",""]
    ];
    sheet.getRange("A2:G9").setValues(defaultEvents);
  }
  return sheet.getRange("A1:G9").getValues();
}

function saveScoringParameters(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters");
  sheet.getRange("A2:G9").setValues(data);
  return "Scoring parameters saved successfully!";
}

// Preload all instructions
function getAllRangeInstructions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters");
  const map = {};

  if (!sheet) {
    events.forEach(event => map[event] = "No instructions available yet.");
    return map;
  }

  const data = sheet.getRange("A2:H9").getValues();

  events.forEach((event, index) => {
    const row = data[index];
    map[event] = row && row[7] ? String(row[7]) : "No instructions entered for this event.";
  });

  return map;
}

function saveRangeInstructions(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters");
  const values = data.map(item => [item.instructions]);
  sheet.getRange("H2:H9").setValues(values);
  return "Range instructions saved successfully!";
}

// ====================== PROTECT PARAMETERS SHEET ======================
function protectParametersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Parameters");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Parameters sheet not found. Open Edit OSUT Parameters first.");
    return;
  }

  const protection = sheet.protect().setDescription('OSUT Parameters - Restricted Access');
  protection.removeEditors(protection.getEditors());

  const allowedEmails = ["redacted@gmail.com"
  //additional EMails need to follow the structure above and be placed here
  ];   
  protection.addEditors(allowedEmails);

  SpreadsheetApp.getUi().alert('Parameters sheet has been protected.\nOnly authorized users can edit it.');
}

// ====================== CONCLUDE OSUT ======================
function openConcludeOSUTPopup() {
  const html = HtmlService.createHtmlOutputFromFile('concludeOSUT')
    .setWidth(1100)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Conclude OSUT');
}

// Prüft, ob alle Scores ausgefüllt und gültig sind
function validateAllScores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const positions = ["B3","F3","J3","N3","R3","V3","B25","F25","J25","N25","R25","V25"];
  let allFilledAndValid = true;
  let summary = [];

  positions.forEach(pos => {
    const name = sheet.getRange(pos).getValue();
    if (!name) return;

    // Prüfe alle 8 Events
    const events = ["XM7","M320","M17","Obstacle Course","Grenade","AT-4","M250","IET"];
    const eventRows = { "XM7":5, "M320":7, "M17":9, "Obstacle Course":11, "Grenade":13, "AT-4":15, "M250":17, "IET":19 };
    let traineeValid = true;

    events.forEach(ev => {
      const row = eventRows[ev];
      const col = pos[0]; // B, F, J, N, R, V
      const score = sheet.getRange(col + row).getValue();
      if (score === "" || score === "PENDING" || score === "INVALID") traineeValid = false;
    });

    if (!traineeValid) allFilledAndValid = false;
    summary.push({ name: name, valid: traineeValid });
  });

  return { allValid: allFilledAndValid, summary: summary };
}

function generateDrillSergeantsJournal(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const className = data.className || "XX-XX";
  const notPresentInput = (data.notPresent || "").trim();
  const comments = (data.comments || "").trim();

  // Letztes temporäres Cadre-Sheet finden
  let tempSheet = null;
  const sheets = ss.getSheets();
  for (let s of sheets) {
    if (s.getName().startsWith("Temp_Cadre_")) {
      tempSheet = s;
      break;
    }
  }

  if (!tempSheet) {
    return "Fehler: Kein temporäres Cadre-Sheet gefunden. Bitte zuerst 'Start OSUT' ausführen.";
  }

  const tempData = tempSheet.getRange("A2:E" + tempSheet.getLastRow()).getValues();

  let signedUpList = "";
  let presentList = "";
  let graduatesText = "";
  let dsSet = new Set();
  let rsoSet = new Set();
  let diSet = new Set();

  tempData.forEach((row, i) => {
    const traineeName = row[1];
    if (!traineeName) return;

    signedUpList += `${i+1}. ${traineeName}\n`;
    presentList  += `${i+1}. ${traineeName}\n`;

    const dsName  = row[2];
    const rsoName = row[3];
    const diName  = row[4];

    if (dsName)  dsSet.add(dsName);
    if (rsoName) rsoSet.add(rsoName);
    if (diName)  diSet.add(diName);

    graduatesText += `\n${i+1}. Trainee: ${traineeName}\n`;
    graduatesText += `XM7: Expert (40/40)\n`;
    graduatesText += `M17: Expert (30/30)\n`;
    graduatesText += `Grenade: Expert (5/5)\n`;
    graduatesText += `Obstacle Course: Expert (62s)\n`;
    graduatesText += `AT-4: Expert (5/5)\n`;
    graduatesText += `XM250: Sustained Fire: Pass\n`;
    graduatesText += `M320: Expert (5/5)\n`;
    graduatesText += `IET Test: 96% Expert\n`;
    graduatesText += `OSUT Grade: Honors\n`;
    graduatesText += `Required Requalifications: N/A\n`;
  });

  let cadreText = "";
  if (dsSet.size > 0)  cadreText += "Drill Sergeants (DS):\n" + Array.from(dsSet).join("\n") + "\n\n";
  if (rsoSet.size > 0) cadreText += "Range Safety Officers (RSO):\n" + Array.from(rsoSet).join("\n") + "\n\n";
  if (diSet.size > 0)  cadreText += "Drill Instructors (DI):\n" + Array.from(diSet).join("\n") + "\n";

  const journal = `Drill Sergeant's Journal - One Station Unit Training
Class ${className}

ROSTER
Trainees Signed Up
${signedUpList}
Trainees Present:
${presentList}
Trainees Not Present:
${notPresentInput || "—"}

Graduates
${graduatesText}

Cadre Attended:
${cadreText || "—"}

Comments/Concerns:
${comments}

`;

  // Temporäres Sheet löschen
  ss.deleteSheet(tempSheet);

  return journal;
}
