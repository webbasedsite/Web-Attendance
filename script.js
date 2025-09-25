function doPost(e) {
  const params = e.parameter;
  const action = params.action;
  const phone = params.phone?.trim();

  const LOCK_KEY = 'lastRequestTimestamp_' + phone;
  const RATE_LIMIT_MS = 5000; // 5 seconds minimum between requests

  // Rate limiting check
  if (phone) {
    const userProperties = PropertiesService.getScriptProperties();
    const lastRequest = userProperties.getProperty(LOCK_KEY);
    const now = new Date().getTime();

    if (lastRequest && (now - Number(lastRequest) < RATE_LIMIT_MS)) {
      Logger.log(`Rate limit exceeded for phone: ${phone}`);
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, message: "Rate limit exceeded. Please wait before trying again." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // Update last request time
    userProperties.setProperty(LOCK_KEY, now.toString());
  }

  Logger.log(`Received action: ${action} from phone: ${phone}`);

  const ss = SpreadsheetApp.openById("1YCVM4b312cVxSyeMjMEL5McfWj5ZTVbLRFa-uf7tTvg");
  const employeesSheet = ss.getSheetByName("Employees");
  const officesSheet = ss.getSheetByName("Offices");
  const attendanceSheet = ss.getSheetByName("Attendance");

  const jsonResponse = (success, message, data = {}) =>
    ContentService.createTextOutput(JSON.stringify({ success, message, ...data }))
      .setMimeType(ContentService.MimeType.JSON);

  // ----------------------
  // 1️⃣ Login
  // ----------------------
  if (action === "login") {
    Logger.log(`Login attempt for phone: ${phone}`);

    const password = params.password?.trim();

    const employees = employeesSheet.getDataRange().getValues();
    const headers = employees[0];
    const rows = employees.slice(1);

    const phoneCol = headers.indexOf("Phone");
    const passwordCol = headers.indexOf("Password");
    const officeCol = headers.indexOf("OfficeID");
    const roleCol = headers.indexOf("Role");
    const nameCol = headers.indexOf("Name");

    if (phoneCol === -1 || passwordCol === -1 || officeCol === -1 || roleCol === -1 || nameCol === -1) {
      Logger.log("One or more required columns are missing in the Employees sheet");
      return jsonResponse(false, "One or more required columns are missing in the Employees sheet");
    }

    // Find employee matching phone and password only
    const matched = rows.find(r =>
      String(r[phoneCol]).trim() === phone &&
      String(r[passwordCol]).trim() === password
    );

    if (!matched) {
      Logger.log(`Invalid login attempt for phone: ${phone}`);
      return jsonResponse(false, "Invalid phone or password");
    }

    const officeID = matched[officeCol];
    const offices = officesSheet.getDataRange().getValues().slice(1);
    const office = offices.find(o => o[0] === officeID);
    const hubName = office ? office[1] : "";

    Logger.log(`Login success for phone: ${phone}, role: ${matched[roleCol]}`);

    return jsonResponse(true, "Login success", {
      role: matched[roleCol],
      name: matched[nameCol],
      hubName: hubName,
      officeID: officeID
    });
  }

  // ----------------------
  // 2️⃣ Get Offices
  // ----------------------
  if (action === "getOffices") {
    Logger.log("Fetching all offices");

    const offices = officesSheet.getDataRange().getValues().slice(1).map(r => ({
      id: r[0],
      name: r[1],
      number: r[2],
      lat: r[3],
      lng: r[4]
    }));
    return jsonResponse(true, "", { offices });
  }

  // ----------------------
  // 3️⃣ Get Office Location
  // ----------------------
  if (action === "getOfficeLocation") {
    Logger.log(`Getting office location for phone: ${phone}`);

    const employees = employeesSheet.getDataRange().getValues().slice(1);
    const emp = employees.find(r => r[0] === phone || r[2] === phone);
    if (!emp) {
      Logger.log("Employee not found for getOfficeLocation");
      return jsonResponse(false, "Employee not found");
    }
    const officeID = emp[4];
    const offices = officesSheet.getDataRange().getValues().slice(1);
    const office = offices.find(r => r[0] === officeID);
    if (!office) {
      Logger.log("Office not found for getOfficeLocation");
      return jsonResponse(false, "Office not found");
    }
    return jsonResponse(true, "", {
      latitude: office[3],
      longitude: office[4]
    });
  }

  // ----------------------
  // 4️⃣ Check-In / Check-Out
  // ----------------------
  if (action === "Check-In" || action === "Check-Out") {
    Logger.log(`${action} request from employee: ${phone}`);

    const employeeId = phone;
    const shift = params.shift;
    const latitude = parseFloat(params.latitude);
    const longitude = parseFloat(params.longitude);
    const timestamp = new Date(params.timestamp);

    const allRows = attendanceSheet.getDataRange().getValues().slice(1);
    const lastRecord = allRows.reverse().find(r => r[1] === employeeId && r[3] === shift);

    if (action === "Check-In") {
      if (lastRecord) {
        const lastTime = new Date(lastRecord[0]);
        const diffHours = (timestamp - lastTime) / 3600000;
        if (diffHours < 10) {
          Logger.log(`Check-in denied: only ${(10 - diffHours).toFixed(1)} hours since last check-in`);
          return jsonResponse(false, `Cannot check-in yet, wait ${(10 - diffHours).toFixed(1)} hours`);
        }
        const today = new Date().toDateString();
        if (lastTime.toDateString() === today) {
          Logger.log("Check-in denied: already checked-in today for this shift");
          return jsonResponse(false, "Already checked-in today for this shift");
        }
      }
    } else {
      if (!lastRecord || lastRecord[4] !== "Check-In") {
        Logger.log("Check-out denied: no active check-in found");
        return jsonResponse(false, "No active check-in found");
      }
    }

    // Find nearest office
    const offices = officesSheet.getDataRange().getValues().slice(1);
    let nearestOffice = offices[0];
    let minDist = getDistance(latitude, longitude, offices[0][3], offices[0][4]);
    offices.forEach(o => {
      const dist = getDistance(latitude, longitude, o[3], o[4]);
      if (dist < minDist) {
        nearestOffice = o;
        minDist = dist;
      }
    });

    attendanceSheet.appendRow([
      timestamp,
      employeeId,
      nearestOffice[0],
      shift,
      action,
      latitude,
      longitude,
      "Active"
    ]);

    Logger.log(`${action} successful for employee: ${employeeId}`);

    return jsonResponse(true, `${action} successful`);
  }

  // ----------------------
  // 5️⃣ Get History
  // ----------------------
  if (action === "getHistory") {
    Logger.log(`Fetching attendance history for phone: ${phone}`);

    const employeeId = phone;
    const allRows = attendanceSheet.getDataRange().getValues().slice(1);
    const records = allRows.filter(r => r[1] === employeeId).map(r => ({
      timestamp: r[0],
      employeeId: r[1],
      officeId: r[2],
      shift: r[3],
      action: r[4],
      latitude: r[5],
      longitude: r[6],
      status: r[7]
    }));

    return jsonResponse(true, "", { records });
  }

  // ----------------------
  // 6️⃣ Get Agents by Office (for CSV download by Incharge)
  // ----------------------
  if (action === "getAgentsByOffice") {
    Logger.log(`Fetching agents for officeID: ${params.officeID}`);

    const officeID = params.officeID;
    if (!officeID) {
      Logger.log("OfficeID missing in getAgentsByOffice");
      return jsonResponse(false, "OfficeID is required");
    }
    const employees = employeesSheet.getDataRange().getValues().slice(1);
    const headers = employeesSheet.getDataRange().getValues()[0];
    const officeCol = headers.indexOf("OfficeID");
    const nameCol = headers.indexOf("Name");
    const phoneCol = headers.indexOf("Phone");
    const roleCol = headers.indexOf("Role");

    if (officeCol === -1 || nameCol === -1 || phoneCol === -1 || roleCol === -1) {
      Logger.log("Required columns missing in Employees sheet for getAgentsByOffice");
      return jsonResponse(false, "Required columns missing");
    }

    // Filter employees by office and role = 'agent'
    const agents = employees.filter(r => r[officeCol] === officeID && r[roleCol] === "agent").map(r => ({
      name: r[nameCol],
      phone: r[phoneCol]
    }));

    Logger.log(`Found ${agents.length} agents for officeID: ${officeID}`);

    return jsonResponse(true, "", { agents });
  }

  // If no valid action
  Logger.log(`Invalid action received: ${action}`);
  return jsonResponse(false, "Invalid action");
}

// --------------------------
// 🌍 Helper: Haversine Formula
// --------------------------
function getDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) *
    Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}
