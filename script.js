function doPost(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    const phone = params.phone?.trim();

    if (!action) {
      return jsonResponse(false, "Action parameter missing");
    }
    // Some actions like getOffices don't require phone param
    if (!phone && action !== "getOffices" && action !== "getAllEmployees") {
      return jsonResponse(false, "Phone parameter missing");
    }

    const ss = SpreadsheetApp.openById("1YCVM4b312cVxSyeMjMEL5McfWj5ZTVbLRFa-uf7tTvg");
    const employeesSheet = ss.getSheetByName("Employees");
    const officesSheet = ss.getSheetByName("Offices");
    const attendanceSheet = ss.getSheetByName("Attendance");

    const employeesData = employeesSheet.getDataRange().getValues();
    const employeesHeaders = employeesData[0];
    const employeesRows = employeesData.slice(1);

    function jsonResponse(success, message, data = {}) {
      return ContentService.createTextOutput(JSON.stringify({ success, message, ...data }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const phoneCol = employeesHeaders.indexOf("Phone");
    const passwordCol = employeesHeaders.indexOf("Password");
    const officeIdCol = employeesHeaders.indexOf("OfficeID");
    const roleCol = employeesHeaders.indexOf("Role");
    const nameCol = employeesHeaders.indexOf("Name");

    if ([phoneCol, passwordCol, officeIdCol, roleCol, nameCol].includes(-1)) {
      Logger.log("One or more required columns missing in Employees sheet");
      return jsonResponse(false, "Employees sheet missing required columns");
    }

    // Rate limiting (same as before)
    if (phone) {
      const LOCK_KEY = 'lastRequestTimestamp_' + phone;
      const RATE_LIMIT_MS = 5000; // 5 seconds

      const userProperties = PropertiesService.getScriptProperties();
      const lastRequest = userProperties.getProperty(LOCK_KEY);
      const now = Date.now();

      if (lastRequest && (now - Number(lastRequest) < RATE_LIMIT_MS)) {
        Logger.log(`Rate limit exceeded for phone: ${phone}`);
        return jsonResponse(false, "Rate limit exceeded. Please wait.");
      }

      userProperties.setProperty(LOCK_KEY, now.toString());
    }

    Logger.log(`Received action: ${action} from phone: ${phone}`);

    // -------- 1️⃣ LOGIN --------
    if (action === "login") {
      const password = params.password?.trim();
      if (!password) {
        return jsonResponse(false, "Password parameter missing");
      }

      const matched = employeesRows.find(r =>
        String(r[phoneCol]).trim() === phone &&
        String(r[passwordCol]).trim() === password
      );

      if (!matched) {
        Logger.log("Invalid login credentials");
        return jsonResponse(false, "Invalid phone or password");
      }

      const officeID = matched[officeIdCol];
      const offices = officesSheet.getDataRange().getValues().slice(1);
      const office = offices.find(o => o[0] === officeID);
      const hubName = office ? office[1] : "";

      Logger.log(`Login successful: ${phone}, role: ${matched[roleCol]}`);

      // Return requested data including phone and officeID
      return jsonResponse(true, "Login success", {
        role: matched[roleCol],
        name: matched[nameCol],
        hubName: hubName,
        officeID: officeID,
        phone: phone
      });
    }

    // -------- 2️⃣ GET OFFICES --------
    if (action === "getOffices") {
      const offices = officesSheet.getDataRange().getValues().slice(1).map(r => ({
        id: r[0],
        name: r[1],
        number: r[2],
        lat: r[3],
        lng: r[4]
      }));
      return jsonResponse(true, "", { offices });
    }

    // -------- 3️⃣ GET OFFICE LOCATION --------
    if (action === "getOfficeLocation") {
      const emp = employeesRows.find(r => String(r[phoneCol]).trim() === phone);
      if (!emp) {
        Logger.log(`Employee not found for phone: ${phone}`);
        return jsonResponse(false, "Employee not found");
      }

      const officeID = emp[officeIdCol];
      const offices = officesSheet.getDataRange().getValues().slice(1);
      const office = offices.find(r => r[0] === officeID);
      if (!office) {
        Logger.log(`Office not found for ID: ${officeID}`);
        return jsonResponse(false, "Office not found");
      }

      return jsonResponse(true, "", {
        latitude: office[3],
        longitude: office[4]
      });
    }

    // -------- 4️⃣ CHECK-IN / CHECK-OUT --------
    if (action === "Check-In" || action === "Check-Out") {
      const employeeId = phone;
      const shift = params.shift;
      if (!shift) {
        return jsonResponse(false, "Shift parameter missing");
      }
      const latitude = parseFloat(params.latitude);
      const longitude = parseFloat(params.longitude);
      if (isNaN(latitude) || isNaN(longitude)) {
        return jsonResponse(false, "Latitude and longitude parameters required");
      }
      const timestamp = new Date(params.timestamp);
      if (isNaN(timestamp.getTime())) {
        return jsonResponse(false, "Invalid timestamp");
      }

      if (!employeesRows.some(r => String(r[phoneCol]).trim() === phone)) {
        Logger.log("Phone not found in Employees");
        return jsonResponse(false, "Phone number not registered");
      }

      const allRows = attendanceSheet.getDataRange().getValues().slice(1);
      const lastRecord = [...allRows].reverse().find(r => r[1] === employeeId && r[3] === shift);

      if (action === "Check-In") {
        if (lastRecord) {
          const lastTime = new Date(lastRecord[0]);
          const diffHours = (timestamp - lastTime) / 3600000;

          if (diffHours < 10) {
            Logger.log(`Check-In denied: Only ${(10 - diffHours).toFixed(1)}h since last`);
            return jsonResponse(false, `Wait ${(10 - diffHours).toFixed(1)} hours to check-in again`);
          }

          const today = new Date().toDateString();
          if (lastTime.toDateString() === today) {
            Logger.log("Check-In denied: Already checked-in today");
            return jsonResponse(false, "Already checked-in today for this shift");
          }
        }
      } else { // Check-Out
        if (!lastRecord || lastRecord[4] !== "Check-In") {
          Logger.log("Check-Out denied: No check-in found");
          return jsonResponse(false, "No active check-in found");
        }
      }

      const offices = officesSheet.getDataRange().getValues().slice(1);
      let nearestOffice = null;
      let minDist = Infinity;

      offices.forEach(o => {
        const dist = getDistance(latitude, longitude, o[3], o[4]);
        if (dist < minDist) {
          minDist = dist;
          nearestOffice = o;
        }
      });

      if (!nearestOffice || minDist > 100) {
        Logger.log(`Too far from office: ${minDist.toFixed(1)}m`);
        return jsonResponse(false, `You are too far from office (${minDist.toFixed(0)} meters)`);
      }

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

      Logger.log(`${action} success at ${nearestOffice[1]}, distance ${minDist.toFixed(1)}m`);

      return jsonResponse(true, `${action} successful at ${nearestOffice[1]}`, {
        officeName: nearestOffice[1]
      });
    }

    // -------- 5️⃣ GET HISTORY --------
    if (action === "getHistory") {
      const employeeId = phone;
      const allRows = attendanceSheet.getDataRange().getValues().slice(1);
      const records = allRows.filter(r => r[1] === employeeId).map(r => ({
        timestamp: new Date(r[0]).toISOString(),
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

    // -------- 6️⃣ GET AGENTS BY OFFICE --------
    if (action === "getAgentsByOffice") {
      const officeID = params.officeID;
      if (!officeID) return jsonResponse(false, "OfficeID required");

      if ([officeIdCol, nameCol, phoneCol, roleCol].includes(-1)) {
        Logger.log("Missing columns in Employees sheet");
        return jsonResponse(false, "Required columns missing");
      }

      const agents = employeesRows
        .filter(r => r[officeIdCol] === officeID && r[roleCol] === "agent")
        .map(r => ({ name: r[nameCol], phone: r[phoneCol] }));

      const offices = officesSheet.getDataRange().getValues().slice(1);
      const office = offices.find(o => o[0] === officeID);
      const officeName = office ? office[1] : "";

      return jsonResponse(true, "", {
        agents,
        officeName
      });
    }

    // -------- 7️⃣ GET ALL EMPLOYEES --------
    if (action === "getAllEmployees") {
      const employees = employeesRows.map(r => ({
        name: r[nameCol],
        phone: r[phoneCol],
        role: r[roleCol],
        officeID: r[officeIdCol]
      }));
      return jsonResponse(true, "", { employees });
    }

    Logger.log(`Invalid action: ${action}`);
    return jsonResponse(false, "Invalid action");

  } catch (err) {
    Logger.log(`Unexpected error: ${err.message}`);
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: "Server error: " + err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// 🌍 Distance Calculation (Haversine formula)
function getDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = x => x * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLon / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}
