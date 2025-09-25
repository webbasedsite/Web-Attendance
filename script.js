function doPost(e) {
  const params = e.parameter;
  const action = params.action;

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
  const phone = params.phone?.trim();
  const password = params.password?.trim();

  const employees = employeesSheet.getDataRange().getValues();
  const headers = employees[0];
  const rows = employees.slice(1);

  const phoneCol = headers.indexOf("Phone");
  const passwordCol = headers.indexOf("Password");
  const roleCol = headers.indexOf("Role");

  Logger.log("Headers: " + headers);
  Logger.log("Phone Col: " + phoneCol + ", Password Col: " + passwordCol + ", Role Col: " + roleCol);
  Logger.log("Attempting login with phone: '" + phone + "', password: '" + password + "'");

  if (phoneCol === -1 || passwordCol === -1 || roleCol === -1) {
    return jsonResponse(false, "Missing required columns in Employees sheet");
  }

  for (let i = 0; i < rows.length; i++) {
    Logger.log(`Row ${i}: Phone='${rows[i][phoneCol]}', Password='${rows[i][passwordCol]}'`);
  }

  const matched = rows.find(r =>
    String(r[phoneCol]).trim() === phone &&
    String(r[passwordCol]).trim() === password
  );

  if (matched) {
    Logger.log("Login success for phone: " + phone);
    return jsonResponse(true, "Login success", {
      role: matched[roleCol]
    });
  }

  Logger.log("Login failed: No matching phone/password");
  return jsonResponse(false, "Invalid phone or password");
}

  // ----------------------
  // 2️⃣ Get Offices
  // ----------------------
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

  // ----------------------
  // 3️⃣ Get Office Location
  // ----------------------
  if (action === "getOfficeLocation") {
    const phone = params.phone;
    const employees = employeesSheet.getDataRange().getValues().slice(1);
    const emp = employees.find(r => r[0] === phone || r[2] === phone);
    if (!emp) return jsonResponse(false, "Employee not found");
    const officeID = emp[4];
    const offices = officesSheet.getDataRange().getValues().slice(1);
    const office = offices.find(r => r[0] === officeID);
    if (!office) return jsonResponse(false, "Office not found");
    return jsonResponse(true, "", {
      latitude: office[3],
      longitude: office[4]
    });
  }

  // ----------------------
  // 4️⃣ Check-In / Check-Out
  // ----------------------
  if (action === "Check-In" || action === "Check-Out") {
    const employeeId = params.phone;
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
        if (diffHours < 10)
          return jsonResponse(false, `Cannot check-in yet, wait ${(10 - diffHours).toFixed(1)} hours`);
        const today = new Date().toDateString();
        if (lastTime.toDateString() === today)
          return jsonResponse(false, "Already checked-in today for this shift");
      }
    } else {
      if (!lastRecord || lastRecord[4] !== "Check-In")
        return jsonResponse(false, "No active check-in found");
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
    return jsonResponse(true, `${action} successful`);
  }

  // ----------------------
  // 5️⃣ Get History
  // ----------------------
  if (action === "getHistory") {
    const employeeId = params.phone;
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

  // If no valid action
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
