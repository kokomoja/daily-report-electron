// main.js
const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const sql = require("mssql");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");
const fs = require("fs");

// Config SQL Server
const config = {
  server: "192.168.7.110",
  database: "opDb",
  user: "sa",
  password: "123456",
  options: {
    encrypt: false,
    trustServerCertificate: true,
    instanceName: "sqlexpress",
  },
};

app.commandLine.appendSwitch("disable-gpu-shader-disk-cache");

let mainWindow;

function createWindow(file) {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  mainWindow.loadFile(path.join(__dirname, file));
}

app.on("ready", () => {
  createWindow("login.html"); // ✅ เริ่มที่หน้า login
});

function toLocalDateTime(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d;
}

function formatDateUTC(dateValue) {
  if (!dateValue) return "";
  const d = new Date(dateValue);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function formatDateTimeUTC(dateValue) {
  if (!dateValue) return "";
  const d = new Date(dateValue);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const hh = String(d.getUTCHours()).padStart(2, "0");
  const min = String(d.getUTCMinutes()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
}

function formatHourMinute(totalMinutes) {
  if (totalMinutes == null) return "";
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${hours.toString().padStart(2, "0")}:${minutes.toString().padStart(2, "0")}`;
}

function getFontPath(fontFile) {
  if (fs.existsSync(path.join(__dirname, "fonts", fontFile))) {
    return path.join(__dirname, "fonts", fontFile); // dev mode
  }
  return path.join(process.resourcesPath, "fonts", fontFile); // build exe
}

// ====================== HANDLERS ======================
ipcMain.handle("login-check", async (event, data) => {
  try {
    const pool = await sql.connect(config);
    const result = await pool
      .request()
      .input("user", sql.VarChar(50), data.user)
      .input("pass", sql.VarChar(50), data.pass)
      .query(
        "SELECT login_user, username FROM loginID WHERE login_user=@user AND login_password=@pass"
      );

    if (result.recordset.length > 0) {
      const user = result.recordset[0];
      return {
        success: true,
        message: "เข้าสู่ระบบสำเร็จ",
        login_user: user.login_user,
        username: user.username,
      };
    } else {
      return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle("add-user", async (event, data) => {
  try {
    // ✅ ตรวจสอบรหัสลับก่อน
    if (data.secret !== "0845535000721") {
      return { success: false, message: "❌ รหัสลับไม่ถูกต้อง" };
    }

    const pool = await sql.connect(config);
    await pool
      .request()
      .input("login_user", sql.VarChar(50), data.login_user)
      .input("login_password", sql.VarChar(50), data.login_password)
      .input("username", sql.VarChar(50), data.username).query(`
    INSERT INTO loginID (login_user, login_password, username)
    VALUES (@login_user, @login_password, @username)
  `);

    return { success: true, message: "✅ เพิ่มผู้ใช้งานสำเร็จ" };
  } catch (err) {
    console.error("add-user error:", err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle("get-lists", async () => {
  try {
    let pool = await sql.connect(config);
    const machines = await pool
      .request()
      .query("SELECT machine FROM machineNumber");
    const operators = await pool
      .request()
      .query("SELECT op_name FROM operatorName");
    return {
      machines: machines.recordset.map((m) => m.machine),
      operators: operators.recordset.map((o) => o.op_name),
    };
  } catch (err) {
    console.error("get-lists error:", err);
    return { error: err.message };
  }
});

ipcMain.handle("insert-report", async (event, data) => {
  try {
    let pool = await sql.connect(config);
    await pool
      .request()
      .input("op_date", sql.Date, data.op_date)
      .input("machine", sql.VarChar(50), data.machine)
      .input("operator", sql.VarChar(50), data.operator)
      .input("job", sql.VarChar(100), data.job)
      .input("start_time", sql.DateTime, toLocalDateTime(data.start_time))
      .input("stop_time", sql.DateTime, toLocalDateTime(data.stop_time))
      .input("recorder_name", sql.VarChar(50), data.recorder_name) // ✅ เพิ่ม
      .query(`
        INSERT INTO dailyReport 
        (op_date, machine, operator, job, start_time, stop_time, op_hour, recorder_name, time_stamp)
        VALUES (@op_date, @machine, @operator, @job, @start_time, @stop_time,
          DATEDIFF(MINUTE, @start_time, @stop_time), @recorder_name, GETDATE())
      `);

    return { success: true, message: "บันทึกเรียบร้อย" };
  } catch (err) {
    console.error("insert error:", err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle("update-report", async (event, data) => {
  try {
    let pool = await sql.connect(config);
    await pool
      .request()
      .input("op_id", sql.Int, data.op_id)
      .input("op_date", sql.Date, data.op_date)
      .input("machine", sql.VarChar(50), data.machine)
      .input("operator", sql.VarChar(50), data.operator)
      .input("job", sql.VarChar(100), data.job)
      .input("start_time", sql.DateTime, toLocalDateTime(data.start_time))
      .input("stop_time", sql.DateTime, toLocalDateTime(data.stop_time))
      .input("recorder_name", sql.VarChar(50), data.recorder_name) // ✅ เพิ่ม
      .query(`
        UPDATE dailyReport
        SET op_date=@op_date, machine=@machine, operator=@operator, job=@job,
            start_time=@start_time, stop_time=@stop_time,
            op_hour = DATEDIFF(MINUTE, @start_time, @stop_time),
            recorder_name=@recorder_name,
            time_stamp=GETDATE()
        WHERE op_id=@op_id
      `);

    return { success: true, message: "อัพเดทเรียบร้อย" };
  } catch (err) {
    console.error("update error:", err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle("get-reports", async (event, date) => {
  try {
    let pool = await sql.connect(config);
    let query = `
      SELECT op_id, op_date, machine, operator, job, start_time, stop_time, 
             op_hour, time_stamp, recorder_name 
      FROM dailyReport
    `;
    let request = pool.request();
    if (date) {
      query += " WHERE CAST(op_date AS DATE) = @date";
      request = request.input("date", sql.Date, date);
    }
    query += " ORDER BY op_date DESC";
    let result = await request.query(query);
    return result.recordset;
  } catch (err) {
    console.error("get-reports error:", err);
    return { error: err.message };
  }
});

ipcMain.handle("delete-multiple", async (event, ids) => {
  try {
    if (!ids || ids.length === 0) {
      return { success: false, error: "ไม่ได้เลือกข้อมูลที่จะลบ" };
    }
    let pool = await sql.connect(config);
    await pool
      .request()
      .query(`DELETE FROM dailyReport WHERE op_id IN (${ids.join(",")})`);
    return { success: true, message: "ลบข้อมูลเรียบร้อย" };
  } catch (err) {
    console.error("delete-multiple error:", err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle("export-excel", async (event, date) => {
  try {
    let pool = await sql.connect(config);
    let query = "SELECT * FROM dailyReport";
    let request = pool.request();
    if (date) {
      query += " WHERE CAST(op_date AS DATE) = @date";
      request = request.input("date", sql.Date, date);
    }
    let result = await request.query(query);
    const reports = result.recordset;

    if (!reports || reports.length === 0) {
      return { success: false, error: "ไม่มีข้อมูล" };
    }

    const data = reports.map((r) => ({
      ลำดับ: r.op_id,
      วันที่: formatDateUTC(r.op_date),
      เครื่องจักร: r.machine,
      พนักงาน: r.operator,
      งาน: r.job,
      เริ่มต้น: formatDateTimeUTC(r.start_time),
      สิ้นสุด: formatDateTimeUTC(r.stop_time),
      เวลาทำงาน: formatHourMinute(r.op_hour),
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reports");

    const exportPath = path.join(
      app.getPath("documents"),
      `daily_report_${Date.now()}.xlsx`
    );
    XLSX.writeFile(wb, exportPath);

    return { success: true, message: "Export สำเร็จ", path: exportPath };
  } catch (err) {
    console.error("export-excel error:", err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle("export-pdf", async (event, date) => {
  try {
    let pool = await sql.connect(config);
    let query = "SELECT * FROM dailyReport";
    let request = pool.request();
    if (date) {
      query += " WHERE CAST(op_date AS DATE) = @date";
      request = request.input("date", sql.Date, date);
    }
    query += " ORDER BY op_date DESC";
    let result = await request.query(query);
    const reports = result.recordset;

    if (!reports || reports.length === 0) {
      return { success: false, error: "ไม่มีข้อมูล" };
    }

    const exportPath = path.join(
      app.getPath("documents"),
      `daily_report_${Date.now()}.pdf`
    );
    const doc = new PDFDocument({ margin: 30 });
    doc.pipe(fs.createWriteStream(exportPath));

    doc.registerFont(
      "THSarabunNew",
      path.join(__dirname, "fonts", "THSarabunNew.ttf")
    );
    doc.registerFont(
      "THSarabunNewBold",
      path.join(__dirname, "fonts", "THSarabunNew Bold.ttf")
    );

    doc
      .font("THSarabunNewBold")
      .fontSize(26)
      .text("บริษัท พี.ซี.ปิโตรเลียมแอนด์เทอร์มินอล จำกัด", {
        align: "center",
      });
    doc.moveDown(0.3);
    doc
      .font("THSarabunNew")
      .fontSize(16)
      .text("แผนการปฏิบัติงานแผนกเทกอง / เครื่องมือหนัก", { align: "center" });
    doc.moveDown(0.2);
    const reportDate = new Date(reports[0].op_date);
    const thaiMonths = [
      "มกราคม",
      "กุมภาพันธ์",
      "มีนาคม",
      "เมษายน",
      "พฤษภาคม",
      "มิถุนายน",
      "กรกฎาคม",
      "สิงหาคม",
      "กันยายน",
      "ตุลาคม",
      "พฤศจิกายน",
      "ธันวาคม",
    ];
    const formattedDate = `${reportDate.getDate()} ${thaiMonths[reportDate.getMonth()]} ${reportDate.getFullYear() + 543}`;
    doc
      .font("THSarabunNew")
      .fontSize(16)
      .text(`วันที่ ${formattedDate}`, { align: "center" });
    doc.moveDown(1);

    const headers = [
      "ลำดับ",
      "วันที่",
      "เครื่องจักร",
      "พนักงาน",
      "งาน",
      "เริ่มต้น",
      "สิ้นสุด",
      "เวลาทำงาน",
    ];
    const columnWidths = [40, 65, 60, 150, 200, 100, 100, 60];
    let startX = 30;
    let startY = doc.y + 10;

    headers.forEach((h, i) => {
      doc.rect(startX, startY, columnWidths[i], 30).stroke();
      doc.text(h, startX + 5, startY + 8, {
        width: columnWidths[i] - 10,
        align: "center",
      });
      startX += columnWidths[i];
    });

    let rowY = startY + 30;
    reports.forEach((r, idx) => {
      startX = 30;
      const row = [
        idx + 1,
        formatDateUTC(r.op_date),
        r.machine || "",
        r.operator || "",
        r.job || "",
        formatDateTimeUTC(r.start_time),
        formatDateTimeUTC(r.stop_time),
        formatHourMinute(r.op_hour),
      ];

      row.forEach((cell, i) => {
        doc.rect(startX, rowY, columnWidths[i], 25).stroke();
        doc.text(cell.toString(), startX + 3, rowY + 7, {
          width: columnWidths[i] - 6,
          align: "center",
        });
        startX += columnWidths[i];
      });
      rowY += 25;

      if (rowY > 550) {
        doc.addPage({ size: "A4", layout: "landscape" });
        rowY = 50;
      }
    });

    doc.end();

    return { success: true, message: "Export PDF สำเร็จ", path: exportPath };
  } catch (err) {
    console.error("export-pdf error:", err);
    return { success: false, error: err.message };
  }
});
