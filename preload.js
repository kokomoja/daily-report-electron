// v2.7 preload setup
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  getLists: () => ipcRenderer.invoke("get-lists"),
  saveReport: (data) => ipcRenderer.invoke("insert-report", data),
  updateReport: (data) => ipcRenderer.invoke("update-report", data),
  getReports: (date) => ipcRenderer.invoke("get-reports", date),
  deleteReport: (id) => ipcRenderer.invoke("delete-report", id),
  deleteMultiple: (ids) => ipcRenderer.invoke("delete-multiple", ids),
  exportExcel: (date) => ipcRenderer.invoke("export-excel", date),
  exportPDF: (date) => ipcRenderer.invoke("export-pdf", date), // ✅ เพิ่ม PDF
});

contextBridge.exposeInMainWorld("auth", {
  login: (data) => ipcRenderer.invoke("login-check", data),
  addUser: (data) => ipcRenderer.invoke("add-user", data),
});

contextBridge.exposeInMainWorld("opre", {
  insert: (data) => ipcRenderer.invoke("insert-opre", data),
  getCodes: () => ipcRenderer.invoke("get-opre-codes"),
  getList: () => ipcRenderer.invoke("get-opre-list"),
});

contextBridge.exposeInMainWorld("excel", {
  exportTable: (reports, filename) => {
    // ✅ ใช้ formatDate / formatDateTime
    const data = reports.map((r) => ({
      ลำดับ: r.op_id,
      วันที่: formatDate(r.op_date),
      เครื่องจักร: r.machine,
      พนักงาน: r.operator,
      งาน: r.job,
      เริ่มต้น: formatDateTime(r.start_time),
      สิ้นสุด: formatDateTime(r.stop_time),
      เวลาทำงาน: formatHourMinute(r.op_hour),
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reports");
    XLSX.writeFile(wb, filename || "daily_report.xlsx");
  },
});
