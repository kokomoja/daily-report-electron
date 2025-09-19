window.addEventListener("DOMContentLoaded", () => {
  const loginBtn = document.getElementById("loginBtn");
  const addUserBtn = document.getElementById("addUserBtn");

  loginBtn.addEventListener("click", async () => {
    const user = document.getElementById("login_user").value;
    const pass = document.getElementById("login_password").value;

    if (!user || !pass) {
      document.getElementById("status").textContent =
        "⚠️ กรุณากรอกชื่อผู้ใช้และรหัสผ่าน";
      return;
    }

    const res = await window.auth.login({ user, pass });

    if (res.success) {
      localStorage.setItem("login_user", res.login_user);
      localStorage.setItem("username", res.username);
      window.location = "index.html";
    } else {
      document.getElementById("status").textContent =
        "❌ " + (res.message || res.error);
    }
  });

  addUserBtn.addEventListener("click", async () => {
    const new_user = document.getElementById("new_user").value;
    const new_pass = document.getElementById("new_pass").value;
    const new_name = document.getElementById("new_name").value;
    const secret = document.getElementById("secret").value;

    if (!new_user || !new_pass || !new_name || !secret) {
      document.getElementById("status").textContent =
        "⚠️ กรุณากรอกข้อมูลให้ครบ";
      document.getElementById("status").style.color = "red";
      return;
    }

    const res = await window.auth.addUser({
      login_user: new_user,
      login_password: new_pass,
      username: new_name,
      secret,
    });

    document.getElementById("status").textContent = res.message || res.error;
    document.getElementById("status").style.color = res.success
      ? "green"
      : "red";

    if (res.success) {
      document.getElementById("new_user").value = "";
      document.getElementById("new_pass").value = "";
      document.getElementById("new_name").value = "";
      document.getElementById("secret").value = "";
    }
  });
});
