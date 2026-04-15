// pm2 配置。启动:  pm2 start ecosystem.config.cjs
// 常用:  pm2 logs eu-label-pdf | pm2 restart eu-label-pdf | pm2 monit
const path = require("path");

module.exports = {
  apps: [
    {
      name: "eu-label-pdf",
      cwd: __dirname,
      // 用绝对路径指向 venv 的 pythonw.exe,避免 pm2 走 PATH 拿到系统 Python
      script: path.join(__dirname, "venv", "Scripts", "pythonw.exe"),
      args: "server.py",
      interpreter: "none",
      autorestart: true,
      max_memory_restart: "500M",
      out_file: "./logs/out.log",
      error_file: "./logs/err.log",
      merge_logs: true,
      time: true,
      env: {
        PYTHONIOENCODING: "utf-8",
      },
    },
  ],
};
