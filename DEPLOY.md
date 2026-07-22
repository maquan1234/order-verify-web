# 部署指南：把审单核对系统部署到云端（永久 URL）

本目录已改造为「云原生」结构，可直接部署到 Railway 或 Render，拿到 **7×24 在线、永久不变** 的公网地址，彻底摆脱本地 SSH 隧道频繁掉线的问题。

---

## 一、改了什么（为什么现在能上云）

| 项目 | 改造前 | 改造后 |
|------|--------|--------|
| 数据路径 | 写死 `C:\Users\EDY\...` | 用环境变量 `DATA_DIR`，默认当前目录 |
| 端口 | 写死 5000 | 读环境变量 `PORT`（云平台自动注入） |
| 货品数据 | 依赖桌面 xlsx | 首次运行自动从仓库内 `products.json` 种子化到持久卷 |
| 部署配置 | 无 | 已备好 `requirements.txt` / `Procfile` / `railway.json` / `render.yaml` / `runtime.txt` |

前端（`templates/index.html`）本来就全部用相对路径 `/api/...`，无需改动即可上云。

---

## 二、部署前准备（一次性）

1. 注册一个 **GitHub** 账号（免费）：https://github.com
2. 注册一个云平台账号（二选一，都免费额度）：
   - **Railway**：https://railway.app
   - **Render**：https://render.com
3. 把本目录（webapp）推送到一个 GitHub 仓库：

```bash
cd C:\Users\EDY\WorkBuddy\2026-07-21-16-17-50\webapp
git init
git add .
git commit -m "电商审单核对系统 - 云原生版本"
git branch -M main
git remote add origin https://github.com/你的用户名/order-verify.git
git push -u origin main
```

> 注：`products.json`（你现有的 82 条货品）已包含在目录里，会一起推上去，云端首次运行自动加载，不会丢数据。

---

## 三、方案 A：Railway 部署（推荐，最简单）

1. 打开 https://railway.app → 用 GitHub 登录
2. 点 **"New Project"** → **"Deploy from GitHub repo"**
3. 选择刚才推送的仓库，确认部署
4. 部署完成后，进入项目 → **"Variables"**，添加一个变量：
   - `DATA_DIR` = `/data`
5. 进入项目 → **"Volumes"** → 添加一个卷，挂载路径填 `/data`（容量 1GB 足够）
6. 回到 **"Settings"** → **"Domains"**，点击生成域名，得到类似 `order-verify.up.railway.app` 的**永久地址**
7. 打开该地址即可使用，任何网络、任何电脑都能访问

> 以后改了代码，只要 `git push`，Railway 会自动重新部署。

---

## 四、方案 B：Render 部署（备选）

1. 打开 https://render.com → 用 GitHub 登录
2. 点 **"New"** → **"Web Service"** → 连接 GitHub 仓库
3. 配置：
   - Runtime: **Python**
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python app.py`
   - Plan: **Free**
4. 展开 **"Advanced"** → **"Environment Variables"**，添加：
   - `DATA_DIR` = `/data`
5. 在 **"Disks"** 里添加一个磁盘，挂载路径 `/data`
6. 创建后，Render 会给一个类似 `order-verify.onrender.com` 的**永久地址**（Free 版首次访问有约 50 秒唤醒延迟，之后正常）

---

## 五、货品数据如何持久化

- 首次部署：云端持久卷为空 → 自动从仓库内 `products.json` 复制 82 条货品到卷里
- 之后你在网页里「新增 / 修改 / 删除」货品：写入都在持久卷 `/data/products.json`，**重启不丢**
- 想换一批初始货品：在网页用「Excel 导入」功能即可，无需动代码

---

## 六、本地临时使用（过渡期）

部署到云端前，本地仍可用。本目录的 `启动服务.bat` 双击启动（含本地公网隧道），或：

```bash
python app.py          # 默认 http://127.0.0.1:5000
```

---

## 七、常见问题

**Q: 部署后打开是空白 / 500 错误？**
A: 检查是否添加了 `DATA_DIR=/data` 变量，以及是否挂载了对应路径的卷。

**Q: 上传大文件失败？**
A: 当前限制 50MB，可在 `app.py` 调整 `MAX_CONTENT_LENGTH`。

**Q: 想绑定自己的域名？**
A: Railway / Render 都支持自定义域名，在 Domains 设置里按提示加 CNAME 解析即可。
