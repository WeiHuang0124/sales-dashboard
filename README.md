# 銷售儀表板

蝦皮銷售資料儀表板，支援拖曳上傳 .xlsx / .csv 檔案。

## 本機開發

```bash
npm install
npm run dev
```

瀏覽器打開 http://localhost:5173

## 部署到 Vercel（免費）

### 第一次部署
1. 去 [github.com](https://github.com) 建立新 repository（例如 `sales-dashboard`）
2. 在這個資料夾執行：
   ```bash
   git init
   git add .
   git commit -m "init"
   git remote add origin https://github.com/你的帳號/sales-dashboard.git
   git push -u origin main
   ```
3. 去 [vercel.com](https://vercel.com) → New Project → 選這個 GitHub repo → Deploy

### 之後每次更新
修改程式碼後：
```bash
git add .
git commit -m "更新內容"
git push
```
Vercel 會自動重新部署，幾秒後生效。

## 功能
- 拖曳 .xlsx / .csv 到畫面任意位置即可匯入
- 自動辨識蝦皮店鋪統計格式
- 每日、每月營業額圖表
- 商品銷售排名
- 資料儲存在瀏覽器 localStorage（不需要後端）
