# LeetCode 每日提醒 + 自动划题

每天两次(GitHub Actions 自动运行):
- 早上:同步你在 LeetCode 做完的题 → 在 Excel 里划掉 → 手机推送"今天要做的 3 题"(带链接)
- 晚上:再同步一次 → 推送"今天做完了哪几题"或"今天还没做"的提醒

## 一次性设置(约 10 分钟)

1. 在 GitHub 新建一个**私有(Private)**仓库,把本文件夹里的所有文件上传(注意 `.github` 文件夹也要上传)。
2. 手机装 **ntfy** App(iOS / Android),订阅一个只有你知道的随机主题名,例如 `chenxi-lc-7f3k29`。
3. 仓库 → Settings → Secrets and variables → Actions:
   - Secrets → New repository secret:`NTFY_TOPIC` = 上一步的主题名
   - Secrets → `LEETCODE_SESSION` = 浏览器里 leetcode.com 的 Cookie 值
     (Chrome:登录 LeetCode → F12 → Application → Cookies → leetcode.com → 复制 `LEETCODE_SESSION` 的值)
   - Variables → New repository variable:`LEETCODE_USERNAME` = 你的 LeetCode 用户名(后备方式,主页需公开)
4. Actions 页面 → leetcode-daily → Run workflow,手动跑一次,确认手机收到消息。

## 注意
- `LEETCODE_SESSION` 等于你的登录凭证,只放进 Secret,不要写进文件或发给别人;它会过期(通常几周),过期后消息里会出现 ⚠ 提示,重新复制一次即可。登录态失效时会自动改用公开主页(只能看到最近 20 条通过记录,够每天用)。
- 定时任务按 UTC 运行,冬令时会晚/早一小时;GitHub 的定时任务可能延迟几分钟到半小时。
- 修改 Excel(填用时、掌握度等)请直接在 GitHub 网页上传新版本,或先 `git pull` 再改;不要和机器人同时改,Excel 是二进制文件,冲突无法自动合并。
- 如果 LeetCode 拦截了 GitHub 的服务器(返回 403),消息会提示同步失败,这时本地运行也可以:`python sync_leetcode.py --xlsx LeetCode_plan.xlsx --paste solved.txt`。
