# 上线后的用户与反馈运营

`metricsEvents` 中的 `app_opened` 是匿名启动事件。云开发控制台中按 `ownerOpenId` 去重即可得到累计用户；按日期筛选后去重即可得到 DAU，按最近 7/30 天筛选得到 WAU/MAU。客户端无该集合的读取权限。

反馈保存在 `feedback` 集合。仅云函数写入；后台在云开发控制台处理 `status`（`new`、`processing`、`resolved`）和 `adminNote`。反馈正文、联系方式和诊断信息只供运营人员查看。

部署前：创建 `metricsEvents` 与 `feedback` 集合，设置客户端读写权限为“仅管理员”，并部署 `metrics`、`feedback` 云函数。
