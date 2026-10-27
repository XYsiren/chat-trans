// app.js
let cloudEnv = '';
try {
  // 本地文件不提交到 Git；新环境可从 env.local.example.js 复制创建。
  cloudEnv = require('./env.local.js').cloudEnv || '';
} catch (error) {
  // 未配置时使用微信云开发控制台中设置的默认环境。
}
App({
  onLaunch: function () {
    this.globalData = {
      // env 参数说明：
      // env 参数决定接下来小程序发起的云开发调用（wx.cloud.xxx）会请求到哪个云环境的资源
      // 此处请填入环境 ID, 环境 ID 可在微信开发者工具右上顶部工具栏点击云开发按钮打开获取
      env: cloudEnv,
    };
    if (!wx.cloud) {
      console.error("请使用 2.2.3 或以上的基础库以使用云能力");
    } else {
      const cloudOptions = { traceUser: true };
      if (this.globalData.env) cloudOptions.env = this.globalData.env;
      wx.cloud.init(cloudOptions);
      // 仅记录匿名 OpenID 对应的一次启动事件，不采集文档正文或个人资料。
      require('./services/metrics-service').track('app_opened');
    }
  },
});
