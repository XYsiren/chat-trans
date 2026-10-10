// index.js
Page({
  data: {
    powerList: [
      {
        title: "AI 内容转 Word",
        tip: "粘贴 AI 文本和公式代码，自动识别结构并导出 Word",
        page: "ai-word",
        showItem: false,
        item: []
      },
    ],
  },

  onClickPowerInfo(e) {
    const index = e.currentTarget.dataset.index;
    const powerList = this.data.powerList;
    const selectedItem = powerList[index];

    if (selectedItem.page) {
      wx.navigateTo({
        url: `/pages/${selectedItem.page}/index`,
      });
    } else {
      selectedItem.showItem = !selectedItem.showItem;
      this.setData({
        powerList,
      });
    }
  },
});
