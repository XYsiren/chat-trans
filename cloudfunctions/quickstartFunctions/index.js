const cloud = require("wx-server-sdk");
cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV,
});

const db = cloud.database();
// 获取openid
const getOpenId = async () => {
  // 获取基础信息
  const wxContext = cloud.getWXContext();
  return {
    openid: wxContext.OPENID,
    appid: wxContext.APPID,
    unionid: wxContext.UNIONID,
  };
};

// 获取小程序二维码
const getMiniProgramCode = async () => {
  // 获取小程序二维码的buffer
  const resp = await cloud.openapi.wxacode.get({
    path: "pages/index/index",
  });
  const { buffer } = resp;
  // 将图片上传云存储空间
  const upload = await cloud.uploadFile({
    cloudPath: "code.png",
    fileContent: buffer,
  });
  return upload.fileID;
};

// 创建集合
const createCollection = async () => {
  try {
    // 创建集合
    await db.createCollection("sales");
    await db.collection("sales").add({
      // data 字段表示需新增的 JSON 数据
      data: {
        region: "华东",
        city: "上海",
        sales: 11,
      },
    });
    await db.collection("sales").add({
      // data 字段表示需新增的 JSON 数据
      data: {
        region: "华东",
        city: "南京",
        sales: 11,
      },
    });
    await db.collection("sales").add({
      // data 字段表示需新增的 JSON 数据
      data: {
        region: "华南",
        city: "广州",
        sales: 22,
      },
    });
    await db.collection("sales").add({
      // data 字段表示需新增的 JSON 数据
      data: {
        region: "华南",
        city: "深圳",
        sales: 22,
      },
    });
    return {
      success: true,
    };
  } catch (e) {
    // 这里catch到的是该collection已经存在，从业务逻辑上来说是运行成功的，所以catch返回success给前端，避免工具在前端抛出异常
    return {
      success: true,
      data: "create collection success",
    };
  }
};

// 查询数据
const selectRecord = async () => {
  // 返回数据库查询结果
  return await db.collection("sales").get();
};

// 更新数据
const updateRecord = async (event) => {
  try {
    // 遍历修改数据库信息
    for (let i = 0; i < event.data.length; i++) {
      await db
        .collection("sales")
        .where({
          _id: event.data[i]._id,
        })
        .update({
          data: {
            sales: event.data[i].sales,
          },
        });
    }
    return {
      success: true,
      data: event.data,
    };
  } catch (e) {
    return {
      success: false,
      errMsg: e,
    };
  }
};

// 新增数据
const insertRecord = async (event) => {
  try {
    const insertRecord = event.data;
    // 插入数据
    await db.collection("sales").add({
      data: {
        region: insertRecord.region,
        city: insertRecord.city,
        sales: Number(insertRecord.sales),
      },
    });
    return {
      success: true,
      data: event.data,
    };
  } catch (e) {
    return {
      success: false,
      errMsg: e,
    };
  }
};

// 删除数据
const deleteRecord = async (event) => {
  try {
    await db
      .collection("sales")
      .where({
        _id: event.data._id,
      })
      .remove();
    return {
      success: true,
    };
  } catch (e) {
    return {
      success: false,
      errMsg: e,
    };
  }
};

// const getOpenId = require('./getOpenId/index');
// const getMiniProgramCode = require('./getMiniProgramCode/index');
// const createCollection = require('./createCollection/index');
// const selectRecord = require('./selectRecord/index');
// const updateRecord = require('./updateRecord/index');
// const fetchGoodsList = require('./fetchGoodsList/index');
// const genMpQrcode = require('./genMpQrcode/index');
// AI 文本转 Word 函数
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, UnorderedList, OrderedList, ListItem, Table, TableCell, TableRow, BorderStyle, convertInchesToTwip } = require("docx");

const generateWordFromBlocks = async (blocks) => {
  const paragraphs = [];

  for (const block of blocks) {
    if (block.type === "heading") {
      const headingLevels = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3];
      paragraphs.push(
        new Paragraph({
          text: block.content,
          heading: headingLevels[Math.min(block.level - 1, 2)],
          spacing: { after: 200 },
        })
      );
    } else if (block.type === "paragraph") {
      paragraphs.push(
        new Paragraph({
          text: block.content,
          spacing: { after: 200 },
        })
      );
    } else if (block.type === "quote") {
      paragraphs.push(
        new Paragraph({
          text: block.content,
          indent: { left: 720 },
          border: {
            left: {
              color: "cccccc",
              space: 1,
              style: BorderStyle.SINGLE,
              size: 12,
            },
          },
          spacing: { after: 200 },
        })
      );
    } else if (block.type === "code") {
      paragraphs.push(
        new Paragraph({
          text: "代码块",
          bold: true,
          spacing: { after: 100 },
        })
      );
      block.content.split("\n").forEach((line) => {
        paragraphs.push(
          new Paragraph({
            text: line || " ",
            font: "Courier New",
            size: 20,
            spacing: { after: 100 },
          })
        );
      });
      paragraphs.push(new Paragraph({ text: "", spacing: { after: 200 } }));
    } else if (block.type === "formula") {
      paragraphs.push(
        new Paragraph({
          text: "公式: " + block.content,
          italics: true,
          spacing: { after: 200 },
        })
      );
    } else if (block.type === "list") {
      const listItems = block.content.map((item) => new ListItem({ children: [new Paragraph(item)] }));
      paragraphs.push(
        new UnorderedList({
          children: listItems,
        })
      );
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: paragraphs,
      },
    ],
  });

  const blob = await Packer.toBuffer(doc);
  return blob;
};

const aiWordConversion = async (event) => {
  try {
    const { blocks } = event;
    if (!blocks || !Array.isArray(blocks)) {
      return { success: false, errMsg: "Invalid blocks format" };
    }

    const wordBuffer = await generateWordFromBlocks(blocks);

    const timestamp = Date.now();
    const cloudPath = `ai-word-conversion/${timestamp}.docx`;

    const uploadResult = await cloud.uploadFile({
      cloudPath: cloudPath,
      fileContent: wordBuffer,
    });

    return {
      success: true,
      fileID: uploadResult.fileID,
      fileName: `AI内容转换_${timestamp}.docx`,
      cloudPath: cloudPath,
    };
  } catch (error) {
    return {
      success: false,
      errMsg: error.message || "Word 生成失败",
    };
  }
};

// 云函数入口函数
exports.main = async (event, context) => {
  switch (event.type) {
    case "getOpenId":
      return await getOpenId();
    case "getMiniProgramCode":
      return await getMiniProgramCode();
    case "createCollection":
      return await createCollection();
    case "selectRecord":
      return await selectRecord();
    case "updateRecord":
      return await updateRecord(event);
    case "insertRecord":
      return await insertRecord(event);
    case "deleteRecord":
      return await deleteRecord(event);
    case "aiWordConversion":
      return await aiWordConversion(event);
  }
};
