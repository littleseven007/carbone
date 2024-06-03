/**
 * 文字相关处理工具  颜色、大小、字体、加粗
 * 支持{d.title:fontColor(.titleColor)}、{d.title:fontColor(#FF0000)}、:fontColor({d.titleColor},{d.title})
 */

const xml2js = require("xml2js");
const builder = new xml2js.Builder({
  headless: true,
  renderOpts: {
    pretty: false,
    indent: "",
    newline: "",
  },
});
const parser = new xml2js.Parser();

/**
 *  颜色、大小、字体、加粗局部替换
 * @param {Object} report - report data
 */
function putFontStyleToDocument(report) {
  const fontStyles = ["fontColor", "fontSize", "fontBold", "fontFamily"];
  const documentIndex = report.files.findIndex(
    (file) => file.name.indexOf("word/document.xml") != -1
  );
  const documentXml = report.files[documentIndex].data;
  // 匹配是否能找到fontColor(val)
  const regex = new RegExp(
    `\\:(?:${fontStyles.join("|")})\\(.*?\\).*?<\\/w:t>[\s\S]*?<\\/w:r>`,
    "g"
  );
  let item;
  const matches = [];

  while ((item = regex.exec(documentXml)) !== null) {
    matches.push({
      value: item[0], // 匹配的值
      index: item.index, // 匹配的索引
    });
  }
  // 计算分组，因为可能一个片段同时使用此功能
  let wrGroup = [];
  matches.forEach((match) => {
    const lastWRTag = documentXml.slice(0, match.index).lastIndexOf("<w:r>");
    let currentWR = documentXml.slice(lastWRTag, match.index) + match.value;
    wrGroup.push(currentWR);
  });
  // 去重
  wrGroup = [...new Set(wrGroup)];
  wrGroup.forEach((wr) => {
    const oldXml = wr;
    const stylesRegStr = `\\:((?:${fontStyles.join("|")}))\\((.*?)\\)`;
    const stylesReg = new RegExp(stylesRegStr, "g");
    const stylesParse = [];
    wr.match(stylesReg).forEach((result) => {
      const matchVal = result.match(new RegExp(stylesRegStr));
      // 如果是2个参数，则说明是特殊处理固定文字 的样式
      const valArr = matchVal[2].split(",");
      stylesParse.push({
        tag: matchVal[1],
        val: valArr[0],
      });
      // 去掉过滤标签
      wr = wr.replace(result, valArr[1] || "");
    });
    // 将XML字符串转换为JavaScript对象
    parser.parseString(wr, (err, result) => {
      if (err) {
        throw err;
      }

      let parentNode = result["w:r"]["w:rPr"][0];
      stylesParse.forEach((style) => {
        // 处理颜色 <w:color w:val= \"#FF0000\" />
        if (style.tag === "fontColor") {
          parentNode["w:color"] = [
            {
              $: {
                "w:val": style.val,
              },
            },
          ];
        }
        // 处理字体大小 <w:sz w:val= \"64\" /> <w:szCs w:val= \"64\" />
        // 相关资料：
        // <w:sz>与<w:szCs>中的数值是实际磅值的两倍（原文翻译为”此元素的数值代表的字体大小，表示半点值）。
        // sz表示Non-Complex Script Font Size，简单理解的话，就是单字节字符（如ASCII编码字符等）的大小
        // szCs表示Complex Script Font Size，可以简单理解为双字节字符（如中日韩文字、阿拉伯文等）的大小。
        if (style.tag === "fontSize") {
          parentNode = {
            ...parentNode,
            "w:sz": [
              {
                $: {
                  "w:val": style.val * 2,
                },
              },
            ],
            "w:szCs": [
              {
                $: {
                  "w:val": style.val * 2,
                },
              },
            ],
          };
        }
        // 处理字体加粗 <w:b /> <w:bCs />
        if (style.tag === "fontBold" && style.val !== "false") {
          parentNode = {
            ...parentNode,
            "w:b": [
              {
                $: {
                  "w:val": "true",
                },
              },
            ],
            "w:bCs": [
              {
                $: {
                  "w:val": "true",
                },
              },
            ],
          };
        }
        // 处理字体 <w:rFonts w:hint= \"eastAsia\" w:ascii= \"仿宋\" w:hAnsi= \"仿宋\" w:eastAsia= \"仿宋\" />
        // w:ascii 定义 ascii 码的字体类型；w:eastAsia 定义中文、日文等文字的字体类型; w:cs 定义复杂文字的字体类型，如阿拉伯文等；w:hAnsi 定义前三者以外的字体类型。
        if (style.tag === "fontFamily") {
          parentNode["w:rFonts"] = [
            {
              $: {
                "w:hint": "eastAsia",
                "w:ascii": style.val,
                "w:hAnsi": style.val,
                "w:eastAsia": style.val,
                "w:cs": style.val,
              },
            },
          ];
        }
      });
      result["w:r"]["w:rPr"][0] = parentNode;
      // 将JavaScript对象转换回XML字符串
      const newXml = builder.buildObject(result);
      // 替换原文档
      report.files[documentIndex].data = report.files[
        documentIndex
      ].data.replaceAll(oldXml, newXml);
    });
  });
}

/**
 * 字体、大小、加粗 整体替换
 * @param {*} report  - report data
 * @param {*} options 构建选项 目前支持fontOptions有fontFamily、fontSize、fontBold
 */
function putFontStyleFullToDocument(report, options) {
  if (!options.fontOptions) return;
  const fontStyles = Object.keys(options.fontOptions);
  if (fontStyles.length === 0) return;
  const documentIndex = report.files.findIndex(
    (file) => file.name.indexOf("word/document.xml") != -1
  );
  const documentXml = report.files[documentIndex].data;
  fontStyles.forEach((styleKey) => {
    // 处理字体 <w:rFonts w:hint= \"eastAsia\" w:ascii= \"仿宋\" w:hAnsi= \"仿宋\" w:eastAsia= \"仿宋\" />
    if (styleKey === "fontFamily") {
      const regex = /(w\:(?:ascii|hAnsi|eastAsia|cs)=\").*?(\")/g;
      const result = documentXml.replace(regex, function (match, p1, p2) {
        return p1 + options.fontOptions[styleKey] + p2;
      });
      report.files[documentIndex].data = result;
    }
    // 处理字体大小 <w:sz w:val= \"64\" /> <w:szCs w:val= \"64\" />
    // 匹配文字加粗修改<w:b w:val= \"0\" /> <w:bCs w:val= \"0\" />
    if (styleKey === "fontSize" || styleKey === "fontBold") {
      const regex = /(\<w\:(?:sz|szCs|b|bCs) w\:val=\").*?(\")/g;
      const result = documentXml.replace(regex, function (match, p1, p2) {
        if (!!p1.match("sz")) {
          return p1 + options.fontOptions["fontSize"] * 2 + p2;
        } else if (!!p1.match("b")) {
          return p1 + options.fontOptions["fontBold"] + p2;
        }
      });
      report.files[documentIndex].data = result;
    }
  });
}

/**
 * 对数据前面空格进行处理 替换为能展示的空格(仅当第一个字符为空格，并且只替换第一个空格，后面就能正常展示，不然存在误差)
 * @param data -需要处理的数据
 */
const spaceObject = (data) => {
  if (data instanceof Array) {
    data.forEach((item, index) => {
      if (typeof item === "string") {
        data[index] = data[index].replace(/^ /, "\u00A0\u00A0");
      } else {
        spaceObject(item);
      }
    });
  } else if (data instanceof Object) {
    Object.keys(data).forEach(function (key) {
      if (typeof data[key] === "string") {
        data[key] = data[key].replace(/^ /, "\u00A0\u00A0");
      } else {
        spaceObject(data[key]);
      }
    });
  }
};

/**
 * 目前没发现渲染纯数组的方式
 * 处理纯数组[1,2]为对象[{value:1},{value:2}]
 * @param data -需要处理的数据
 */
const arrayObject = (data) => {
  if (data instanceof Array) {
    if (data.length > 0) {
      if (data[0] instanceof Object) {
        data.forEach((item) => {
          arrayObject(item);
        });
      } else {
        data.forEach((item, index) => {
          data[index] = { value: item };
        });
      }
    }
  } else if (data instanceof Object) {
    Object.keys(data).forEach(function (key) {
      arrayObject(data[key]);
    });
  }
};

module.exports = {
  putFontStyleToDocument,
  putFontStyleFullToDocument,
  spaceObject,
  arrayObject,
};
