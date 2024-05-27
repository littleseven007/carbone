/**
 * 删除元素相关工具
 * 使用方式{d.imgUrl:ifEM():drop(p)}
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
 * 删除元素  目前支持p、tr、tc、tbl等
 * 其中tc最为复杂，需要考虑移除列后宽度均分给其他列的情况
 * @param {Object} report - report data
 */
function dropElementToDocument(report) {
  const documentIndex = report.files.findIndex(
    (file) => file.name.indexOf("word/document.xml") != -1
  );
  let documentXml = report.files[documentIndex].data;
  // 匹配是否能找到:drop(val)
  const regex = /\:drop\(.*?\)/;
  let item;
  let whileCount = 100; // 防止死循环
  while ((item = documentXml.match(regex)) !== null) {
    whileCount--;
    if (whileCount < 0) {
      console.log("死循环了");
      break;
    }
    const match = {
      value: item[0], // 匹配的值
      index: item.index, // 匹配的索引
    };
    let tagValue = match.value.match(/\((.*?)\)/)[1];
    const realTagValue = tagValue;
    // 如果是移除列，那么需要移除所有行中的列,并重新计算宽度
    if (realTagValue === "tc") {
      // 想了想只能操作当前所属整体table的xml最方便了
      tagValue = "tbl";
    }
    const leftTag = `<w:${tagValue}>`;
    const rightTag = `</w:${tagValue}>`;
    const firstTagIndex = documentXml
      .slice(0, match.index)
      .lastIndexOf(leftTag);
    const lastTagIndex =
      match.index + documentXml.slice(match.index).indexOf(rightTag);
    if (firstTagIndex === -1 || lastTagIndex === -1) return;
    const currentStr = documentXml.slice(
      firstTagIndex,
      lastTagIndex + rightTag.length
    );
    if (realTagValue === "tc") {
      const newCurrentStr = dropTCByIndex(currentStr);
      documentXml = documentXml.replace(currentStr, newCurrentStr);
    } else {
      documentXml = documentXml.replace(currentStr, "");
    }
  }
  report.files[documentIndex].data = documentXml;
}

/**
 * 删除表格列，并重新计算分配宽度
 * @param {*} tableXml
 * @param {*} index
 */
function dropTCByIndex(tableXml) {
  let newXml = ''
  // :drop(tc)可能存在多个，找到当前tc所处的index集合后一次性处理
  // 将XML字符串转换为JavaScript对象
  parser.parseString(tableXml, (err, result) => {
    if (err) {
      throw err;
    }
    const dropIndexs = [];
    // 移除tc
    result["w:tbl"]["w:tr"].forEach((tr) => {
      tr["w:tc"] = tr["w:tc"].filter((tc, index) => {
        const tcStr = builder.buildObject(tc); // 将JavaScript对象转换回XML字符串
        if (dropIndexs.includes(index)) return false
        if (tcStr.indexOf(":drop(tc)") != -1) {
          dropIndexs.push(index);
          return false;
        }
        return true;
      });
    });
    // 重新计算宽度
    let oldAllWidth = 0;
    let newAllWidth = 0;
    let gridCols = result["w:tbl"]["w:tblGrid"][0]["w:gridCol"];
    gridCols = gridCols.filter((gridCol, index) => {
      oldAllWidth += +gridCol["$"]["w:w"]
      if (!dropIndexs.includes(index)) {
        newAllWidth += +gridCol["$"]["w:w"]
        return true
      }
      return false
    })
    const rate = oldAllWidth / newAllWidth;
    gridCols.forEach((gridCol) => {
      gridCol["$"]["w:w"] = Math.floor(+gridCol["$"]["w:w"] * rate) + '';
    })
    result["w:tbl"]["w:tblGrid"][0]["w:gridCol"] = gridCols
    // 将JavaScript对象转换回XML字符串
    newXml = builder.buildObject(result);
  });
  return newXml;
}

module.exports = { dropElementToDocument };
