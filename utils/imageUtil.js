/**
 * 基于@rodewitsch/carbone库的图片处理改造
 * 1.新增了图片的宽高比自适应处理
 * 2.支持重复图片能多次展示不同的尺寸
 * 3.支持网络图片（url必须携带图片格式后缀）
 * 使用方式 {d.imgUrl:imageSize(100,80)}
 */
const sizeOf = require("image-size");
const axios = require("axios");

/**
 * 获取一个36位的随机字符串
 * @returns -随机字符串
 */
function uuid() {
  let d = new Date().getTime();
  const uuid = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(
    /[xy]/g,
    function (c) {
      const random_16 = Math.random() * 16;
      const r = (d + random_16) % 16 | 0;
      d = Math.floor(d / 16);
      return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
    }
  );
  return uuid;
}

/**
 * 将图片地址转换为base64
 * @param {*} urls
 */
function getBase64ImagesFromUrls(urls) {
  // 去重
  urls = [...new Set(urls)];
  const imgUrlReg =
    /https?:\/\/.*\.(?:png|jpg|jpeg|gif|bmp|svg|webp)(?:\?.*)?$/;
  const promiseArr = urls
    .filter((url) => imgUrlReg.test(url))
    .map((url) => {
      return axios({
        method: "get",
        url,
        responseType: "arraybuffer", // Important to get the image as a binary buffer
      }).then((response) => {
        // Convert the Buffer to a Base64 string
        const base64Image = Buffer.from(response.data, "base64");
        return {
          fileName: url,
          data: base64Image,
        };
      });
    });
  if (promiseArr.length == 0) return Promise.resolve([]);
  return Promise.all(promiseArr);
}

/**
 * 从data提取base64数据源，并修改data中对应字段为简单的标识，减少xml大小
 * @param {Object} data  - data to substitute
 * @returns {Array<{fileName: string, data: string}>}
 */
function getBase64ImagesFromData(data) {
  // 递归查找符合条件的字段
  let result = [];
  const base64Reg = /^data:image\/(png|jpeg|gif|bmp|svg\+xml|webp);base64,/;
  const getImage = (data) => {
    if (data instanceof Array) {
      data.forEach((item, index) => {
        if (typeof item === "string" && base64Reg.test(item)) {
          const randomKey = uuid();
          result.push({
            fileName: `${randomKey}`,
            data: Buffer.from(
              item.replace(/data:image\/(jpeg|png);base64,/, ""),
              "base64"
            ),
          });
          data[index] = `${randomKey}`;
        } else {
          getImage(item);
        }
      });
    } else if (data instanceof Object) {
      Object.keys(data).forEach(function (key) {
        if (typeof data[key] === "string" && base64Reg.test(data[key])) {
          const randomKey = uuid();
          result.push({
            fileName: `${randomKey}`,
            data: Buffer.from(
              data[key].replace(/data:image\/(jpeg|png);base64,/, ""),
              "base64"
            ),
          });
          data[key] = `${randomKey}`;
        } else {
          getImage(data[key]);
        }
      });
    }
  };
  getImage(data);
  return result;
}
/**
 * 替换图片
 * @param {Object} report - report data
 * @param {Array} base64Images - base64图片数据
 */
function putImagesToDocument(report, base64Images, callback) {
  const documentIndex = report.files.findIndex(
    (file) => file.name.indexOf("word/document.xml") != -1
  );
  const documentXml = report.files[documentIndex].data;
  // imageSize(url,val)
  const regex = new RegExp(`\\:imageSize\\(.*?,.*?\\)`, "g");
  const matches = documentXml.match(regex);
  const imgInfoList = [];
  matches &&
    matches.forEach((match) => {
      const regexItem = new RegExp(`\\:imageSize\\((.*?),.*?\\)`);
      imgInfoList.push(match.match(regexItem)[1]);
    });
  getBase64ImagesFromUrls(imgInfoList).then((result) => {
    let images = [...base64Images, ...result];
    if (!images.length) return callback();
    images = images.map((image) => {
      let encodedImageName = Buffer.from(image.fileName).toString("base64");
      report.files.push({
        data: image.data,
        isMarked: false,
        name: `word/media/${encodedImageName}.png`,
        parent: "",
      });
      return {
        ...image,
        encodedImageName,
      };
    });
    report.files[0].data = report.files[0].data.replace(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>'
    );
    let documentRelsIndex = report.files.findIndex(
      (file) => file.name.indexOf("document.xml.rels") != -1
    );
    let maxDocumentRelsId = Math.max(
      0,
      ...(report.files[documentRelsIndex].data.match(/Id=".*?"/g) || []).map(
        (elem) => +elem.substring(elem.indexOf("rId") + 3, elem.length - 1)
      )
    );
    images.forEach((image, index) => {
      let imageRid = maxDocumentRelsId + index + 1;
      report.files[documentRelsIndex].data = report.files[
        documentRelsIndex
      ].data.replace(
        "</Relationships>",
        `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${image.encodedImageName}.png"/></Relationships>`
      );
      const base64Name = Buffer.from(image.encodedImageName, "base64")
        .toString("utf8")
        .replace(/(\[|\]|\.)/g, "\\$1");
      let imageSizeFormatters =
        report.files[documentIndex].data.match(
          new RegExp(`<w:t>\\:imageSize\\(${base64Name},.*?\\)<\\/w:t>`, "g")
        ) || [];
      // 去重优化
      imageSizeFormatters = [...new Set(imageSizeFormatters)];
      imageSizeFormatters.forEach((formattersItem) => {
        const imageSizeFormatter = formattersItem.match(
          new RegExp(`<w:t>\\:imageSize\\(${base64Name},(.*?)\\)<\\/w:t>`)
        )[1];
        // 使用image-size库获取base64图片真实的宽度和高度
        const dimensions = sizeOf(image.data);
        const w_h = dimensions.width / dimensions.height;
        let imageWidth, imageHeight;
        if (imageSizeFormatter) {
          [imageWidth, imageHeight] = imageSizeFormatter.split("*");
          // 如果未设置宽度，就自动按比例计算
          if (imageHeight === undefined) {
            imageHeight = imageWidth / w_h;
          }
          // 如果宽高
          imageWidth *= 12700;
          imageHeight *= 12700;
        } else {
          imageWidth = dimensions.width * 12700;
          imageHeight = dimensions.height * 12700;
        }
        report.files[documentIndex].data = report.files[
          documentIndex
        ].data.replaceAll(
          formattersItem,
          `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${imageWidth}" cy="${imageHeight}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Рисунок 1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Аннотация 2019-04-04 111910.jpg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageWidth}" cy="${imageHeight}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
        );
      });
    });
    callback();
  });
}

module.exports = {
  getBase64ImagesFromData,
  putImagesToDocument,
};
