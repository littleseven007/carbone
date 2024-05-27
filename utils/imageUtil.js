/**
 * 基于@rodewitsch/carbone库的图片处理改造，代码基本完整保留，
 * 1.新增了图片的宽高比自适应处理
 * 2.支持重复图片能多次展示不同的尺寸
 * 使用方式 {d.imgUrl:imageSize(100,80)}
 */
const sizeOf = require("image-size");
/**
 * Add file buffer to image list
 * @param {Object} data  - data to substitute
 * @returns {Array<{fileName: string, data: string}>}
 */
function getImagesFromData(data) {
  let acc = [];
  for (let key in data) {
    if (!data[key]) continue;
    if (Array.isArray(data[key])) {
      for (let i = 0; i < data[key].length; i++) {
        for (let nestedKey in data[key][i]) {
          if (
            (data[key][i][nestedKey] &&
              data[key][i][nestedKey].indexOf &&
              data[key][i][nestedKey].indexOf("base64") != -1) ||
            (data[key][i][nestedKey] &&
              (Buffer.isBuffer(data[key][i][nestedKey]) ||
                data[key][i][nestedKey].type == "Buffer" ||
                data[key][i][nestedKey].BYTES_PER_ELEMENT != undefined))
          ) {
            if (
              data[key][i][nestedKey].type == "Buffer" ||
              data[key][i][nestedKey].BYTES_PER_ELEMENT != undefined
            )
              data[key][i][nestedKey] = Buffer.from(data[key][i][nestedKey]);
            acc.push({
              fileName: `${key}[${i}].${nestedKey}`,
              data: Buffer.from(
                data[key][i][nestedKey]
                  .toString("utf8")
                  .replace(/data:image\/(jpeg|png);base64,/, ""),
                "base64"
              ),
            });
            data[key][i][nestedKey] = `${key}[${i}].${nestedKey}`;
          }
        }
      }
    } else {
      if (data[key].indexOf && data[key].indexOf("base64") != -1) {
        acc.push({
          fileName: `${key}`,
          data: Buffer.from(
            data[key].replace(/data:image\/(jpeg|png);base64,/, ""),
            "base64"
          ),
        });
        data[key] = `${key}`;
      }
    }
  }
  return acc;
}
/**
 * Insert images into document
 * @param {Object} report - report data
 * @param {Array} images - images
 */
function putImagesToDocument(report, images = []) {
  if (!images.length) return;
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
  let documentIndex = report.files.findIndex(
    (file) => file.name.indexOf("word/document.xml") != -1
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
        new RegExp(`<w:t>${base64Name}:(\\d+\\*\\d+|\\d+)<\\/w:t>`, "g")
      ) || [];
    // 去重优化
    imageSizeFormatters = [...new Set(imageSizeFormatters)]; 
    imageSizeFormatters.forEach((formattersItem) => {
      const imageSizeFormatter = formattersItem.match(
        new RegExp(`<w:t>${base64Name}:(\\d+\\*\\d+|\\d+)<\\/w:t>`)
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
      ].data.replace(
        new RegExp(formattersItem, "g"),
        `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${imageWidth}" cy="${imageHeight}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Рисунок 1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Аннотация 2019-04-04 111910.jpg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageWidth}" cy="${imageHeight}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
      );
    });
  });
}
/**
 * Insert images into sheets
 * @param {Object} report - report data
 * @param {Array} images - images
 */
function putImagesToSheets(report, images = []) {
  if (!images.length) return;
  images = images.map((image) => {
    let encodedImageName = Buffer.from(image.fileName).toString("base64");
    report.files.push({
      data: image.data,
      isMarked: false,
      name: `xl/media/${encodedImageName}.png`,
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
  let sheets = [];
  report.files.forEach((file, index) => {
    if (file.name.indexOf("xl/worksheets/sheet") != -1)
      sheets.push({
        file,
        sheetNumber: file.name.substring(
          file.name.indexOf("/sheet") + 6,
          file.name.lastIndexOf(".")
        ),
        index,
      });
  });
  sheets.forEach((sheet) => {
    let sheetHasImages = images.reduce((acc, image) => {
      if (sheet.file.data.indexOf(image.fileName) != -1) acc++;
      return acc;
    }, 0);
    if (!sheetHasImages) return;
    let sheetDrawingRelsIndex = 0;
    let drawingId = 0;
    let sheetRelsIndex = report.files.findIndex(
      (file) => file.name.indexOf(`sheet${sheet.sheetNumber}.xml.rels`) != -1
    );
    // нет sheetRels у листа
    if (sheetRelsIndex == -1) {
      // ищем все drawings
      let allDrawings = report.files.filter(
        (file) => file.name.indexOf("drawings/drawing") != -1
      );
      // берем максимальное название
      if (allDrawings.length) {
        drawingId = Math.max(
          0,
          ...allDrawings.map(
            (file) =>
              +file.name.substring(
                file.name.indexOf("xl/drawings/drawing") + 19,
                file.name.lastIndexOf(".")
              )
          )
        );
        drawingId++;
      } else {
        drawingId = 1;
      }
      // создаем новый файл с макс названием +1
      report.files.push(createBlankDrawingRelsFile(drawingId));
      sheetDrawingRelsIndex = report.files.length - 1;
      // добавляем sheetRels с указателем на drawings
      report.files.push(createBlankSheetRelsFile(sheet.sheetNumber));
      sheetRelsIndex = report.files.length - 1;
      let maxSheetRelsId = Math.max(
        0,
        ...(report.files[sheetRelsIndex].data.match(/Id=".*?"/g) || []).map(
          (elem) => +elem.substring(elem.indexOf("rId") + 3, elem.length - 1)
        )
      );
      report.files[sheetRelsIndex].data = report.files[
        sheetRelsIndex
      ].data.replace(
        "</Relationships>",
        `<Relationship Id="rId${
          maxSheetRelsId + 1
        }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingId}.xml"/></Relationships>`
      );
      // создание drawingRels
      let sheetDrawingIndex = 0;
      report.files.push(createBlankDrawingFile(drawingId));
      sheetDrawingIndex = report.files.length - 1;
      // добавление ссылки на drawing в content-type
      report.files[0].data = report.files[0].data.replace(
        "</Types>",
        `<Override PartName="/xl/drawings/drawing${drawingId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
      );
      // добавление ссылки на drawing в sheet
      report.files[sheet.index].data = report.files[sheet.index].data.replace(
        "</worksheet>",
        `<drawing r:id="rId${maxSheetRelsId + 1}"/></worksheet>`
      );
      let maxSheetDrawingRelsId = Math.max(
        0,
        ...(
          report.files[sheetDrawingRelsIndex].data.match(/Id=".*?"/) || []
        ).map(
          (elem) => +elem.substring(elem.indexOf("rId") + 3, elem.length - 1)
        )
      );
      const sheetJson = JSON.parse(
        xmljs.xml2json(sheet.file.data, { compact: true })
      ).worksheet;
      images.forEach((image, index) => {
        let imageParams = getImagePosition(sheetJson, image.fileName);
        let imageRid = maxSheetDrawingRelsId + index + 1;
        report.files[sheetDrawingRelsIndex].data = report.files[
          sheetDrawingRelsIndex
        ].data.replace(
          "</Relationships>",
          `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${image.encodedImageName}.png"/></Relationships>`
        );
        report.files[sheetDrawingIndex].data = report.files[
          sheetDrawingIndex
        ].data.replace(
          "</xdr:wsDr>",
          `<xdr:oneCellAnchor><xdr:from><xdr:col>${imageParams.coords.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${imageParams.coords.row}</xdr:row><xdr:rowOff>1</xdr:rowOff></xdr:from><xdr:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${imageRid}" name="Изображение"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C3C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns="" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`
        );
      });
    } else {
      // sheetRels есть, ищем указатель на drawings
      let drawingLink = report.files[sheetRelsIndex].data.match(
        /drawings\/drawing.?\.xml/
      );
      if (!drawingLink) {
        // ищем все drawings
        let allDrawings = report.files.filter(
          (file) => file.name.indexOf("drawings/drawing") != -1
        );
        // берем максимальное название
        if (allDrawings.length) {
          drawingId = Math.max(
            0,
            ...allDrawings.map(
              (file) =>
                +file.name.substring(
                  file.name.indexOf("xl/drawings/drawing") + 19,
                  file.name.lastIndexOf(".")
                )
            )
          );
          drawingId++;
        }
        // создаем новый файл с макс названием +1
        report.files.push(createBlankDrawingRelsFile(drawingId));
        sheetDrawingRelsIndex = report.files.length - 1;
        // поиск максимального идентификатор в ссылка листа
        let maxSheetRelsId = Math.max(
          0,
          ...(report.files[sheetRelsIndex].data.match(/Id=".*?"/) || []).map(
            (elem) => +elem.substring(elem.indexOf("rId") + 3, elem.length - 1)
          )
        );
        // вставка указателя на drawingRels
        report.files[sheetRelsIndex].data = report.files[
          sheetRelsIndex
        ].data.replace(
          "</Relationships>",
          `<Relationship Id="rId${
            maxSheetRelsId + 1
          }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingId}.xml"/></Relationships>`
        );
        // создание drawingRels
        let sheetDrawingIndex = 0;
        report.files.push(createBlankDrawingFile(drawingId));
        sheetDrawingIndex = report.files.length - 1;
        // добавление ссылки на drawing в content-type
        report.files[0].data = report.files[0].data.replace(
          "</Types>",
          `<Override PartName="/xl/drawings/drawing${drawingId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
        );
        // добавление ссылки на drawing в sheet
        report.files[sheet.index].data = report.files[sheet.index].data.replace(
          "</worksheet>",
          `<drawing r:id="rId${maxSheetRelsId + 1}"/></worksheet>`
        );

        let maxSheetDrawingRelsId = Math.max(
          0,
          ...(
            report.files[sheetDrawingRelsIndex].data.match(/Id=".*?"/g) || []
          ).map(
            (elem) => +elem.substring(elem.indexOf("rId") + 3, elem.length - 1)
          )
        );
        const sheetJson = JSON.parse(
          xmljs.xml2json(sheet.file.data, { compact: true })
        ).worksheet;
        images.forEach((image, index) => {
          let imageParams = getImagePosition(sheetJson, image.fileName);
          let imageRid = maxSheetDrawingRelsId + index + 1;
          report.files[sheetDrawingRelsIndex].data = report.files[
            sheetDrawingRelsIndex
          ].data.replace(
            "</Relationships>",
            `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${image.encodedImageName}.png"/></Relationships>`
          );
          report.files[sheetDrawingIndex].data = report.files[
            sheetDrawingIndex
          ].data.replace(
            "</xdr:wsDr>",
            `<xdr:oneCellAnchor><xdr:from><xdr:col>${imageParams.coords.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${imageParams.coords.row}</xdr:row><xdr:rowOff>1</xdr:rowOff></xdr:from><xdr:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${imageRid}" name="Изображение"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C3C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns="" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`
          );
        });
      } else {
        // если ссылка на drawings есть
        let sheetDrawingIndex = 0;
        sheetDrawingIndex = report.files.findIndex(
          (file) => file.name.indexOf(drawingLink) != -1
        );
        let relNumber = report.files[sheetDrawingIndex].name.substring(
          report.files[sheetDrawingIndex].name.indexOf("xl/drawings/drawing") +
            19,
          report.files[sheetDrawingIndex].name.lastIndexOf(".")
        );
        sheetDrawingRelsIndex = report.files.findIndex(
          (file) =>
            file.name.indexOf(
              `xl/drawings/_rels/drawing${relNumber}.xml.rels`
            ) != -1
        );
        if (sheetDrawingRelsIndex == -1) {
          report.files.push(createBlankDrawingRelsFile(relNumber));
          sheetDrawingRelsIndex = report.files.length - 1;
        }
        let maxSheetDrawingRelsId = Math.max(
          0,
          ...(
            report.files[sheetDrawingRelsIndex].data.match(/Id=".*?"/) || []
          ).map(
            (elem) => +elem.substring(elem.indexOf("rId") + 3, elem.length - 1)
          )
        );
        const sheetJson = JSON.parse(
          xmljs.xml2json(sheet.file.data, { compact: true })
        ).worksheet;
        images.forEach((image, index) => {
          let imageParams = getImagePosition(sheetJson, image.fileName);
          let imageRid = maxSheetDrawingRelsId + index + 1;
          report.files[sheetDrawingRelsIndex].data = report.files[
            sheetDrawingRelsIndex
          ].data.replace(
            "</Relationships>",
            `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${image.encodedImageName}.png"/></Relationships>`
          );
          report.files[sheetDrawingIndex].data = report.files[
            sheetDrawingIndex
          ].data.replace(
            "</xdr:wsDr>",
            `<xdr:oneCellAnchor><xdr:from><xdr:col>${imageParams.coords.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${imageParams.coords.row}</xdr:row><xdr:rowOff>1</xdr:rowOff></xdr:from><xdr:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${imageRid}" name="Изображение"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C3C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns="" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`
          );
        });
      }
    }
  });
}
/**
 * Creating an empty sheet drawing file
 * @param {number} sheetNumber
 * @returns {Object}
 */
function createBlankDrawingFile(sheetNumber) {
  return {
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>`,
    isMarked: true,
    name: `xl/drawings/drawing${sheetNumber}.xml`,
    parent: "",
  };
}
/**
 * @param {Object} sheet  - sheet data
 * @param {string} imagePlaceHolder
 * @returns {Object}
 */
function getImagePosition(sheet, imagePlaceHolder) {
  const rows = sheet.sheetData.row;
  const colsAttributes = sheet.cols;
  let imageParams = {
    coords: {
      row: 0,
      col: 0,
    },
    size: {
      row: 0,
      col: 0,
    },
  };
  rows.forEach((row, rowIndex) => {
    let columns;
    if (!Array.isArray(row.c)) columns = [row.c];
    else columns = row.c;
    columns.forEach((column, colIndex) => {
      if (!column) return;
      if (!column.is) return;
      if (!column.is.t) return;
      if (!column.is.t._text) return;
      if (column.is.t._text.indexOf(imagePlaceHolder) != -1) {
        imageParams.coords.row = rowIndex;
        imageParams.coords.col = colIndex;
        let imageSizeFormatter =
          column.is.t._text.indexOf(":") != -1
            ? column.is.t._text.substring(column.is.t._text.indexOf(":") + 1)
            : null;
        let imageWidth, imageHeight;
        if (imageSizeFormatter) {
          [imageWidth, imageHeight = imageWidth] =
            imageSizeFormatter.split("*");
          imageParams.size.row = imageHeight * 12700;
          imageParams.size.col = imageWidth * 12700;
        } else {
          if (row && row._attributes && row._attributes.ht)
            imageParams.size.row = (row._attributes.ht * 12500).toFixed(0);
          else imageParams.size.row = (100 * 12500).toFixed(0);
          if (
            colsAttributes &&
            colsAttributes.col &&
            !Array.isArray(colsAttributes.col)
          )
            colsAttributes.col = [colsAttributes.col];
          if (
            colsAttributes &&
            colsAttributes.col &&
            colsAttributes.col[imageParams.coords.col] &&
            colsAttributes.col[imageParams.coords.col]._attributes &&
            colsAttributes.col[imageParams.coords.col]._attributes.width
          )
            imageParams.size.col = (
              colsAttributes.col[imageParams.coords.col]._attributes.width *
              70000
            ).toFixed(0);
        }
      }
    });
  });
  return imageParams;
}

module.exports = { getImagesFromData, putImagesToDocument, putImagesToSheets };
