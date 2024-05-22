var fs = require('fs');
var os = require('os');
var path = require('path');
var file = require('./file');
var params = require('./params');
var helper = require('./helper');
var format = require('./format');
var builder = require('./builder');
var input = require('./input');
var preprocessor = require('./preprocessor');
var translator = require('./translator');
var converter = require('./converter');
var debug = require('debug')('carbone');
var dayjs = require('dayjs');
var locales = require('../formatters/_locale');

var carbone = {

  /**
   * This function is NOT asynchronous (It may create the template or temp directory synchronously)
   * @param {Object} options {
   *                           tempPath     : system temp directory by default
   *                           templatePath : it will create the directory if it does not exists
   *                           renderPath   : where rendered files are temporary saved. It will create the directory if it does not exists
   *                           lang         : set default lang of carbone, can be overwrite in carbone.render options.lang
   *                           timezone     : set default timezone of carbone, can be overwrite in carbone.render options.timezone
   *                           translations : overwrite carbone translations object
   *                           currencySource : currency of data, it depends on the locale if empty
   *                           currencyTarget : default target currency when the formatter convCurr is used without target
   *                                            it depends on the locale if empty
   *                           currencyRates  : rates, based on EUR { EUR : 1, USD : 1.14 }
   *                         }
   */
  set : function (options) {
    for (var attr in options) {
      if (params[attr] !== undefined) {
        params[attr] = options[attr];
      }
      else {
        throw Error('Undefined options :' + attr);
      }
    }
    if (options.templatePath !== undefined) {
      if (fs.existsSync(params.templatePath) === false) {
        fs.mkdirSync(params.templatePath, '0755');
      }
      if (fs.existsSync(path.join(params.templatePath, 'lang')) === false) {
        fs.mkdirSync(path.join(params.templatePath, 'lang'), '0755');
      }
      if (options.translations === undefined) {
        translator.loadTranslations(params.templatePath);
      }
    }
    if (options.tempPath !== undefined && fs.existsSync(params.tempPath) === false) {
      fs.mkdirSync(params.tempPath, '0755');
    }
    if (fs.existsSync(params.renderPath) === false) {
      fs.mkdirSync(params.renderPath, '0755');
    }
    if (options.factories !== undefined || options.startFactory !== undefined) {
      converter.init();
    }
    dayjs.tz.setDefault(params.timezone);
    dayjs.locale(params.lang.toLowerCase());
  },

  /**
   * Reset parameters (for test purpose)
   */
  reset : function () {
    // manage node 0.8 / 0.10 differences
    var _nodeVersion = process.versions.node.split('.');
    var _tmpDir = (parseInt(_nodeVersion[0], 10) === 0 && parseInt(_nodeVersion[1], 10) < 10) ? os.tmpDir() : os.tmpdir();

    params.tempPath                = _tmpDir;
    params.templatePath            = process.cwd();
    params.renderPath              = path.join(params.tempPath, 'carbone_render');
    params.factories               = 1;
    params.attempts                = 2;
    params.startFactory            = false;
    params.factoryMemoryFileSize   = 1;
    params.factoryMemoryThreshold  = 50;
    params.converterFactoryTimeout = 60000;
    params.uidPrefix               = 'c';
    params.pipeNamePrefix          = '_carbone';
    params.lang                    = 'en';
    params.timezone                = 'Europe/Paris';
    params.translations            = {};
    params.currencySource          = '';
    params.currencyTarget          = '';
    params.currencyRates           = { EUR : 1, USD : 1.14 };
  },

  /**
   * Add a template in Carbone datastore (template path)
   * @param {String}            fileId   Unique file name. All templates will be saved in the same folder (templatePath). It will overwrite if the template already exists.
   * @param {String|Buffer}     data     The content of the template
   * @param {Function}          callback(err) called when done
   */
  addTemplate : function (fileId, data, callback) {
    /* if(path.isAbsolute(fileId)===true){  //possible with Node v0.11
      return callback('The file id should not be an absolute path: '+fileId);
    }*/
    var _fullPath = path.join(params.templatePath, fileId);
    fs.writeFile(_fullPath, data, function (err) {
      callback(err);
    });
  },

  /**
   * add formatters
   * @param {Object} formatters {toInt: function(d, args, agrs, ...)}
   */
  addFormatters : function (customFormatters) {
    for (var f in customFormatters) {
      input.formatters[f] = customFormatters[f];
    }
  },

  /**
   * Remove a template from the Carbone datastore (template path)
   * @param  {String}   fileId   Unique file name.
   * @param  {Function} callback(err)
   */
  removeTemplate : function (fileId, callback) {
    var _fullPath = path.join(params.templatePath, fileId);
    fs.unlink(_fullPath, callback);
  },


  /**
   * Return the list of possible conversion format
   * @param  {String} documentType  Must be 'document', 'web', 'graphics', 'spreadsheet', 'presentation'
   * @return {Array}                List of format
   */
  listConversionFormats : function (documentType) {
    var _res = [];
    if (format[documentType] === undefined) {
      throw Error('Unknown document type');
    }
    var _doc = format[documentType];
    for (var attr in _doc) {
      var _format = _doc[attr];
      _format.id = attr;
      _res.push(_format);
    }
    return _res;
  },

  /**
   * Render XML directly
   *
   * @param {String}        xml          The XML
   * @param {Object|Array}  data         The data
   * @param {Object}        optionsRaw   The options raw
   * @param {Function}      callbackRaw  The callback raw
   */
  renderXML : function (xml, data, optionsRaw, callbackRaw) {
    input.parseOptions(optionsRaw, callbackRaw, function (options, callback) {
      // Clean XML tags inside Carbone markers and translate
      xml = preprocessor.preParseXML(xml, options);
      return builder.buildXML(xml, data, options, callback);
    });
  },

  /**
   * Renders a template with given datas and return result to the callback function
   *
   * @param {String}       templatePath : file name of the template (or absolute path)
   * @param {Object|Array} data : Datas to be inserted in the template represented by the {d.****}
   * @param {Object}       optionsRaw [optional] : {
   *                          'complement'   : {}    data which is represented by the {c.****}
   *                          'convertTo'    : 'pdf' || { 'formatName', 'formatOptions'} Convert the document in the format specified
   *                          'extension'    : 'odt' || undefined Specify the template extension
   *                          'variableStr'  : ''    pre-declared variables,
   *                          'lang'         : overwrite default lang. Ex. "fr"
   *                          'timezone'     : set timezone for date formatters (Europe/Paris) by default
   *                          'translations' : overwrite all loaded translations {fr: {}, en: {}, es: {}}
   *                          'enum'         : { ORDER_STATUS : ['open', 'close', 'sent']
   *                          'currencySource'   : currency of data, 'EUR'
   *                          'currencyTarget' : default target currency when the formatter convCurr is used without target
   *                          'currencyRates'  : rates, based on EUR { EUR : 1, USD : 1.14 }
   *                          'hardRefresh'  : (default: false) if true, LibreOffice is used to render and refresh the content of the report at the end of Carbone process
   *                          'renderPrefix' : If defined, it returns a path instead of a buffer, and it adds this prefix in the filename
   *                                           The filename will contains also the report name URL Encoded
   *                       }
   * @param {Function}     callbackRaw(err, bufferOrPath, reportName) : Function called after generation with the result
   */
  render : function (templatePath, data, optionsRaw, callbackRaw) {
    input.parseOptions(optionsRaw, callbackRaw, function (options, callback) {
      // open the template (unzip if necessary)
      file.openTemplate(templatePath, function (err, template) {
        if (err) {
          return callback(err, null);
        }
        // Determine the template extension.
        var _extension = file.detectType(template);
        // It takes the user defined one, or use the file type.
        options.extension = optionsRaw.extension || _extension;
        if (options.extension === null) {
          return callback('Unknown input file type. It should be a docx, xlsx, pptx, odt, ods, odp, xhtml, html or an xml file');
        }
        // check and clean convertTo object, options.convertTo contains a clean version of optionsRaw.convertTo
        var _error = input.parseConvertTo(options, optionsRaw.convertTo);
        if (_error) {
          return callback(_error);
        }
        template.reportName = options.reportName;
        template.extension = options.extension;
        preprocessor.execute(template, options, function (err, template) {
          if (err) {
            return callback(err, null);
          }
          // 获取图片
          let images = getImagesFromData(data);
          // parse all files of the template
          walkFiles(template, data, options, 0, function (err, report) {
            if (err) {
              return callback(err, null);
            }
            // 插入图片
            if (template.extension == 'xlsx') putImagesToSheets(report, images);
            if (template.extension == 'docx') putImagesToDocument(report, images);
            // assemble all files and zip if necessary
            file.buildFile(report, function (err, result) {
              if (err) {
                return callback(err, null);
              }
              convert (result, report.reportName, options, function (err, bufferOrFile) {
                if (report.reportName === undefined && typeof bufferOrFile === 'string') {
                  report.reportName = path.basename(bufferOrFile);
                }
                callback(err, bufferOrFile, report.reportName, (options.isDebugActive === true ? options.debugInfo : null) );
              });
            });
          });
        });
      });
    });
  },

  /**
   * Decodes a rendered filename.
   *
   * When carbone.render is called with the options renderPrefix, the callback returns a path instead of a buffer
   * The filename is built like this (3 distinct parts), with only alphanumeric characters to be able to write it on the disk safely
   *
   * <prefix><22-random-chars><encodedReportName.extension>
   *
   * This function decodes the part `<encodedReportName.extension>` `
   *
   * @param  {String}   pathOrFilename  The path or filename
   * @param  {Integer}  prefixLength    The prefix length used in options.renderPrefix
   * @return {Object}   {
   *                      extension  : 'pdf',
   *                      reportName : 'decoded filename'
   *                    }
   */
  decodeRenderedFilename : function (pathOrFilename, prefixLength = 0) {
    var _filename = path.basename(pathOrFilename);
    var _extension = path.extname(_filename);
    var _onlyReportName = _filename.slice(prefixLength + helper.RANDOM_STRING_LENGTH, -_extension.length);

    return {
      reportName : helper.decodeSafeFilename(_onlyReportName),
      extension  : _extension.slice(1)
    };
  },

  /**
   * Return the file extension
   * @param {String} filePath File path
   * @param {Function} callback
   */
  getFileExtension : function (filePath, callback) {
    file.openTemplate(filePath, function (err, template) {
      if (err) {
        return callback(err);
      }

      var ext = file.detectType(template);

      if (ext === null) {
        return callback('Cannot detect file extension');
      }

      return callback(null, ext);
    });
  },

  /**
   * Convert a file format to another
   *
   * @param  {Buffer}   data      raw data returned by fs.readFile (no utf-8)
   * @param  {Object}   options   Same as carbone.render
   *                              {
   *                                convertTo : pdf || {formatName, formatOptions}
   *                                extension : 'csv'  // extension of input file returned by path.extname(filename).
   *                                                   // It helps LibreOffice to understand the source format (mandatory for CSV files)
   *                              }
   * @param  {Function} callback  (err, result) result is a buffer (file converted)
   */
  convert : function (fileBuffer, optionsRaw, callbackRaw) {
    input.parseOptions(optionsRaw, callbackRaw, function (options, callback) {
      options.extension = optionsRaw.extension;
      if (options.extension === null) {
        return callback('Unknown input file type. options.extension should be equals to docx, xlsx, pptx, odt, ods, odp, xhtml, html or an xml');
      }
      var _error = input.parseConvertTo(options, optionsRaw.convertTo);
      if (_error) {
        return callback(_error);
      }
      convert (fileBuffer, undefined, options, callback);
    });
  },

  formatters : input.formatters
};

/** ***************************************************************************************************************/
/* Privates methods */
/** ***************************************************************************************************************/

/**
 * Parse and compute XML for all files of the template
 * @param  {Object}   template     template file returned by file.js
 * @param  {Object}   data         data to insert
 * @param  {Object}   options      {'complement', 'variables', ...}
 * @param  {Integer}  currentIndex currently visited files in the template
 * @param  {Function} callback(err, template)
 */
function walkFiles (template, data, options, currentIndex, callback) {
  if (currentIndex >= template.files.length) {
    // we have parsed all files, now parse the reportName
    if (template.reportName !== undefined) {
      builder.buildXML(template.reportName, data, options, function (err, reportNameResult) {
        template.reportName = reportNameResult;
        callback(null, template);
      });
    }
    else {
      callback(null, template);
    }
    return;
  }
  var _file = template.files[currentIndex];
  if (_file.isMarked===true) {
    builder.buildXML(_file.data, data, options, function (err, xmlResult) {
      if (err) {
        return callback(err, template);
      }
      _file.data = xmlResult;
      process.nextTick(function () {
        walkFiles(template, data, options, ++currentIndex, callback);
      });
    });
  }
  else {
    walkFiles(template, data, options, ++currentIndex, callback);
  }
}

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
          if (data[key][i][nestedKey] && data[key][i][nestedKey].indexOf && data[key][i][nestedKey].indexOf('base64') != -1 || data[key][i][nestedKey] && (Buffer.isBuffer(data[key][i][nestedKey]) || data[key][i][nestedKey].type == 'Buffer' || data[key][i][nestedKey].BYTES_PER_ELEMENT != undefined)) {
            if (data[key][i][nestedKey].type == 'Buffer' || data[key][i][nestedKey].BYTES_PER_ELEMENT != undefined) data[key][i][nestedKey] = Buffer.from(data[key][i][nestedKey]);
            acc.push({
              fileName: `${key}[${i}].${nestedKey}`,
              data: Buffer.from(data[key][i][nestedKey].toString('utf8').replace(/data:image\/(jpeg|png);base64,/, ''), 'base64')
            });
            data[key][i][nestedKey] = `${key}[${i}].${nestedKey}`;
          }
        }
      }
    } else {
      if (data[key].indexOf && data[key].indexOf('base64') != -1) {
        acc.push({
          fileName: `${key}`,
          data: Buffer.from(data[key].replace(/data:image\/(jpeg|png);base64,/, ''), 'base64')
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
  images = images.map(image => {
    let encodedImageName = Buffer.from(image.fileName).toString('base64');
    report.files.push({
      data: image.data,
      isMarked: false,
      name: `word/media/${encodedImageName}.png`,
      parent: ""
    });
    return {
      ...image,
      encodedImageName
    };
  })
  report.files[0].data = report.files[0].data.replace(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>'
  );
  let documentRelsIndex = report.files.findIndex(file => file.name.indexOf('document.xml.rels') != -1);
  let documentIndex = report.files.findIndex(file => file.name.indexOf('word/document.xml') != -1);
  let maxDocumentRelsId = Math.max(0, ...(report.files[documentRelsIndex].data.match(/Id=".*?"/g) || []).map(elem => +elem.substring(elem.indexOf('rId') + 3, elem.length - 1)));
  images.forEach((image, index) => {
    let imageRid = maxDocumentRelsId + index + 1;
    report.files[documentRelsIndex].data = report.files[documentRelsIndex].data.replace(
      '</Relationships>',
      `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${image.encodedImageName}.png"/></Relationships>`
    );
    let imageSizeFormatter = (report.files[documentIndex].data.match(new RegExp(`<w:t>${Buffer.from(image.encodedImageName, 'base64').toString('utf8').replace(/(\[|\]|\.)/g, '\\$1')}:(\\d+\\*\\d+|\\d+)<\\/w:t>`)) || [])[1];
    let imageWidth, imageHeight;
    if (imageSizeFormatter) {
      ([imageWidth, imageHeight = imageWidth] = imageSizeFormatter.split('*'));
      imageWidth *= 12700;
      imageHeight *= 12700;
    } else {
      imageWidth = 2095500;
      imageHeight = 2095500;
    }
    report.files[documentIndex].data = report.files[documentIndex].data.replace(
      new RegExp(`<w:t>${Buffer.from(image.encodedImageName, 'base64').toString('utf8').replace(/(\[|\]|\.)/g, '\\$1')}(:\\d+\\*\\d+|:\\d+|)<\\/w:t>`),
      `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${imageWidth}" cy="${imageHeight}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Рисунок 1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Аннотация 2019-04-04 111910.jpg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageWidth}" cy="${imageHeight}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
    );
  })
}
/**
 * Insert images into sheets
 * @param {Object} report - report data
 * @param {Array} images - images
 */
function putImagesToSheets(report, images = []) {
  if (!images.length) return;
  images = images.map(image => {
    let encodedImageName = Buffer.from(image.fileName).toString('base64');
    report.files.push({
      data: image.data,
      isMarked: false,
      name: `xl/media/${encodedImageName}.png`,
      parent: ""
    });
    return {
      ...image,
      encodedImageName
    };
  })
  report.files[0].data = report.files[0].data.replace(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>'
  );
  let sheets = [];
  report.files.forEach((file, index) => {
    if (file.name.indexOf('xl/worksheets/sheet') != -1) sheets.push({ file, sheetNumber: file.name.substring(file.name.indexOf('/sheet') + 6, file.name.lastIndexOf('.')), index });
  })
  sheets.forEach((sheet) => {
    let sheetHasImages = images.reduce((acc, image) => { if (sheet.file.data.indexOf(image.fileName) != -1) acc++; return acc }, 0);
    if (!sheetHasImages) return;
    let sheetDrawingRelsIndex = 0;
    let drawingId = 0;
    let sheetRelsIndex = report.files.findIndex(file => file.name.indexOf(`sheet${sheet.sheetNumber}.xml.rels`) != -1);
    // нет sheetRels у листа
    if (sheetRelsIndex == -1) {
      // ищем все drawings
      let allDrawings = report.files.filter(file => file.name.indexOf('drawings/drawing') != - 1);
      // берем максимальное название
      if (allDrawings.length) {
        drawingId = Math.max(0, ...allDrawings.map(file => +file.name.substring(file.name.indexOf('xl/drawings/drawing') + 19, file.name.lastIndexOf('.'))));
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
      let maxSheetRelsId = Math.max(0, ...(report.files[sheetRelsIndex].data.match(/Id=".*?"/g) || []).map(elem => +elem.substring(elem.indexOf('rId') + 3, elem.length - 1)));
      report.files[sheetRelsIndex].data = report.files[sheetRelsIndex].data.replace(
        '</Relationships>',
        `<Relationship Id="rId${maxSheetRelsId + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingId}.xml"/></Relationships>`
      );
      // создание drawingRels
      let sheetDrawingIndex = 0;
      report.files.push(createBlankDrawingFile(drawingId));
      sheetDrawingIndex = report.files.length - 1;
      // добавление ссылки на drawing в content-type
      report.files[0].data = report.files[0].data.replace(
        '</Types>',
        `<Override PartName="/xl/drawings/drawing${drawingId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
      );
      // добавление ссылки на drawing в sheet
      report.files[sheet.index].data = report.files[sheet.index].data.replace(
        '</worksheet>',
        `<drawing r:id="rId${maxSheetRelsId + 1}"/></worksheet>`
      );
      let maxSheetDrawingRelsId = Math.max(0, ...(report.files[sheetDrawingRelsIndex].data.match(/Id=".*?"/) || []).map(elem => +elem.substring(elem.indexOf('rId') + 3, elem.length - 1)));
      const sheetJson = JSON.parse(xmljs.xml2json(sheet.file.data, { compact: true })).worksheet;
      images.forEach((image, index) => {
        let imageParams = getImagePosition(sheetJson, image.fileName);
        let imageRid = maxSheetDrawingRelsId + index + 1;
        report.files[sheetDrawingRelsIndex].data = report.files[sheetDrawingRelsIndex].data.replace(
          '</Relationships>',
          `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${image.encodedImageName}.png"/></Relationships>`
        );
        report.files[sheetDrawingIndex].data = report.files[sheetDrawingIndex].data.replace(
          '</xdr:wsDr>',
          `<xdr:oneCellAnchor><xdr:from><xdr:col>${imageParams.coords.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${imageParams.coords.row}</xdr:row><xdr:rowOff>1</xdr:rowOff></xdr:from><xdr:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${imageRid}" name="Изображение"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C3C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns="" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`
        );
      })
    } else {
      // sheetRels есть, ищем указатель на drawings
      let drawingLink = report.files[sheetRelsIndex].data.match(/drawings\/drawing.?\.xml/);
      if (!drawingLink) {
        // ищем все drawings
        let allDrawings = report.files.filter(file => file.name.indexOf('drawings/drawing') != - 1);
        // берем максимальное название
        if (allDrawings.length) {
          drawingId = Math.max(0, ...allDrawings.map(file => +file.name.substring(file.name.indexOf('xl/drawings/drawing') + 19, file.name.lastIndexOf('.'))));
          drawingId++;
        }
        // создаем новый файл с макс названием +1
        report.files.push(createBlankDrawingRelsFile(drawingId));
        sheetDrawingRelsIndex = report.files.length - 1;
        // поиск максимального идентификатор в ссылка листа
        let maxSheetRelsId = Math.max(0, ...(report.files[sheetRelsIndex].data.match(/Id=".*?"/) || []).map(elem => +elem.substring(elem.indexOf('rId') + 3, elem.length - 1)));
        // вставка указателя на drawingRels
        report.files[sheetRelsIndex].data = report.files[sheetRelsIndex].data.replace(
          '</Relationships>',
          `<Relationship Id="rId${maxSheetRelsId + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingId}.xml"/></Relationships>`
        );
        // создание drawingRels
        let sheetDrawingIndex = 0;
        report.files.push(createBlankDrawingFile(drawingId));
        sheetDrawingIndex = report.files.length - 1;
        // добавление ссылки на drawing в content-type
        report.files[0].data = report.files[0].data.replace(
          '</Types>',
          `<Override PartName="/xl/drawings/drawing${drawingId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
        );
        // добавление ссылки на drawing в sheet
        report.files[sheet.index].data = report.files[sheet.index].data.replace(
          '</worksheet>',
          `<drawing r:id="rId${maxSheetRelsId + 1}"/></worksheet>`
        );

        let maxSheetDrawingRelsId = Math.max(0, ...(report.files[sheetDrawingRelsIndex].data.match(/Id=".*?"/g) || []).map(elem => +elem.substring(elem.indexOf('rId') + 3, elem.length - 1)));
        const sheetJson = JSON.parse(xmljs.xml2json(sheet.file.data, { compact: true })).worksheet;
        images.forEach((image, index) => {
          let imageParams = getImagePosition(sheetJson, image.fileName);
          let imageRid = maxSheetDrawingRelsId + index + 1;
          report.files[sheetDrawingRelsIndex].data = report.files[sheetDrawingRelsIndex].data.replace(
            '</Relationships>',
            `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${image.encodedImageName}.png"/></Relationships>`
          );
          report.files[sheetDrawingIndex].data = report.files[sheetDrawingIndex].data.replace(
            '</xdr:wsDr>',
            `<xdr:oneCellAnchor><xdr:from><xdr:col>${imageParams.coords.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${imageParams.coords.row}</xdr:row><xdr:rowOff>1</xdr:rowOff></xdr:from><xdr:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${imageRid}" name="Изображение"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C3C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns="" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`
          );
        })
      } else {
        // если ссылка на drawings есть
        let sheetDrawingIndex = 0;
        sheetDrawingIndex = report.files.findIndex(file => file.name.indexOf(drawingLink) != -1);
        let relNumber = report.files[sheetDrawingIndex].name.substring(report.files[sheetDrawingIndex].name.indexOf('xl/drawings/drawing') + 19, report.files[sheetDrawingIndex].name.lastIndexOf('.'));
        sheetDrawingRelsIndex = report.files.findIndex(file => file.name.indexOf(`xl/drawings/_rels/drawing${relNumber}.xml.rels`) != -1);
        if (sheetDrawingRelsIndex == -1) {
          report.files.push(createBlankDrawingRelsFile(relNumber));
          sheetDrawingRelsIndex = report.files.length - 1;
        }
        let maxSheetDrawingRelsId = Math.max(0, ...(report.files[sheetDrawingRelsIndex].data.match(/Id=".*?"/) || []).map(elem => +elem.substring(elem.indexOf('rId') + 3, elem.length - 1)));
        const sheetJson = JSON.parse(xmljs.xml2json(sheet.file.data, { compact: true })).worksheet;
        images.forEach((image, index) => {
          let imageParams = getImagePosition(sheetJson, image.fileName);
          let imageRid = maxSheetDrawingRelsId + index + 1;
          report.files[sheetDrawingRelsIndex].data = report.files[sheetDrawingRelsIndex].data.replace(
            '</Relationships>',
            `<Relationship Id="rId${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${image.encodedImageName}.png"/></Relationships>`
          );
          report.files[sheetDrawingIndex].data = report.files[sheetDrawingIndex].data.replace(
            '</xdr:wsDr>',
            `<xdr:oneCellAnchor><xdr:from><xdr:col>${imageParams.coords.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${imageParams.coords.row}</xdr:row><xdr:rowOff>1</xdr:rowOff></xdr:from><xdr:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${imageRid}" name="Изображение"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${imageRid}"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C3C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns="" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${imageParams.size.col}" cy="${imageParams.size.row}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`
          );
        })
      }
    }
  })
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
    parent: ""
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
      col: 0
    },
    size: {
      row: 0,
      col: 0
    }
  };
  rows.forEach((row, rowIndex) => {
    let columns;
    if (!Array.isArray(row.c)) columns = [row.c];
    else columns = row.c;
    columns.forEach((column, colIndex) => {
      if (!column) return
      if (!column.is) return;
      if (!column.is.t) return;
      if (!column.is.t._text) return;
      if (column.is.t._text.indexOf(imagePlaceHolder) != -1) {
        imageParams.coords.row = rowIndex;
        imageParams.coords.col = colIndex;
        let imageSizeFormatter = column.is.t._text.indexOf(':') != -1 ? column.is.t._text.substring(column.is.t._text.indexOf(':') + 1) : null;
        let imageWidth, imageHeight;
        if (imageSizeFormatter) {
          ([imageWidth, imageHeight = imageWidth] = imageSizeFormatter.split('*'));
          imageParams.size.row = imageHeight * 12700;
          imageParams.size.col = imageWidth * 12700;
        } else {
          if (row && row._attributes && row._attributes.ht) imageParams.size.row = (row._attributes.ht * 12500).toFixed(0); else imageParams.size.row = (100 * 12500).toFixed(0);
          if (colsAttributes && colsAttributes.col && !Array.isArray(colsAttributes.col)) colsAttributes.col = [colsAttributes.col];
          if (colsAttributes && colsAttributes.col &&
            colsAttributes.col[imageParams.coords.col] &&
            colsAttributes.col[imageParams.coords.col]._attributes &&
            colsAttributes.col[imageParams.coords.col]._attributes.width) imageParams.size.col = (colsAttributes.col[imageParams.coords.col]._attributes.width * 70000).toFixed(0);
        }
      }
    })
  })
  return imageParams;
}

/**
 * { function_description }
 *
 * @param {Buffer}    inputFileBuffer  The input file buffer
 * @param {String}    customReportName user-defined report name computed by Carbone
 * @param {Function}  options          options coming from input.parseOptions and input.parseConvertTo
 *                                     {
 *                                       convertTo : {
 *                                         extension  : 'pdf',
 *                                         format     : 'writer_pdf_Export' // coming from lib/format.js
 *                                         optionsStr : '44,34,76',         // only for CSV
 *                                         filters    : {                   // only for PDF, JPG, ...
 *                                           ReduceImageResolution : true
 *                                         }
 *                                       }
 *                                       extension    : 'odt' || Force input template extension
 *                                       hardRefresh  : (default: false) if true, LibreOffice is used to render and refresh the content of the report at the end of Carbone process
 *                                       renderPrefix : If defined, it add a prefix to the report name
 *                                     }
 * @param {Function}  callback(err, bufferOrPath) return a path if renderPrefix is defined, a buffer otherwise
 */
function convert (inputFileBuffer, customReportName, options, callback) {
  const _convertTo         = options.convertTo;
  const _hasConversion     = options.extension !== _convertTo.extension || options.hardRefresh === true;
  const _isReturningBuffer = options.renderPrefix === undefined || options.renderPrefix === null;

  // If there is no conversion, and there is no renderPrefix, we return the buffer directly
  if (_hasConversion === false && _isReturningBuffer === true) {
    return callback(null, inputFileBuffer);
  }

  // generate a unique random & safe filename
  const _renderPrefix     = (options.renderPrefix || '').replace(/[^0-9a-z-]/gi, '');
  const _randomNamePart   = helper.getRandomString();
  const _customReportName = customReportName !== undefined ? customReportName : 'report';
  const _renderFilename   = _renderPrefix + _randomNamePart + helper.encodeSafeFilename(_customReportName) + '.' + _convertTo.extension;
  const _renderFile       = path.join(params.renderPath, _renderFilename);

  // no conversion, but return a path
  if (_hasConversion === false) {
    return fs.writeFile(_renderFile, inputFileBuffer, function (err) {
      if (err) {
        debug('Cannot write rendered file on disk' + err);
        return callback('Cannot write rendered file on disk', null);
      }
      return callback(null, _renderFile);
    });
  }

  // A conversion is necessary, generate a intermediate file for the converter
  const _intermediateFilename = _renderPrefix + _randomNamePart + '_tmp.' + options.extension;
  const _intermediateFile     = path.join(params.renderPath, _intermediateFilename);
  fs.writeFile(_intermediateFile, inputFileBuffer, function (err) {
    if (err) {
      debug('Cannot write rendered file on disk' + err);
      return callback('Cannot write rendered file on disk', null);
    }
    // call the converter and tell him to generate directly the wanted filename
    converter.convertFile(_intermediateFile, _convertTo.format, _convertTo.optionsStr, _renderFile, function (errConvert, outputFile) {
      fs.unlink(_intermediateFile, function (err) {
        if (err) {
          debug('Cannot remove intermediate file before conversion ' + err);
        }
      });
      if (errConvert) {
        return callback(errConvert, null);
      }
      if (_isReturningBuffer === false) {
        return callback(null, outputFile);
      }
      fs.readFile(outputFile, function (err, outputBuffer) {
        fs.unlink(outputFile, function (err) {
          if (err) {
            debug('Cannot remove rendered file ' + err);
          }
        });
        if (err) {
          debug('Cannot returned file buffer ' + err);
          return callback('Cannot returned file buffer', null);
        }
        callback(null, outputBuffer);
      });
    });
  });
}


// add default formatters
carbone.addFormatters(require('../formatters/array.js'));
carbone.addFormatters(require('../formatters/condition.js'));
carbone.addFormatters(require('../formatters/date.js'));
carbone.addFormatters(require('../formatters/number.js'));
carbone.addFormatters(require('../formatters/string.js'));
translator.loadTranslations(params.templatePath);

// We must include all locales like this for PKG
require('dayjs/locale/af.js');
require('dayjs/locale/am.js');
require('dayjs/locale/ar-dz.js');
require('dayjs/locale/ar-kw.js');
require('dayjs/locale/ar-ly.js');
require('dayjs/locale/ar-ma.js');
require('dayjs/locale/ar-sa.js');
require('dayjs/locale/ar-tn.js');
require('dayjs/locale/ar.js');
require('dayjs/locale/az.js');
require('dayjs/locale/be.js');
require('dayjs/locale/bg.js');
require('dayjs/locale/bi.js');
require('dayjs/locale/bm.js');
require('dayjs/locale/bn.js');
require('dayjs/locale/bo.js');
require('dayjs/locale/br.js');
require('dayjs/locale/bs.js');
require('dayjs/locale/ca.js');
require('dayjs/locale/cs.js');
require('dayjs/locale/cv.js');
require('dayjs/locale/cy.js');
require('dayjs/locale/da.js');
require('dayjs/locale/de-at.js');
require('dayjs/locale/de-ch.js');
require('dayjs/locale/de.js');
require('dayjs/locale/dv.js');
require('dayjs/locale/el.js');
require('dayjs/locale/en-au.js');
require('dayjs/locale/en-ca.js');
require('dayjs/locale/en-gb.js');
require('dayjs/locale/en-ie.js');
require('dayjs/locale/en-il.js');
require('dayjs/locale/en-in.js');
require('dayjs/locale/en-nz.js');
require('dayjs/locale/en-sg.js');
require('dayjs/locale/en-tt.js');
require('dayjs/locale/en.js');
require('dayjs/locale/eo.js');
require('dayjs/locale/es-do.js');
require('dayjs/locale/es-pr.js');
require('dayjs/locale/es-us.js');
require('dayjs/locale/es.js');
require('dayjs/locale/et.js');
require('dayjs/locale/eu.js');
require('dayjs/locale/fa.js');
require('dayjs/locale/fi.js');
require('dayjs/locale/fo.js');
require('dayjs/locale/fr-ca.js');
require('dayjs/locale/fr-ch.js');
require('dayjs/locale/fr.js');
require('dayjs/locale/fy.js');
require('dayjs/locale/ga.js');
require('dayjs/locale/gd.js');
require('dayjs/locale/gl.js');
require('dayjs/locale/gom-latn.js');
require('dayjs/locale/gu.js');
require('dayjs/locale/he.js');
require('dayjs/locale/hi.js');
require('dayjs/locale/hr.js');
require('dayjs/locale/ht.js');
require('dayjs/locale/hu.js');
require('dayjs/locale/hy-am.js');
require('dayjs/locale/id.js');
require('dayjs/locale/is.js');
require('dayjs/locale/it-ch.js');
require('dayjs/locale/it.js');
require('dayjs/locale/ja.js');
require('dayjs/locale/jv.js');
require('dayjs/locale/ka.js');
require('dayjs/locale/kk.js');
require('dayjs/locale/km.js');
require('dayjs/locale/kn.js');
require('dayjs/locale/ko.js');
require('dayjs/locale/ku.js');
require('dayjs/locale/ky.js');
require('dayjs/locale/lb.js');
require('dayjs/locale/lo.js');
require('dayjs/locale/lt.js');
require('dayjs/locale/lv.js');
require('dayjs/locale/me.js');
require('dayjs/locale/mi.js');
require('dayjs/locale/mk.js');
require('dayjs/locale/ml.js');
require('dayjs/locale/mn.js');
require('dayjs/locale/mr.js');
require('dayjs/locale/ms-my.js');
require('dayjs/locale/ms.js');
require('dayjs/locale/mt.js');
require('dayjs/locale/my.js');
require('dayjs/locale/nb.js');
require('dayjs/locale/ne.js');
require('dayjs/locale/nl-be.js');
require('dayjs/locale/nl.js');
require('dayjs/locale/nn.js');
require('dayjs/locale/oc-lnc.js');
require('dayjs/locale/pa-in.js');
require('dayjs/locale/pl.js');
require('dayjs/locale/pt-br.js');
require('dayjs/locale/pt.js');
require('dayjs/locale/ro.js');
require('dayjs/locale/ru.js');
require('dayjs/locale/rw.js');
require('dayjs/locale/sd.js');
require('dayjs/locale/se.js');
require('dayjs/locale/si.js');
require('dayjs/locale/sk.js');
require('dayjs/locale/sl.js');
require('dayjs/locale/sq.js');
require('dayjs/locale/sr-cyrl.js');
require('dayjs/locale/sr.js');
require('dayjs/locale/ss.js');
require('dayjs/locale/sv.js');
require('dayjs/locale/sw.js');
require('dayjs/locale/ta.js');
require('dayjs/locale/te.js');
require('dayjs/locale/tet.js');
require('dayjs/locale/tg.js');
require('dayjs/locale/th.js');
require('dayjs/locale/tk.js');
require('dayjs/locale/tl-ph.js');
require('dayjs/locale/tlh.js');
require('dayjs/locale/tr.js');
require('dayjs/locale/tzl.js');
require('dayjs/locale/tzm-latn.js');
require('dayjs/locale/tzm.js');
require('dayjs/locale/ug-cn.js');
require('dayjs/locale/uk.js');
require('dayjs/locale/ur.js');
require('dayjs/locale/uz-latn.js');
require('dayjs/locale/uz.js');
require('dayjs/locale/vi.js');
require('dayjs/locale/yo.js');
require('dayjs/locale/zh-cn.js');
require('dayjs/locale/zh-hk.js');
require('dayjs/locale/zh-tw.js');
require('dayjs/locale/zh.js');

// if DayJS does not have a locale defined with country code, define it
// For example "de-de" does not exists in DaysJS, but "de" exists.
// So add locale "de-de" in DaysJS
for (let _locale in locales) {
  if (dayjs.Ls[_locale] === undefined) {
    let _localeWithoutCountry = _locale.replace(/-\S+/g, '');
    if (dayjs.Ls[_localeWithoutCountry] !== undefined) {
      dayjs.locale(_locale, dayjs.Ls[_localeWithoutCountry]);
    }
  }
}

dayjs.extend(require('dayjs/plugin/advancedFormat'));    // Support Quarter, ...
dayjs.extend(require('dayjs/plugin/localizedFormat'));   // Support L LL LLLL
dayjs.extend(require('dayjs/plugin/customParseFormat')); // Support custom format as input
dayjs.extend(require('dayjs/plugin/utc'));
dayjs.extend(require('dayjs-timezone-iana-plugin'));
dayjs.extend(require('dayjs/plugin/isoWeek'));

dayjs.tz.setDefault('Europe/Paris');
dayjs.locale('en');

if (fs.existsSync(params.renderPath) === false) {
  fs.mkdirSync(params.renderPath, '0755');
}

module.exports = carbone;
