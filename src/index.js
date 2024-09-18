const JSZip = require('jszip');
const Style = require('./merge-styles');
const Media = require('./merge-media');
const RelContentType = require('./merge-relations-and-content-type');
const bulletsNumbering = require('./merge-bullets-numberings');

class DocxMerger {
  constructor() {
    this._body = [];
    this._pageBreak = true;
    this._style = [];
    this._numbering = [];
    this._files = [];
    this._contentTypes = {};
    this._media = {};
    this._rel = {};
    this._builder = this._body;
  }

  async initialize(options, files) {
    files = files || [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;

    for (const file of files) {
      this._files.push(await new JSZip().loadAsync(file));
    }
    if (this._files.length > 0) {
      await this.mergeBody(this._files)
    }
  }

  insertPageBreak = function () {
    const pb = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

    this._builder.push(pb);
  };

  insertRaw = function (xml) {
    this._builder.push(xml);
  };

  async mergeBody(files) {
    this._builder = this._body;

    RelContentType.mergeContentTypes(files, this._contentTypes);
    Media.prepareMediaFiles(files, this._media);
    RelContentType.mergeRelations(files, this._rel);

    bulletsNumbering.prepareNumbering(files);
    bulletsNumbering.mergeNumbering(files, this._numbering);
    Style.prepareStyles(files, this._style);
    Style.mergeStyles(files, this._style);

    for (let zip of files) {
      let xmlString = await zip.file('word/document.xml').async('string');

      xmlString = xmlString.match(/<w:body>([\s\S]*?)<\/w:body>/)[1].trim();
      if (xmlString.lastIndexOf('<w:sectPr') === 0) {
        xmlString = xmlString.substring(xmlString.lastIndexOf('</w:sectPr>') + 11);
      } else {
        xmlString = xmlString.substring(0, xmlString.lastIndexOf('<w:sectPr'));
      }

      this.insertRaw(xmlString);
      if (this._pageBreak) this.insertPageBreak();
    }
  };

  async save(type) {
    const zip = this._files[0];

    let xmlString = await zip.file('word/document.xml').async('string');

    const startIndex = xmlString.indexOf('<w:body>') + 8;
    const endIndex = xmlString.lastIndexOf('<w:sectPr');
    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), this._body.join(''));

    await RelContentType.generateContentTypes(zip, this._contentTypes);
    await Media.copyMediaFiles(zip, this._media, this._files);
    await RelContentType.generateRelations(zip, this._rel);
    await bulletsNumbering.generateNumbering(zip, this._numbering);
    await Style.generateStyles(zip, this._style);

    zip.file('word/document.xml', xmlString);

    return await zip.generateAsync({
      type: type,
      compression: 'DEFLATE',
      compressionOptions: {
        level: 4
      }
    })
  };
}

module.exports = DocxMerger;
