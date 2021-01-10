const PDFDocument = require("pdfkit");
const fs = require("fs");
const XLSX = require("xlsx");

// https://github.com/foliojs/pdfkit/issues/346
// Check out the link for spacing calculations
// You can calculate whether the zhuyin is longer or the character
// Then offset both accordingly

const A4 = [595.28, 841.89];
const font = `${__dirname}\\msyh.ttf`;
const fontSize = 18; // Font size of the chinese characters
const zhuyinSize = 6; // Font size of the zhuyin
const titleSize = 30; // Font size of the title
const characterSpacing = 10; // Distance between letters
const margin = 36; // Margin top, bottom, left and right


const files = process.argv.slice(2)


files.forEach(file => {
  const doc = new PDFDocument({ autoFirstPage: false, bufferPages: true });
  const filename = file.replace(".xlsx", "")
  doc.filename = filename;
  doc.addPage({
    margin: 0,
    size: "A4"
  });

  doc.pipe(fs.createWriteStream(`${filename}.pdf`));
  parseFileData(doc, file);
});

function parseFileData(doc, filename) {
  const workbook = XLSX.readFile(filename);
  var sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  createPDF(doc, data);
}

function createPDF(doc, phrases) {
  let anchorRow = 0;
  phrases.forEach(phrase => {
    let x = margin;

    switch (phrase.align) {
      case "newline":
        return;
      case "break":
        anchorRow = 1;
        doc.addPage({
          margin: 0,
          size: "A4"
        });
        return;
      case "right":
        x =
          A4[0] -
          margin -
          getWidth(doc, phrase.chinese, fontSize, characterSpacing);
        break;
      case "center":
        x =
          (A4[0] - getWidth(doc, phrase.chinese, fontSize, characterSpacing)) /
          2;
        break;
      case "centerTitle":
        x =
          (A4[0] - getWidth(doc, phrase.chinese, titleSize, characterSpacing)) /
          2;
        break;
      default:
        // Else left align is just the margin
        x = margin;
        break;
    }

    let y = margin + (phrase.row - anchorRow) * (characterSpacing + fontSize);

    if (y >= A4[1] - margin) {
      y = margin;
      anchorRow = phrase.row;
      doc.addPage({
        margin: 0,
        size: "A4"
      });
    }
    if (phrase.row == 0) {
      writeZhuyin(doc, phrase, titleSize, x, y / 2, characterSpacing);
    } else {
      writeZhuyin(doc, phrase, fontSize, x, y, characterSpacing);
    }
  });

  // Adding page numbers
  const range = doc.bufferedPageRange(); // => { start: 0, count: 2 }

  for (
    i = range.start, end = range.start + range.count, range.start <= end;
    i < end;
    i++
  ) {
    doc.switchToPage(i);
    let pageNum = `${doc.filename} ${i + 1}/${range.count}`;
    doc
      .font(font)
      .fontSize(10)
      .text(pageNum, 565 - getWidth(doc, pageNum, 10), 815);
  }
  doc.end();
}

function getHeight(doc, text, fontSize, characterSpacing = null) {
  return doc
    .font(font)
    .fontSize(fontSize)
    .heightOfString(text, {
      characterSpacing
    });
}
function getWidth(doc, text, fontSize, characterSpacing = null) {
  return doc
    .font(font)
    .fontSize(fontSize)
    .widthOfString(text, {
      characterSpacing
    });
}
function writeZhuyin(doc, text, fontSize, x, y, characterSpacing = null) {
  // Chinese characters
  doc
    .font(font)
    .fontSize(fontSize)
    .text(text.chinese, x, y, {
      characterSpacing,
      lineBreak: false
    });

  // Zhuyin
  const words = text.zhuyin.split(" ");

  // Splitting the phrase into words
  words.forEach(word => {
    let offset = 0;
    x += fontSize;

    // Splitting the zhuyin into glyphs
    word.split("").forEach(symbol => {
      const zhuyinHeight = getHeight(doc, symbol, zhuyinSize);

      let toneSize = 0;
      let toneOffset = 0;
      if ("`ˇˊ".includes(symbol)) {
        toneSize = 10;
        toneOffset =
          (getWidth(doc, symbol, toneSize + zhuyinSize) - zhuyinSize) / 2 -
          zhuyinSize +
          1; // offset for the tone to bring it 1 unit closer to the character
        offset = (offset - zhuyinHeight) / 2;
      }

      doc
        .font(font)
        .fontSize(zhuyinSize + toneSize)
        .text(symbol, x - toneOffset, y + offset);
      offset += zhuyinHeight;
    });
    x += characterSpacing;
  });
}
