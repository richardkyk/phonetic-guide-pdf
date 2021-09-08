const PDFDocument = require("pdfkit");
const fs = require("fs");
const XLSX = require("xlsx");

// https://github.com/foliojs/pdfkit/issues/346
// Check out the link for spacing calculations
// You can calculate whether the pinyin is longer or the character
// Then offset both accordingly

const A4 = [595.28, 841.89];
// const font = `${__dirname}\\msyh.ttf`;
const font = `${__dirname}\\kaiu.ttf`;
const fontEnglish = `${__dirname}\\msyh.ttf`;
const fontSize = 22; // Font size of the chinese characters default: 20
const pinyinSize = 8; // Font size of the pinyin default: 10
const pinyinOffset = 2; // Distance pinyin is above character
const titleSize = 30; // Font size of the title default: 30
const characterSpacing = 8; // Distance between letters default: 5
const margin = 64; // Margin top, bottom, left and right default: 64

const files = process.argv.slice(2);

files.forEach((file) => {
  const doc = new PDFDocument({ autoFirstPage: false, bufferPages: true });
  const filename = file.replace(".xlsx", "");
  doc.filename = filename;
  doc.addPage({
    margin: 0,
    size: "A4",
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
  phrases.forEach((phrase) => {
    let x = margin;

    switch (phrase.align) {
      case "newline":
        return;
      case "break":
        anchorRow = 1;
        doc.addPage({
          margin: 0,
          size: "A4",
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

    let y =
      margin +
      (phrase.row - anchorRow) * (pinyinSize + fontSize + characterSpacing);

    if (y >= A4[1] - margin) {
      y = margin;
      anchorRow = phrase.row;
      doc.addPage({
        margin: 0,
        size: "A4",
      });
    }
    if (phrase.row == 0) {
      writeText(doc, phrase, titleSize, x, y / 2, characterSpacing);
    } else {
      writeText(doc, phrase, fontSize, x, y, characterSpacing);
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
    let pageNum = `${doc.filename.replace(".\\", "")} ${i + 1}/${range.count}`;
    doc
      .fill([7, 5, 0, 83])
      .font(font)
      .fontSize(10)
      .text(pageNum, 565 - getWidth(doc, pageNum, 10), 810);
  }
  doc.end();
}

function getWidth(doc, text, fontSize, characterSpacing = null) {
  return doc.font(font).fontSize(fontSize).widthOfString(text, {
    characterSpacing,
  });
}
function getPinyinWidth(doc, text, fontSize, characterSpacing = null) {
  return doc.font(fontEnglish).fontSize(fontSize).widthOfString(text, {
    characterSpacing,
  });
}

function writeText(doc, text, fontSize, x, y, characterSpacing = null) {
  // Chinese characters
  doc
    .fill([7, 5, 0, 83])
    .font(font)
    .fontSize(fontSize)
    .text(text.chinese.replace(/ /g, "　"), x, y + pinyinSize + pinyinOffset, {
      characterSpacing,
      lineBreak: false,
    });

  // Pinyin
  const words = text.pinyin.split(" ");
  const chars = text.chinese.replace(/ /g, "　").split("");
  const charWidth = getWidth(doc, text.chinese.slice(0, 1), fontSize);
  for (const [i, char] of chars.entries()) {
    if (char == "　") {
      words.splice(i, 0, " ");
    }
  }

  for (const word of words) {
    const pinyinWidth = getPinyinWidth(doc, word, pinyinSize);
    const offset = (charWidth - pinyinWidth) / 2;

    doc
      .fill([7, 5, 0, 83])
      .font(fontEnglish)
      .fontSize(pinyinSize)
      .text(word, x + offset, y);
    x += charWidth + characterSpacing;
  }
}
