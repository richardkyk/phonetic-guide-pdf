const PDFDocument = require("pdfkit");
const fs = require("fs");
const XLSX = require("xlsx");

// https://github.com/foliojs/pdfkit/issues/346
// Check out the link for spacing calculations
// You can calculate whether the zhuyin is longer or the character
// Then offset both accordingly

const A4 = [595.28, 841.89];
const font = `${__dirname}\\DENGL.ttf`;
// const font = "C://WINDOWS//FONTS//DENGL.TTF";
const fontSize = 18; // Font size of the chinese characters
const zhuyinSize = 6; // Font size of the zhuyin
const titleSize = 30; // Font size of the title
const characterSpacing = 10; // Distance between letters
let titleSpacing = characterSpacing * 5; // Distance below the title
const margin = 36; // Margin top, bottom, left and right

// This is to create PDFs for individual ceremonies
// const files = ["三天法會"];
// const files = ["初一（十五）禮"]
// const files = ["參（辭）駕禮"]
// const files = ["安座禮"];
// const files = ["早晚香禮"]
// const files = ["獻供禮"]
// const files = ["老中大典禮"]
// const files = ["謝恩禮"];
// const files = ["辦道禮"]
// const files = ["過年禮"]
// const files = ["道喜（祝壽）禮"]
// const files = ["開班禮"]

// This is to create PDFs for all the ceremonies in bulk
const files = [
  "三天法會",
  "初一（十五）禮",
  "參（辭）駕禮",
  "安座禮",
  "早晚香禮",
  "獻供禮",
  "老中大典禮",
  "謝恩禮",
  "辦道禮",
  "過年禮",
  "道喜（祝壽）禮",
  "開班禮"
];

files.forEach(file => {
  const doc = new PDFDocument({ autoFirstPage: false });
  let pageNumber = 1;
  doc.on("pageAdded", () => {
    //Add page number to the bottom of the every page
    doc
      .font(font)
      .fontSize(10)
      .text(pageNumber, 570, 820);
    pageNumber++;
  });

  doc.addPage({
    margin: 0,
    size: "A4"
  });

  const ceremony = `${__dirname}\\ceremonies\\${file}.xlsx`;
  doc.pipe(fs.createWriteStream(`${file}.pdf`));
  parseLectureData(ceremony, doc);
});

function parseLectureData(filename, doc) {
  const workbook = XLSX.readFile(filename);
  var sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  createPDF(data, doc);
}

function createPDF(phrases, doc) {
  let anchorRow = 0;
  phrases.forEach(phrase => {
    let x = margin;

    switch (phrase.align) {
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
          getWidth(phrase.chinese, fontSize, doc, characterSpacing);
        break;
      case "center":
        x =
          (A4[0] - getWidth(phrase.chinese, fontSize, doc, characterSpacing)) /
          2;
        break;
      case "centerTitle":
        x =
          (A4[0] - getWidth(phrase.chinese, titleSize, doc, characterSpacing)) /
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
      writeZhuyin(phrase, titleSize, x, y / 2, doc, characterSpacing);
    } else {
      writeZhuyin(phrase, fontSize, x, y, doc, characterSpacing);
    }
  });
  doc.end();
}

function getHeight(text, fontSize, doc, characterSpacing = null) {
  return doc
    .font(font)
    .fontSize(fontSize)
    .heightOfString(text, {
      characterSpacing
    });
}
function getWidth(text, fontSize, doc, characterSpacing = null) {
  return doc
    .font(font)
    .fontSize(fontSize)
    .widthOfString(text, {
      characterSpacing
    });
}
function writeZhuyin(text, fontSize, x, y, doc, characterSpacing = null) {
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
      const zhuyinHeight = getHeight(symbol, zhuyinSize, doc);

      let toneSize = 0;
      let toneOffset = 0;
      if ("`ˇˊ".includes(symbol)) {
        toneSize = 10;
        toneOffset =
          (getWidth(symbol, toneSize + zhuyinSize, doc) - zhuyinSize) / 2 -
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
