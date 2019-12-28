const pinyin = require("chinese-to-pinyin");
const zhuyin = require("zhuyin");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const XLSX = require("xlsx");

// const font = "C://WINDOWS//FONTS//DENGL.TTF";
const font =
  "E://Users//Xepht//Documents//Xepht//Scripts//Temple//phonetic guide//DENGL2.ttf";

const A4 = [595.28, 841.89];
const doc = new PDFDocument({ autoFirstPage: false });
let pageNumber = 1;
doc.on("pageAdded", () => {
  //Add page number to the bottom of the every page
  doc
    .font(font)
    .fontSize(10)
    .text(pageNumber, 540, 780);
  pageNumber++;
});

doc.addPage({
  margin: 0,
  size: "A4"
});

// https://github.com/foliojs/pdfkit/issues/346
// Check out the link for spacing calculations
// You can calculate whether the pinyin is longer or the character
// Then offset both accordingly

// doc.pipe(fs.createWriteStream(`${filename}.pdf`));
const fontSize = 20; // Font size of the chinese characters
const pinyinSize = 10; // Font size of the pinyin
const titleSize = 30; // Font size of the title
const characterSpacing = 5; // Distance between letters
let titleSpacing = characterSpacing * 5; // Distance below the title
const margin = 72; // Margin top, bottom, left and right

function createPDF(phrases) {
  let anchorRow = 0;
  phrases.forEach(phrase => {
    let x = margin;

    switch (phrase.align) {
      case "right":
        x =
          A4[0] - margin - getWidth(phrase.chinese, fontSize, characterSpacing);
        // x = rightAlign(phrase.chinese, fontSize, characterSpacing);
        break;
      case "center":
        x = (A4[0] - getWidth(phrase.chinese, fontSize, characterSpacing)) / 2;
        // return (A4[0] - size) / 2;
        // x = centerAlign(phrase.chinese, fontSize, characterSpacing);
        break;
      case "centerTitle":
        x = (A4[0] - getWidth(phrase.chinese, titleSize, characterSpacing)) / 2;
        // x = centerAlign(phrase.chinese, titleSize, characterSpacing);
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
        size: "A4"
      });
    }
    if (phrase.row == 0) {
      writeText(phrase, titleSize, x, y - titleSpacing, characterSpacing);
    } else {
      writeText(phrase, fontSize, x, y, characterSpacing);
    }
  });
  doc.end();
}

function writeText(text, fontSize, x, y, characterSpacing = null) {
  // Chinese characters
  doc
    .font(font)
    .fontSize(fontSize)
    .text(text.chinese, x, y + pinyinSize, {
      characterSpacing,
      lineBreak: false
    });

  // Pinyin
  const words = text.pinyin.split(" ");
  words.forEach(word => {
    const pinyinWidth = getWidth(word, pinyinSize);
    const offset = (fontSize - pinyinWidth) / 2;
    doc
      .font(font)
      .fontSize(pinyinSize)
      .text(word, x + offset, y);
    x += fontSize + characterSpacing;
  });
}

function getWidth(text, fontSize, characterSpacing = null) {
  return doc
    .font(font)
    .fontSize(fontSize)
    .widthOfString(text, {
      characterSpacing
    });
}

function parseLectureData(filename) {
  const myRegexp = /([^/]+)\.xlsx$/;
  const title = myRegexp.exec(filename)[1];
  // console.log(title);
  const workbook = XLSX.readFile(filename);
  var sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  // console.log(data);
  createPDF(data);
}

const path = "E://Users//Xepht//Documents//Temple//Ceremonies//";
// "C://Users//RichardKYK//Documents//Temple//Ceremonies//";

// const filename = "三天法會";
// const filename = "初一（十五）禮";
// const filename = "參（辭）駕禮";
// const filename = "安座禮";
// const filename = "早晚香禮";
// const filename = "獻供禮";
// const filename = "老中大典禮";
// const filename = "謝恩禮";
// const filename = "辦道禮";
// const filename = "過年禮";
// const filename = "道喜（祝壽）禮";
const filename = "開班禮";

const ceremony = path + filename + ".xlsx";
doc.pipe(fs.createWriteStream(`${filename}.pdf`));

// generatePhonetics(ceremony);
parseLectureData(ceremony);

console.log(zhuyin.fromPinyinSyllable("lv"));
// console.log(pinyin("大家一起懺悔"));
console.log(zhuyin("gè wèi fǎ lv zhù"));

function generatePhonetics(filename) {
  const workbook = XLSX.readFile(filename);
  var sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  const exportData = [];
  data.forEach(phrase => {
    if (!phrase.pinyin) {
      phrase.pinyin = pinyin(phrase.chinese);
    }
    if (!phrase.zhuyin) {
      phrase.zhuyin = zhuyin(phrase.pinyin).join(" ");
    }
    exportData.push(phrase);
  });

  download(exportData);
}

function download(exportData) {
  const data = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, data);
  const name = `Pinyin.xlsx`;
  XLSX.writeFile(wb, name);
}
