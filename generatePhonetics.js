const pinyin = require("chinese-to-pinyin");
const zhuyin = require("zhuyin");
const XLSX = require("xlsx");

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
// const filename = "開班禮";
const filename = "彌勒救苦真經"

const ceremony = `${__dirname}\\ceremonies\\${filename}.xlsx`;

generatePhonetics(ceremony);
