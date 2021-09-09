const pinyin = require("chinese-to-pinyin");
const zhuyin = require("zhuyin");
const XLSX = require("xlsx");

function generatePhonetics(filename) {
  const workbook = XLSX.readFile(filename);
  var sheet_name_list = workbook.SheetNames;
  let data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

  const exportData = [];
  for (phrase of data) {
    if (!["break", "newline"].includes(phrase.align)) {
      const chinesePhrases = phrase.chinese
        .replace(/[。：，； 　]/g, ";")
        .split(";");
      const output = [];
      for (const chinesePhrase of chinesePhrases) {
        output.push(pinyin(chinesePhrase));
      }
      phrase.pinyin = output.join("  ");

      if (!phrase.zhuyin) {
        phrase.zhuyin = zhuyin(phrase.pinyin).join(" ");
      }
    }
    exportData.push(phrase);
  }

  download(filename, exportData);
}

function download(filename, exportData) {
  const data = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, data);
  const name = `${filename}`;
  XLSX.writeFile(wb, name);
}

const files = process.argv.slice(2);
files.forEach((file) => {
  generatePhonetics(file);
});
