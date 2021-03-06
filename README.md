# Phonetic Guide


## About The Project
I wasn't satisfied with Microsoft Word's implementation as it was slow to input and hard to format.  
So I created my own tool which adds pinyin/zhuyin phonetic guides to chinese characters and exports it to a PDF.


## Getting Started

### Installation
1. Clone the repo  
```git clone https://github.com/richardkyk/phonetic-guide-pdf.git```
2. Install NPM packages  
```npm install```

### Usage
1. Create a XLSX file similar to the "青花瓷.xlsx" in the examples folders
2. Fill out the row, align and chinese columns
3. Generate the phonetics (prefills the pinyin/zhuyin columns if there isn't any set)  
```node .\generatePhonetics.js 青花瓷.xlsx```  
**You can simply replace the pinyin if it is incorrect**
4. Export the PDF  
```node .\pinyin.js 青花瓷.xlsx```

#### Row
Denotes how many rows the respective row is relative from the previous "break".

#### Align
"centerTitle"  
Centers a title in a larger font

"center"  
Centers the chinese characters

"left"  
Left-aligns the chinese characters

"right"  
Right-aligns the chinese characters

"newline"  
Creates a new line

"break"  
Creates a new page  
**If a break is used, the succeeding rows are relative to this new point (top of a new page)**
