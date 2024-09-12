# Docx-merger

Javascript Library for Merging Docx files.
This is a fork of https://github.com/scholarcy/docx-merge with updated dependencies.

It's only been tested with NodeJS but according to the original authors it should work in the browser too.

The fork has replaced webpack with esbuild as the build tool, and yarn for npm.

## Purpose

To merge docx file using javascript which has rich contents.

The Library Preserves the Styles, Tables, Images, Bullets and Numberings of input files.

## Table of Contents

1. [Installation](#installation)
2. [Usage Nodejs](#usage-nodejs)
3. [Usage Browser](#usage-browser)
4. [TODO](#todo)
5. [Known Issues](#known-issues)

## Installation

  ```bash
  npm install @scholarcy/docx-merger
  ```

**[Back to top](#table-of-contents)**

### Usage in Nodejs

Read input files as binary and pass it to the `DocxMerger` constructor fuction as a array of files.

Then call the save function with first argument as `nodebuffer`, check the example below.

  ```javascript
  const DocxMerger = require('./../src/index');

const fs = require('fs');
const path = require('path');

(async () => {
  const file1 = fs.readFileSync(path.resolve(__dirname, 'template.docx'), 'binary');
  const file2 = fs.readFileSync(path.resolve(__dirname, 'template1.docx'), 'binary');

  const docx = new DocxMerger();
  await docx.initialize({}, [file1, file2]);
  const data = await docx.save('nodebuffer');
  fs.writeFileSync("output.docx", data);
})()
  ```

**[Back to top](#table-of-contents)**

### TODO

- CLI Support
- Unit Tests
- ES6 Conversion

**[Back to top](#table-of-contents)**

### Known Issues

- Microsoft Word in windows Shows some error due to numbering.

**[Back to top](#table-of-contents)**
