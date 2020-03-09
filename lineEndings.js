const fs = require('fs');
const path = require('path');

getLineColumn(3745);

function getLineEndings() {
    var total = [];
    var lines = getSplitLines();
    for (var i = 0; i < lines.length; i++) {
        total.push(lines[i].length + 2);
    }
    console.log(total);

    // var charCount = 0;
    // for (var i = 0; i < lines.length; i++) {
    //     charCount += total[i];
    // }
    // console.log(charCount);
    return total;

}

function getLineColumn(offset) {
    const lineEndings = getLineEndings();
    const lines = getSplitLines();
    var columnNumber = 0;
    var lineNumber = 0;
    var charCount = 0;
    for (var i = 0; i < lineEndings.length; i++) {
        charCount += lineEndings[i];
        if (offset < charCount) {
            lineNumber = i;
            columnNumber = charCount - offset;
            break;
        }
    }
    console.log(`lineNumber: ${lineNumber}, columnNumber: ${columnNumber}`)
}

function getSplitLines() {
    var textResult = fs.readFileSync(path.join(__dirname, './dist/taskpane/taskpane.js'), 'utf8');
    var lines = textResult.split(/\r\n|\n|\r/);
    return lines;
}