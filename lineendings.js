const fs = require('fs');
const path = require('path');

getLineEndings();

function getLineEndings() {

        var textResult = fs.readFileSync(path.join(__dirname, './dist/taskpane.js'), 'utf8');
            var total = [];
            var lines = textResult.split(/\r\n|\n|\r/);
            for (var i = 0; i < lines.length; i++) {
                total.push(lines[i].length + 1);
            }
            console.log(total);

            var charCount = 0;
            for (var i = 0; i < 275; i++) {
                charCount += total[i];
            }
            console.log(charCount);           

}