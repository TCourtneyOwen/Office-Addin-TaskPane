var fs = require('fs'),
    path = require('path'),
    sourceMap = require('source-map');

// file output by Webpack, Uglify, etc.
var GENERATED_FILE = path.join(__dirname, '/dist/taskpane/taskpane.js.map');

// line and column located in your generated file (for example, the source of your error
// from your minified file)
var GENERATED_LINE_AND_COLUMN = { line: 64, column: 41 };

var rawSourceMap = fs.readFileSync(GENERATED_FILE).toString();
new sourceMap.SourceMapConsumer(rawSourceMap)
    .then(function (smc) {
        var pos = smc.originalPositionFor(GENERATED_LINE_AND_COLUMN);

        // should see something like:
        // { source: 'original.js', line: 57, column: 9, name: 'myfunc' }
        console.log(pos);
    });