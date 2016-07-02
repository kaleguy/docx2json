#!/usr/local/bin/node

var Converter = require('./import');
var fs = require('fs');
var argv = require('minimist')(process.argv.slice(2));

if (argv.help){
  displayHelp();
}

function displayHelp(){
    console.log(`
    Script to convert docx to json.

    Usage: node w2j.js input/myfile.docx

    Or, to run at command line, check first line of w2j.js and
    then run ./w2j.js path_to_file

    Options:
      --help: this message
      --cleanup: delete intermediate files in output directory
      --outdir=path/to/output: output directory, default is 'output'
      --datauri: if true, convert image files to inline data
    `);
    process.exit(0);
}

var path = (argv._[0]);
if (! path){
    console.log("\n=== No path specified! More info about this program:");
    displayHelp();
    process.exit(0);
}
var converter = new Converter(argv);

fs.stat(path, (err, stat) => {
    if(err == null) {
        converter.import(path, argv.outdir);
    } else {
        console.log('Invalid file path:', err)
    }
});


