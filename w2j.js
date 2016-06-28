#!/usr/local/bin/node

var Converter = require('./import');
var fs = require('fs');
var argv = require('minimist')(process.argv.slice(2));

if (argv.help){
    console.log(`
    Script to convert docx to json and html.

    Usage: node w2j.js input/myfile.docx

    Or, to run at command line, check first line of w2j.js and
    then run ./w2j.js path_to_file

    Options:
      --help: this message
      --outdir: output directory, default is 'output'
    `);
    process.exit(0);
}

var converter = new Converter(this);
var path = (argv._[0]);
if (! path){
    console.log("Full file path required.");
    process.exit(0);
}

fs.stat(path, (err, stat) => {
    if(err == null) {
        converter.import(path, argv.outdir);
    } else {
        console.log('Invalid file path:', err)
    }
});


