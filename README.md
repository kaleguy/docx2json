# Docx to JSON

This is a basic command-line utility for node that will read a .docx file 
and convert it to JSON. It will also opitonally convert to Markdown.

The utility uses XSL to create an intermediate XML document that is then
converted to JSON. Working with the intermediate file is simpler than
working with the original xml files that are zipped into the docx file. The code
in this project that converts the intermediate file to JSON can be adapted
to produce other formats. The main purpose of this project is to provide some sample
 code for 
extracting the content out of a docx file using node.js. For a general purpose conversion 
tool, see [pandoc](http://www.pandoc.org).

## Installation

This script uses the npm java module. In order to get this to run with more recent versions of Java,
you will need to edit the following file (on Mac):

    /Library/Java/JavaVirtualMachines/<version>.jdk/Contents/Info.plist 

Edit the above file to include the lines below:

    <key>JVMCapabilities</key>
    <array>
        ...
        <string>JNI</string>
    </array>


To install, download the project and:

    npm install
    
If you get node-gyp errors, it may be because you need to install Java.     

## Use

To convert a document, place it in the 'input' folder, then run (example, from
terminal window in project folder):

    node w2j.js input/myfile.docx

Or run as command line script:

    chmod +x w2j.js

    node w2j.js input/myfile.docx
    
The output files will be written to the 'output' folder.

For more options run

    node w2j.js input/myfile.docx --help
    
## About
    
Most of the conversion is accomplished via XSL. Parts of the XSL file were 
adapted from [docx2md](https://github.com/matb33/docx2md)

Docx files are zipped collections of xml and image files. This utility creates
a directory with the same name as the file in the 'output' directory. It then 
unzips the docs file to that named directory. Then it creates a file called 
'imported_document.xml' with all of the unzipped xml files put into one. This
xml file is then transformed with an XSL stylesheet into the file 'output_document.xml'. 
This output file is then converted to JSON, and a file 'myfile.json' is written
to the directory with the same name as the file. 

Several examples are in the 'input' and 'output' folders in the project.
 

## License
 
MIT    