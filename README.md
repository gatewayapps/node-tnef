# node-tnef
NodeJS project that will parse Transport Neutral Encapsulation Format (TNEF) files and extract any attachments from the files. This is useful for people who do not have a Microsoft created email client but receive email attachments from people who are using an email client like Outlook.

Based on the GO project: https://github.com/Teamwork/tnef

# What are TNEF files?
Read here for more information: https://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format

In a nutshell, the TNEF format was created by Microsoft as a proprietary format for sending Rich Text Format emails and attachments. Unfortunately, those who are using email clients that are not created by Microsoft(so, not Outlook or Exchange) cannot open these emails/attachments. Most TNEF files are named `winmail.dat` but other names are common as well such as `win.dat` and `Part 1.2`

# How to Install
To install globally: 
`npm install node-tnef -g`

To install to your local folder:
`npm install node-tnef`

# How to Use as a Console application
Run: `node-tnef --help` for a condensed outline of how to use the command line application

`node-tnef` can parse entire directories or just single files.

If you are attempting to parse an entire directory of files, find or create the directory that contains the TNEF files you wish to parse. Once you've identified the directory, run:
`node-tnef parse <path to your TNEF files>` or `node-tnef parse -d <path to your TNEF files>` or `node-tnef parse --directory <path to your TNEF files>`

If you are attempting to parse just a single file, find the file that is the supposed TNEF file you wish to parse. Once you've identified the file, run:
`node-tnef parse -f <path to your TNEF file>` or `node-tnef parse --file <path to your TNEF file>`

The TNEF parser will enumerate every file in the directory. If the file does not contain the TNEF signature, it will output to the console and move to the next file. If the file contains the TNEF signature, the parser will extract the attachment contents and write them to the new folder `<path to your TNEF files>/processed`.

# How to use within a NodeJS project
Just require in `node-tnef` into a NodeJS project. The `node-tnef` library currently has a `parse` method that:
- takes a path to a TNEF formatted file
- a callback function which returns a Buffer containing the decoded TNEF content

There is also a `parseBuffer` method that:
- takes a Buffer representation of the TNEF file
- a callback function which returns a Buffer containing the decoded TNEF content

Object Properties:
- Title - The attachment's title(including file extension)
- Data - The attachment data. Can be used to write to a file using `fs` or another mechanism. Create a new Buffer passing in Data.

Examples:
```
var tnef = require('node-tnef')
var fs = require('fs')
var path = require('path')

tnef.parse('/path/to/the/tnef/file', function (err, content) {
    // here you could write the data to a file for example
    var firstAttachment = content[0]
    fs.writeFile(path.join(aPath, firstAttachment.Title), new Buffer(firstAttachment.Data), (err) => {
        console.log('success!')
    })
    ...
})

tnef.parseBuffer([your buffer of data], function (err, content) {
    // content would contain the result data as a Buffer
    // from here you can write to a file, evaluate the contents(ex. content.Attachments or content.Body)
})
```

# Issues/Feedback?
Create an issue at this repo and I will try my best to fix the problem or implement the suggestion.
