# node-tnef
NodeJS project that will parse Transport Neutral Encapsulation Format (TNEF) files and extract any attachments from the files. This is useful for people who do not have a Microsoft created email client but receive email attachments from people who are using an email client like Outlook.

Based on the GO project: https://github.com/Teamwork/tnef

# How to Install
To install globally: 
`npm install node-tnef -g`

To install to your local folder:
`npm install node-tnef`

# How to Use
Run: `node-tnef --help` for a condensed outline of how to use the command line application

`node-tnef` can parse entire directories or just single files.

If you are attempting to parse an entire directory of files, find or create the directory that contains the TNEF files you wish to parse. Once you've identified the directory, run:
`node-tnef parse <path to your TNEF files>` or `node-tnef parse -d <path to your TNEF files>` or `node-tnef parse --directory <path to your TNEF files>`

If you are attempting to parse just a single file, find the file that is the supposed TNEF file you wish to parse. Once you've identified the file, run:
`node-tnef parse -f <path to your TNEF file>` or `node-tnef parse --file <path to your TNEF file>`

The TNEF parser will enumerate every file in the directory. If the file does not contain the TNEF signature, it will output to the console and move to the next file. If the file contains the TNEF signature, the parser will extract the attachment contents and write them to the new folder `<path to your TNEF files>/processed`.

# Issues/Feedback?
Create an issue at this repo and I will try my best to fix the problem or implement the suggestion.
