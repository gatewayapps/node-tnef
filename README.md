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

Find or create a directory that contains the TNEF files you wish to parse. Once you've identified the directory, run:
`node-tnef parse <path to your TNEF files>`

The TNEF parser will enumerate every file in the directory. If the file does not contain the TNEF signature, it will output to the console and move to the next file. If the file contains the TNEF signature, the parser will extract the attachment contents and write them to the new folder `<path to your TNEF files>/processed`.

# Issues/Feedback?
Create an issue at this repo and I will try my best to fix the problem or implement the suggestion.
