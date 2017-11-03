'use strict';

Object.defineProperty(exports, "__esModule", {
    value: true
});
exports.builder = exports.desc = exports.command = undefined;
exports.handler = handler;

var _bluebird = require('bluebird');

var _bluebird2 = _interopRequireDefault(_bluebird);

var _convertString = require('convert-string');

var _convertString2 = _interopRequireDefault(_convertString);

var _util = require('../bin/util');

var _util2 = _interopRequireDefault(_util);

var _mapi = require('../bin/mapi');

var _mapi2 = _interopRequireDefault(_mapi);

var _fs = require('fs');

var _fs2 = _interopRequireDefault(_fs);

var _path = require('path');

var _path2 = _interopRequireDefault(_path);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

const command = exports.command = 'parse <directory>';

const desc = exports.desc = 'Parses TNEF files inside a specified directory';

const builder = exports.builder = {
    directory: {
        alias: 'd',
        default: undefined,
        type: 'string',
        describe: 'Directory to scan and parse',
        demandOption: true
    }

    // standard TNEF signature
};const tnefSignature = 0x223e9f78;

const lvlMessage = 0x01;
const lvlAttachment = 0x02;

// These can be used to figure out the type of attribute
// an object is
const ATTOWNER = 0x0000; // Owner
const ATTSENTFOR = 0x0001; // Sent For
const ATTDELEGATE = 0x0002; // Delegate
const ATTDATESTART = 0x0006; // Date Start
const ATTDATEEND = 0x0007; // Date End
const ATTAIDOWNER = 0x0008; // Owner Appointment ID
const ATTREQUESTRES = 0x0009; // Response Requested
const ATTFROM = 0x8000; // From
const ATTSUBJECT = 0x8004; // Subject
const ATTDATESENT = 0x8005; // Date Sent
const ATTDATERECD = 0x8006; // Date Received
const ATTMESSAGESTATUS = 0x8007; // Message Status
const ATTMESSAGECLASS = 0x8008; // Message Class
const ATTMESSAGEID = 0x8009; // Message ID
const ATTPARENTID = 0x800a; // Parent ID
const ATTCONVERSATIONID = 0x800b; // Conversation ID
const ATTBODY = 0x800c; // Body
const ATTPRIORITY = 0x800d; // Priority
const ATTATTACHDATA = 0x800f; // Attachment Data
const ATTATTACHTITLE = 0x8010; // Attachment File Name
const ATTATTACHMETAFILE = 0x8011; // Attachment Meta File
const ATTATTACHCREATEDATE = 0x8012; // Attachment Creation Date
const ATTATTACHMODIFYDATE = 0x8013; // Attachment Modification Date
const ATTDATEMODIFY = 0x8020; // Date Modified
const ATTATTACHTRANSPORTFILENAME = 0x9001; // Attachment Transport File Name
const ATTATTACHRENDDATA = 0x9002; // Attachment Rendering Data
const ATTMAPIPROPS = 0x9003; // MAPI Properties
const ATTRECIPTABLE = 0x9004; // Receipients
const ATTATTACHMENT = 0x9005; // Attachment
const ATTTNEFVERSION = 0x9006; // TNEF Version
const ATTOEMCODEPAGE = 0x9007; // OEM Codepage
const ATTORIGNINALMESSAGECLASS = 0x9008; //Original Message Class

function handler(argv) {
    const opts = parseOptions(argv);

    console.log('Begin iterating through the directory');
    ProcessDirectory(opts);
}

function parseOptions(argv) {
    if (!argv) {
        throw new Error('No arguments provided!');
    }

    if (!argv.directory) {
        throw new Error('No directory specified!');
    }

    return argv;
}

// right now, adds just the attachment title and data
var addAttr = (obj, attachment) => {
    switch (obj.Name) {
        case ATTATTACHTITLE:
            var byteString = _convertString2.default.bytesToString(obj.Data);
            attachment.Title = byteString.replaceAll('\x00', '');
            break;
        case ATTATTACHDATA:
            attachment.Data = obj.Data;
            break;
    }
};

// DecodeFile is a utility function that reads the file into memory
// before calling the normal Decode function on the data.
var DecodeFile = path => {
    return new _bluebird2.default((resolve, reject) => {
        console.log('Read the supposed TNEF file: ' + path);

        _fs2.default.readFile(path, (err, data) => {
            if (!err) {
                var arr = [...data];
                resolve(Decode(arr, path));
            } else {
                console.log(err);
                reject(err);
            }
        });
    });
};

// Decode will accept a stream of bytes in the TNEF format and extract the
// attachments and body into a Data object.
var Decode = (data, path) => {

    // get the first 32 bits of the file
    var signature = _util2.default.processBytesToInteger(data, 0, 4);

    // if the signature we get doesn't match the TNEF signature, exit
    if (signature !== tnefSignature) {
        console.log('Value of ' + signature + ' did not equal the expected value of ' + tnefSignature);
        return null;
    }

    console.log('Found a valid TNEF signature for ' + path);

    // set the starting offset past the signature
    var offset = 6;
    var attachment = null;
    var tnef = {};
    tnef.Attachments = [];

    // iterate through the data
    while (offset < data.length) {
        // get only the data within the range of offset and the array length
        var tempData = _util2.default.processBytes(data, offset, data.length);
        // decode the TNEF objects
        var obj = decodeTNEFObject(tempData);

        if (!obj) {
            console.log('Did not get a TNEF object back from file ' + path + ', exit');
            break;
        }

        // increment offset based on the returned object's length
        offset += obj.Length;

        // append attributes and attachments
        if (obj.Name === ATTATTACHRENDDATA) {
            // create an empty attachment object to prepare for population
            attachment = {};
            tnef.Attachments.push(attachment);
        } else if (obj.Level === lvlAttachment) {
            // add the attachments
            addAttr(obj, attachment);
        } else if (obj.Name === ATTMAPIPROPS) {
            tnef.Attributes = _mapi2.default.decodeMapi(obj.Data);

            // get the body property if it exists
            for (var attr in tnef.Attributes) {
                switch (attr.Name) {
                    case _mapi2.default.MAPIBody:
                        tnef.Body = attr.Data;
                    case _mapi2.default.MAPIBodyHTML:
                        tnef.BodyHTML = attr.Data;
                }
            }
        }
    }

    // return the final TNEF object
    return tnef;
};

var decodeTNEFObject = data => {
    var tnefObject = {};
    var offset = 0;
    var object = {};

    // level
    object.Level = _util2.default.processBytesToInteger(data, offset, 1);
    offset++;

    // name
    object.Name = _util2.default.processBytesToInteger(data, offset, 2);
    offset += 2;

    // type
    object.Type = _util2.default.processBytesToInteger(data, offset, 2);
    offset += 2;

    // attribute length
    var attLength = _util2.default.processBytesToInteger(data, offset, 4);
    offset += 4;

    // data
    object.Data = _util2.default.processBytes(data, offset, attLength);
    offset += attLength;
    offset += 2;

    // length
    object.Length = offset;

    return object;
};

var ProcessDirectory = opts => {
    // get the directory path from commandline arguments
    // iterate through each file, and run DecodeFile
    _fs2.default.readdir(opts.directory, (err, files) => {
        if (err) {
            console.error('Could not list the directory: ', err);
            process.exit(1);
        }

        _bluebird2.default.each(files, file => {
            const fullPath = _path2.default.join(opts.directory, file);

            const processedPath = _path2.default.join(opts.directory, 'processed');

            if (!_fs2.default.existsSync(processedPath)) {
                _fs2.default.mkdirSync(processedPath);
            }

            console.log('ATTEMPTING TO PARSE ' + fullPath);

            return DecodeFile(fullPath).then(result => {
                // if there is an attachment, extract it and save to file
                if (result && result.Attachments && result.Attachments.length > 0) {
                    for (var a in result.Attachments) {
                        var attachment = result.Attachments[a];

                        _fs2.default.writeFile(_path2.default.join(processedPath, attachment.Title), new Buffer(attachment.Data), err => {
                            console.log(err);
                        });
                    }

                    console.log('Done decoding ' + fullPath + '!!');
                } else {
                    console.log('Something went wrong with parsing ' + fullPath + '. Make sure this is a TNEF file. If you are certain it is, possibly the file is corrupt');
                }
            }).catch(err => {
                console.error('Something went wrong parsing ' + fullPath, err);
            });
        });
    });
};