import Promise from 'bluebird'
import convertString from 'convert-string'
import utils from '../bin/util'
import mapi from '../bin/mapi'
import fs from 'fs'
import path from 'path'
import bunyan from 'bunyan'

export const command = 'parse [directory] [file]'

export const desc = 'Parses TNEF files inside a specified directory'

export const builder = {
    directory: {
        alias: 'd',
        default: undefined,
        type: 'string',
        describe: 'Directory to scan and parse',
        demandOption: false
    },
    file: {
        alias: 'f',
        default: undefined,
        type: 'string',
        describe: 'File to scan and parse',
        demandOption: false
    }
}

const log = bunyan.createLogger({ name: 'node-tnef' })

// standard TNEF signature
const tnefSignature = 0x223e9f78

const lvlMessage = 0x01
const lvlAttachment = 0x02

// These can be used to figure out the type of attribute
// an object is
const ATTOWNER = 0x0000 // Owner
const ATTSENTFOR = 0x0001 // Sent For
const ATTDELEGATE = 0x0002 // Delegate
const ATTDATESTART = 0x0006 // Date Start
const ATTDATEEND = 0x0007 // Date End
const ATTAIDOWNER = 0x0008 // Owner Appointment ID
const ATTREQUESTRES = 0x0009 // Response Requested
const ATTFROM = 0x8000 // From
const ATTSUBJECT = 0x8004 // Subject
const ATTDATESENT = 0x8005 // Date Sent
const ATTDATERECD = 0x8006 // Date Received
const ATTMESSAGESTATUS = 0x8007 // Message Status
const ATTMESSAGECLASS = 0x8008 // Message Class
const ATTMESSAGEID = 0x8009 // Message ID
const ATTPARENTID = 0x800a // Parent ID
const ATTCONVERSATIONID = 0x800b // Conversation ID
const ATTBODY = 0x800c // Body
const ATTPRIORITY = 0x800d // Priority
const ATTATTACHDATA = 0x800f // Attachment Data
const ATTATTACHTITLE = 0x8010 // Attachment File Name
const ATTATTACHMETAFILE = 0x8011 // Attachment Meta File
const ATTATTACHCREATEDATE = 0x8012 // Attachment Creation Date
const ATTATTACHMODIFYDATE = 0x8013 // Attachment Modification Date
const ATTDATEMODIFY = 0x8020 // Date Modified
const ATTATTACHTRANSPORTFILENAME = 0x9001 // Attachment Transport File Name
const ATTATTACHRENDDATA = 0x9002 // Attachment Rendering Data
const ATTMAPIPROPS = 0x9003 // MAPI Properties
const ATTRECIPTABLE = 0x9004 // Receipients
const ATTATTACHMENT = 0x9005 // Attachment
const ATTTNEFVERSION = 0x9006 // TNEF Version
const ATTOEMCODEPAGE = 0x9007 // OEM Codepage
const ATTORIGNINALMESSAGECLASS = 0x9008 //Original Message Class

export function handler(argv) {
    const opts = parseOptions(argv)

    if (opts && opts.directory) {
        log.info('Begin iterating through the directory:' + opts.directory)
        ProcessDirectory(opts.directory)
    } else if (opts && opts.file) {
        log.info('Begin parsing the file:' + opts.file)
        ProcessFile(opts.file)
    } else {
        log.warn('No arguments specified!')
    }
}

/**
 * This callback type is called `parseCallback` and is displayed as a global symbol.
 *
 * @callback parseCallback
 * @param {Object[]} Attachments - Array of TNEF attachments
 */

/**
 * parse a single
 * TNEF file given the file path and a callback
 * @param {string} filePath - The path to the TNEF file
 * @param {parseCallback} callback - The callback
 *  
 */
// method that can be used within another Node module to parse a single
// TNEF file given the file path and a callback
export function parse(filePath, callback) {
    log.info('ATTEMPTING TO PARSE ' + filePath);

    DecodeFile(filePath).then((result) => {
        // if there is an attachment, extract it and save to file
        if (result && result.Attachments && result.Attachments.length > 0) {
            log.info('Done decoding ' + filePath + ' and found attachments!!')
            callback(false, result)
        } else if (result && (result.BodyHTML || result.Body)) {
            log.info('Done decoding ' + filePath + ' and found email body!!')
            callback(false, result)
        } else {
            log.warn('Something went wrong with parsing ' + filePath + '. Make sure this is a TNEF file. If you are certain it is, possibly the file is corrupt')
            callback(true, null)
        }
    }).catch((err) => {
        log.error('Something went wrong parsing ' + filePath, err)
        callback(true, err)
    })
}

function parseOptions(argv) {
    if (!argv) {
        throw new Error('No arguments provided!')
    }

    return argv
}

// right now, adds just the attachment title and data
var addAttr = ((obj, attachment) => {
    switch (obj.Name) {
        case ATTATTACHTITLE:
            var byteString = convertString.bytesToString(obj.Data)
            attachment.Title = byteString.replaceAll('\x00', '')
            break;
        case ATTATTACHDATA:
            attachment.Data = obj.Data
            break;
    }
})

// DecodeFile is a utility function that reads the file into memory
// before calling the normal Decode function on the data.
var DecodeFile = ((path) => {
    return new Promise((resolve, reject) => {
        log.info('Read the supposed TNEF file: ' + path)

        fs.readFile(path, (err, data) => {
            if (!err) {
                var arr = [...data]
                resolve(Decode(arr, path))
            } else {
                log.error(err);
                reject(err)
            }
        })
    })
})

// Decode will accept a stream of bytes in the TNEF format and extract the
// attachments and body into a Data object.
var Decode = ((data, path) => {

    // get the first 32 bits of the file
    var signature = utils.processBytesToInteger(data, 0, 4)

    // if the signature we get doesn't match the TNEF signature, exit
    if (signature !== tnefSignature) {
        log.warn('Value of ' + signature + ' did not equal the expected value of ' + tnefSignature)
        return null
    }

    log.info('Found a valid TNEF signature for ' + path)

    // set the starting offset past the signature
    var offset = 6
    var attachment = null
    var tnef = {}
    tnef.Attachments = []

    // iterate through the data
    while (offset < data.length) {
        // get only the data within the range of offset and the array length
        var tempData = utils.processBytes(data, offset, data.length)
        // decode the TNEF objects
        var obj = decodeTNEFObject(tempData)

        if (!obj) {
            log.error('Did not get a TNEF object back from file ' + path + ', exit')
            break;
        }

        // increment offset based on the returned object's length
        offset += obj.Length

        // append attributes and attachments
        if (obj.Name === ATTATTACHRENDDATA) {
            // create an empty attachment object to prepare for population
            attachment = {}
            tnef.Attachments.push(attachment)
        } else if (obj.Level === lvlAttachment) {
            // add the attachments
            addAttr(obj, attachment)
        } else if (obj.Name === ATTMAPIPROPS) {
            tnef.Attributes = mapi.decodeMapi(obj.Data)
            if (tnef.Attributes) {
                // get the body property if it exists
                for (var attr of tnef.Attributes) {
                    switch (attr.Name) {
                        case mapi.MAPIBody:
                            tnef.Body = attr.Data
                        case mapi.MAPIBodyHTML:
                            tnef.BodyHTML = attr.Data
                    }
                }
            }
        }
    }

    // return the final TNEF object
    return tnef
})

var decodeTNEFObject = ((data) => {
    var tnefObject = {}
    var offset = 0
    var object = {}

    // level
    object.Level = utils.processBytesToInteger(data, offset, 1)
    offset++

    // name
    object.Name = utils.processBytesToInteger(data, offset, 2)
    offset += 2

    // type
    object.Type = utils.processBytesToInteger(data, offset, 2)
    offset += 2

    // attribute length
    var attLength = utils.processBytesToInteger(data, offset, 4)
    offset += 4

    // data
    object.Data = utils.processBytes(data, offset, attLength)
    offset += attLength
    offset += 2

    // length
    object.Length = offset

    return object
})

var ProcessDirectory = ((directory) => {
    // get the directory path from commandline arguments
    // iterate through each file, and run DecodeFile
    fs.readdir(directory, (err, files) => {
        if (err) {
            log.error('Could not list the directory: ', err)
            process.exit(1)
        }

        Promise.each(files, (file) => {
            return ProcessFile(file, directory)
        })
    })
})

// process a single file
var ProcessFile = ((file, directory) => {
    return new Promise((resolve, reject) => {
        var fullPath = file
        var processedPath = path.join(path.dirname(file), 'processed')

        if (directory) {
            fullPath = path.join(directory, file)
            processedPath = path.join(directory, 'processed')
        }

        if (!fs.existsSync(processedPath)) {
            fs.mkdirSync(processedPath)
        }

        log.info('ATTEMPTING TO PARSE ' + fullPath);

        DecodeFile(fullPath).then((result) => {
            // if there is an attachment, extract it and save to file
            if (result && result.Attachments && result.Attachments.length > 0) {
                for (var a in result.Attachments) {
                    var attachment = result.Attachments[a]

                    fs.writeFile(path.join(processedPath, attachment.Title), new Buffer(attachment.Data), (err) => {
                        log.error(err)
                        reject(err)
                    })
                }

                if (result.Body || result.BodyHTML) {
                    fs.writeFile(path.join(processedPath, 'htmlbody.html'), new Buffer(result.Body || result.BodyHTML), (err) => {
                        log.error(err)
                        reject(err)
                    })
                }

                log.info('Done decoding ' + fullPath + '!!')
                resolve(null)
            } else if (result && (result.Body || result.BodyHTML)) {
                fs.writeFile(path.join(processedPath, 'htmlbody.html'), new Buffer(result.Body || result.BodyHTML), (err) => {
                    log.error(err)
                    reject(err)
                })
            } else {
                log.warn('Something went wrong with parsing ' + fullPath + '. Make sure this is a TNEF file. If you are certain it is, possibly the file is corrupt')
                resolve(null)
            }
        }).catch((err) => {
            log.error('Something went wrong parsing ' + fullPath, err)
            reject(err)
        })
    })
})