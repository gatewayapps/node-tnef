String.prototype.replaceAll = function (search, replacement) {
    var target = this
    return target.replace(new RegExp(search, 'g'), replacement)
}

// Convert the byte array to an int value
var byteArrayToInt = ((bytes) => {
    var num = 0
    var n = 0

    for(var b in bytes) {
        num += (parseInt(bytes[b]) << n)
        n += 8
    }

    return num
})

// process bytes from array using offset and length
var processBytes = ((data, offset, length) => {
    var bytesToProcess = []
    for(var i = offset; i < (offset + length); i++) {
        bytesToProcess.push(data[i])
    }

    return bytesToProcess
})

// process bytes from an array to int value using offset and length
var processBytesToInteger = ((data, offset, length) => {
    var bytesToProcess = []
    
    for(var i = offset; i < (offset + length); i++) {
        bytesToProcess.push(data[i])
    }

    return byteArrayToInt(bytesToProcess)
})

module.exports = {
    byteArrayToInt: byteArrayToInt,
    processBytes,
    processBytesToInteger
}