const XLSX = require('xlsx')
const { XMLParser, XMLBuilder } = require('fast-xml-parser')
const { writeFile } = require('fs/promises')
const yargs = require('yargs')
const fs = require('fs')

const TAG_TABLE_NAME = 'Imported SPS-Liste'
const LOGFILE_NAME = 'makeplctags.log'

let logfile = {
    alleReihen: 0,
    ImportierbareReihen: 0,
    DoppelteAdressen: [],
}

// Init yargs cli arguments and define help
const { spath: sourcePath, dpath: destPath } = yargs
    .option('spath', {
        type: 'string',
        description: 'Path to source xlsx file exported from EPlan',
        demandOption: false,
        default: 'SPS-Liste.xlsx',
    })
    .option('dpath', {
        type: 'string',
        description: 'Path to destination xml file for import into TIA Portal',
        demandOption: false,
        default: 'PlcTags.xml',
    })
    .help()
    .alias('help', 'h').argv

console.log(`Started ./src/makePlcTags.js, Data is in ${sourcePath}`)

// --
// 1) Read SPS-Liste.xlsx - use npm xlsx
// 2) Parse rows
// 3) Filter column E "SPSAdresse" not empty
// 4) Build Tags according to PLCTags.example.xml
// 5) Write new XML file PlcTags.xml (override existing) - use fast-xml-parser
// --

// 1) Read SPS-Liste.xlsx
let sheet
try {
    const workbook = XLSX.readFile(`${sourcePath}`)

    // Get first worksheet
    const sheetName = workbook.SheetNames[0]
    console.log(`Sheet name: ${sheetName}`)
    sheet = workbook.Sheets[sheetName]
} catch (error) {
    console.error('Source file dose not exist or no xlsx file' + error)
    const logStream = fs.createWriteStream(LOGFILE_NAME, { flags: 'a' })
    logStream.end(
        new Date().toLocaleString() +
            '\n Error: ' +
            error +
            '\n---------------------------------------------------------------------------------\n\n'
    )
    process.exitCode = 1
    return
    //throw new Error('Exit')
}

const plcTags = []
let plcTag = {
    listNumber: 0,
    cpuName: '',
    tagIdPlcModule: '',
    connector: '',
    ioAddress: '',
    dataType: '',
    signalType: '',
    settings: '',
    direction: '',
    symbolicAddress: '',
    functionText: '',
    text: '',
}

// 2) Parse rows
XLSX.utils.sheet_to_json(sheet, { header: 1 }).forEach((row, index) => {
    if (index < 3) return // skip header
    logfile.alleReihen++
    const [
        listNumber,
        cpuName,
        tagIdPlcModule,
        connector,
        ioAddress,
        dataType,
        signalType,
        settings,
        direction,
        symbolicAddress,
        functionText,
        text,
    ] = row
    if (!ioAddress) return // skip empty address
    plcTag = {
        listNumber: parseInt(listNumber) || 0,
        cpuName,
        tagIdPlcModule,
        connector,
        ioAddress,
        dataType,
        signalType,
        settings,
        direction,
        symbolicAddress,
        functionText,
        text,
    }
    plcTags.push(plcTag)
})

// return plcTags with unique ioAddress
const plcTagsFiltered = plcTags.filter((tag, index) => {
    const ioAddress = tag.ioAddress
    const firstIndex = plcTags.findIndex(
        (t, newindex) => t.ioAddress === ioAddress && newindex !== index
    )
    return firstIndex !== -1
})

console.log(`IO - Adresses with more then one definition:`)
plcTagsFiltered.forEach((e) => {
    console.log(`Adresse: ${e.ioAddress}; SPS-Karte: ${e.tagIdPlcModule}`)
    logfile.DoppelteAdressen.push(
        `\tNummer: ${e.listNumber}; Adresse: ${e.ioAddress}; SPS-Karte: ${e.tagIdPlcModule}`
    )
})
logfile.ImportierbareReihen = plcTags.length

// 4) Build Tags according to PLCTags.example.xml
const plcTagsObj = {
    '?xml': {
        '@@version': '1.0',
        '@@encoding': 'utf-8',
    },
    Tagtable: {
        '@@name': TAG_TABLE_NAME,
        Tag: plcTags.map((tag) => {
            const prefixSafety = tag.signalType.includes('SAFETY') ? 'F' : ''
            const prefixInOut =
                tag.direction === 'Eingang' ? 'I' : tag.direction === 'Ausgang' ? 'Q' : 'TODO'
            const text = `${prefixSafety}${prefixInOut}_(${tag.ioAddress})`

            return {
                '@@type': tag.dataType,
                '@@hmiVisible': 'True',
                '@@hmiWriteable': 'False',
                '@@hmiAccessible': 'True',
                '@@retain': 'False',
                '@@remark': `${tag.text} | Plc-Module ${tag.tagIdPlcModule} : ${tag.connector}`,
                '@@addr': `%${tag.ioAddress}`,
                '#text': text,
            }
        }),
    },
}

// Use fast-xml-parser
const options = {
    ignoreAttributes: false,
    attributeNamePrefix: '@@',
    format: true,
    processEntities: false,
}
const builder = new XMLBuilder(options)
const output = builder.build(plcTagsObj)

// 5) Write new XML file PlcTags.xml (override existing)
writeFile(`${destPath}`, output)
    .then(() => {
        console.log(`${destPath} written`)
    })
    .catch((err) => {
        const logStream = fs.createWriteStream(LOGFILE_NAME, { flags: 'a' })
        logStream.end(
            new Date().toLocaleString() +
                '\n Error: ' +
                err +
                '\n---------------------------------------------------------------------------------\n\n'
        )
        console.error(err)
        process.exitCode = 1
        return
        // throw new Error('Exit')
    })

let logfileText =
    new Date().toLocaleString() +
    '\nInput Pfad:' +
    sourcePath +
    '\nOutput Pfad:' +
    destPath +
    '\n' +
    'AlleReihen:           ' +
    logfile.alleReihen +
    '\nImportierte Reihen: ' +
    logfile.ImportierbareReihen +
    '\nDoppelte Adressen:\n' +
    logfile.DoppelteAdressen.join('\n') +
    '\n---------------------------------------------------------------------------------\n\n'

var logStream = fs.createWriteStream(LOGFILE_NAME, { flags: 'a' })
logStream.end(logfileText)
