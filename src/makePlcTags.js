const XLSX = require('xlsx')
const { XMLParser, XMLBuilder } = require('fast-xml-parser')
const { writeFile } = require('fs/promises')

const PATH_TO_DATA = 'data'
const TAG_TABLE_NAME = 'Imported SPS-Liste'

console.log(`Started ./src/makePlcTags.js, Data is in ${PATH_TO_DATA}`)

// --
// 1) Read SPS-Liste.xlsx - use npm xlsx
// 2) Parse rows
// 3) Filter column E "SPSAdresse" not empty
// 4) Build Tags according to PLCTags.example.xml
// 5) Write new XML file PlcTags.xml (override existing) - use fast-xml-parser
// --

// 1) Read SPS-Liste.xlsx
const workbook = XLSX.readFile(`${PATH_TO_DATA}/SPS-Liste.xlsx`)

// Get first worksheet
const sheetName = workbook.SheetNames[0]
console.log(`Sheet name: ${sheetName}`)
const sheet = workbook.Sheets[sheetName]

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
const plcTagsFiltered = plcTags.filter((tag, index, self) => {
    const ioAddress = tag.ioAddress
    const firstIndex = plcTags.findIndex(
        (t, newindex) => t.ioAddress === ioAddress && newindex !== index
    )
    return firstIndex !== -1
})

console.log(`IO - Adresses with more then one definition:`)
plcTagsFiltered.forEach((e) =>
    console.log(`Adresse: ${e.ioAddress}; SPS-Karte: ${e.tagIdPlcModule}`)
)

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
writeFile(`${PATH_TO_DATA}/PlcTags.xml`, output)
    .then(() => {
        console.log('PlcTags.xml written')
    })
    .catch((err) => {
        console.error(err)
    })
