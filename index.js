import * as XLSX from 'xlsx/xlsx.mjs'
import * as fs from 'fs/promises'
import * as cli from 'cli-progress'
import { col, categories, dataStartRow } from './references.js'

const run = async () => {
    const path = process.argv[2]
    if (!path.length) {
        console.log("Missing file path to cutsheet collection")
        process.exit(1)
    }

    console.log(`Beginning scan of cutsheet files in location: ${process.argv[2]}`)
    const cutsheets = []
    const startTime = new Date().getTime()
    try {
        const dir = await fs.opendir(path);
        for await (const file of dir) {
            const cs = await openCutsheet(file)
            if (cs) cutsheets.push(cs)
        }
    } catch (err) {
        console.error(err);
    } finally {
        const seconds = (new Date().getTime() - startTime) / 1000
        console.log(`Complete. Scanned ${cutsheets.length} cutsheets in ${seconds} seconds.`)
    }
}

const openCutsheet = async (file) => {
    if (!checkIsXlsx(file)) {
        return
    }
    if (!(await checkIsCutsheet(file))) {
        return
    }
    let cutsheet = createCutsheetObject(file)
    return cutsheet
}

const checkIsXlsx = (file) => {
    if (!file.isFile()) return false
    if (file.name.slice(file.name.lastIndexOf('.') + 1).toLowerCase() != 'xlsx') return false
    return true
}

const checkIsCutsheet = async (file) => {
    let check = false;
    try {
        const workbook = await createWorkbook(file)
        if (workbook.Sheets['Cut List'].A1.v == 'CrateMaker')
            check = true
    } finally {
        return check
    }
}

const createCutsheetObject = async (file) => {
    let cutsheet = {}
    const bar = new cli.SingleBar({
        format: '|{bar}| {percentage}% || ' + file.name,
        hideCursor: true,
        barCompleteChar: '\u2588',
        barIncompleteChar: '\u2591',
    })
    bar.start(100, 0)

    try {
        const workbook = await createWorkbook(file)
        const sheet = workbook.Sheets['Cut List']
        cutsheet = await parseCutsheet(sheet, (x) => bar.increment(x))
    } catch (err) {
        console.log(err)
    } finally {
        bar.update(100)
        bar.stop()
    }
    return cutsheet
}

const parseCutsheet = async (sheet, bar) => {
    const cutsheet = {
        bomLines: 0,
        lumberCount: 0,
        lumberLength: 0,
        plyCount: 0,
        plySquareFeet: 0,
        foamCount: 0,
        otherCount: 0,
    }
    cutsheet.bomLines = getBomLines(sheet)
    const increment = 100 / cutsheet.bomLines
    for (let i = dataStartRow; i < cutsheet.bomLines + dataStartRow; i++) {
        const category = getCategory(sheet[cell(col.pn, i)].v)
        switch (category) {
            case 'lumber':
                cutsheet.lumberCount += sheet[cell(col.qty, i)].v
                cutsheet.lumberLength += sheet[cell(col.length, i)].v
                break
            case 'ply':
                cutsheet.plyCount += sheet[cell(col.qty, i)].v
                cutsheet.plySquareFeet += Math.round(sheet[cell(col.length, i)].v * sheet[cell(col.width, i)].v / 144, 2)
                break
            case 'foam':
                cutsheet.foamCount += sheet[cell(col.qty, i)].v
                break
            default:
                cutsheet.otherCount += sheet[cell(col.qty, i)].v
        }
        await sleep(3) // just to make it look cool
        bar(increment)
    }
    return cutsheet
}

const getBomLines = (sheet) => {
    let i = dataStartRow
    while (sheet[cell(0, i)]) i++
    return i - dataStartRow
}

const getCategory = (pn) => {
    return categories[pn.slice(0, 3)]
}

const createWorkbook = async (file) => {
    const data = await fs.readFile(process.argv[2].concat(file.name))
    return await XLSX.read(data)
}

const cell = (c, r) => {
    return XLSX.utils.encode_cell({ c: c, r: r })
}

function sleep(ms) {
    return new Promise((resolve) => {
        setTimeout(resolve, ms)
    })
}

run()
