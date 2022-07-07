import * as XLSX from 'xlsx/xlsx.mjs'
import * as fs from 'fs/promises'
import * as cli from 'cli-progress'

const run = async () => {
    const path = process.argv[2]
    if (!path.length) {
        console.log("Missing file path to cutsheet collection")
        process.exit(1)
    }

    try {
        const dir = await fs.opendir(path);
        let cutsheets = []
        for await (const file of dir) {
            let cs = await parseCutsheet(file)
            if (cs) cutsheets.push(cs)
        }
    } catch (err) {
        console.error(err);
    }
}

const parseCutsheet = async (file) => {
    if (!checkIsXlsx(file)) {
        return
    }
    if (!(await checkIsCutsheet(file))) {
        return
    }
    let cutsheet = {}
    createCutsheetObject(file)
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
        let data = await fs.readFile(process.argv[2].concat(file.name))
        let workbook = await XLSX.read(data)
        if (workbook.Sheets['Cut List'].A1.v == 'CrateMaker')
            check = true
    } finally {
        return check
    }
}

const createCutsheetObject = async (file) => {
    const cutsheet = {}
    const bar = new cli.SingleBar({ format: '|{bar}| {percentage}% :: ' + file.name, hideCursor: true })
    bar.start(100, 0)
    for (let i = 0; i < 10; i++) {
        bar.increment(10)
        await sleep(500)
    }
    bar.stop()
}

function sleep(ms) {
    return new Promise((resolve) => {
        setTimeout(resolve, ms)
    })
}

run()


