const XLSX = require('xlsx');
const parseString = require('xml2js').parseString;
const AdmZip = require('adm-zip');

class TestService {

    static async getObjectFromXml(xml) {
        return new Promise( (resolve, reject) => {
            parseString(xml, function (err, result) {
                if (err !== null) reject(err)
                else resolve(result);
            })
        })
    }

    static rowConvert(xmlObject) {

        return xmlObject;
    }

    static async data(params) {

        let filepath = "/private/var/www/portal.js/backend/tmp/111.xlsx";

        let workbook = XLSX.readFile(filepath);

        let Sheets = {};
        let SheetsInfo = {};
        let zip = new AdmZip(filepath);

        let styleFile = {};

        let zipEntries = zip.getEntries();
        for (let key in zipEntries) {
            let zipEntry = zipEntries[key];
            if (zipEntry.entryName.includes("xl/worksheets/sheet")) {
                const regex = /sheet(\d+)/gm;
                let name = zipEntry.name;
                let indexFind = [...name.matchAll(regex)];
                let indexSheet = indexFind[0][1];

                let xml = zipEntry.getData().toString("utf8");
                let xmlObject = await this.getObjectFromXml(xml);

                let sheetInfo = {};
                if (xmlObject.worksheet.sheetFormatPr !== undefined) {
                    sheetInfo["sheetFormatPr"] = xmlObject.worksheet.sheetFormatPr[0]["$"];
                }
                if (xmlObject.worksheet.cols !== undefined) {
                    sheetInfo["cols"] = xmlObject.worksheet.cols[0].col;
                }

                SheetsInfo[indexSheet] = sheetInfo;
                Sheets[indexSheet] = xmlObject.worksheet.sheetData[0];
            }
            else if (zipEntry.entryName.includes("xl/style")) {
                let xml = zipEntry.getData().toString("utf8");
                styleFile = await this.getObjectFromXml(xml);
            }
        }


        let SheetsRaw = {};
        let SheetsStyle = {};
        for (let key in workbook.SheetNames) {
            let nameSheet = workbook.SheetNames[key];
            let index = parseInt(key)+1;
            if (SheetsInfo[index] !== undefined) {
                SheetsStyle[nameSheet] = SheetsInfo[index];
                SheetsRaw[nameSheet] = Sheets[index];
            } else {
                SheetsStyle[nameSheet] = {};
                SheetsRaw[nameSheet] = {};
            }
        }
        SheetsStyle.gStyle = styleFile;

        return {SheetNames: workbook.SheetNames, Sheets:workbook.Sheets, SheetsRaw: SheetsRaw, SheetsStyle: SheetsStyle};
    }

}

module.exports = TestService;
