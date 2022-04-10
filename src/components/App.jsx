import React from "react";
import { read, utils } from "xlsx";
import DocxTemplater from "docxtemplater";
import PizZip from "pizzip";
import PizZipUtils from "pizzip/utils/index.js";
import { saveAs } from "file-saver";
import DocxMerger from "docx-merger";
global.Buffer = global.Buffer || require("buffer").Buffer;

let dateColumns = [9];
let currencyColumns = [5, 6];
let fileNameColumn = 4;

function pause(msec) {
    return new Promise((resolve, reject) => {
        setTimeout(resolve, msec || 1000);
    });
}

function getDate(date) {
    return new Date(-2209075200000 + (date - (date < 61 ? 0 : 1)) * 86400000).toLocaleDateString("en-IN");
}

function App() {
    return (
        <div>
            Template File: <input type="file" id="template-file" accept=".docx" onChange={getTemplateFile} />
            <br />
            <br />
            Data File: <input type="file" id="data-file" accept=".xlsx" onChange={createDocFiles} />
            <br />
            <br />
            File Merge:{" "}
            <input type="file" id="merge-file" multiple="multiple" accept=".docx" onChange={mergeDocFiles} />
        </div>
    );
}
export default App;

function getTemplateFile(e) {
    // let files = e.target.files;
    // let allFiles = [];
    // for (let i = 0; i < files.length; i++) {
    //     let reader = new FileReader();
    //     reader.readAsBinaryString(files[i]);
    //     allFiles.push(reader);
    // }
    // console.log(`🚀 -------------------------------------------------------------`);
    // console.log(`🚀 ~ file: App.jsx ~ line 78 ~ mergeDocFiles ~ files`, files);
    // console.log(`🚀 -------------------------------------------------------------`);
    // let docx = new DocxMerger({}, files);
    // docx.save("blob", (data) => saveAs(data, "output.docx"));
}

function createDocFiles(e) {
    const [file] = e.target.files;

    const reader = new FileReader();
    reader.readAsBinaryString(file);
    reader.onload = (evt) => {
        const bstr = evt.target.result;
        const wb = read(bstr, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        let data = utils.sheet_to_json(ws, { header: 1 });
        data = data.map((e) => {
            dateColumns.forEach((i) => e.length > i && (e[i] = getDate(e[i])));
            currencyColumns.forEach((i) => (e[i] = e.length > i && e[i].toLocaleString("en-IN")));
            return e;
        });
        console.log(`🚀 --------------------------------------`);
        console.log(`🚀 ~ file: App.jsx ~ line 34 ~ data`, data);
        console.log(`🚀 --------------------------------------`);

        let allFiles = [];
        PizZipUtils.getBinaryContent("./Template.docx", async (error, content) => {
            if (error) {
                throw error;
            }
            for (let i = 0; i < data.length; i++) {
                try {
                    const zip = new PizZip(content);
                    const doc = new DocxTemplater(zip, { paragraphLoop: true, linebreaks: true });
                    doc.render(data[i]);
                    // allFiles.push({
                    //     doc,
                    //     processedDoc: doc.getZip().generate({
                    //         type: "blob",
                    //         mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    //     }),
                    //     fileName: `${data[i][fileNameColumn]}.docx`,
                    // });
                    allFiles.push(
                        doc.getZip().generate({
                            type: "arraybuffer",
                            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        })
                    );
                } catch (e) {
                    console.log(e.message);
                }
            }

            let docx = new DocxMerger({}, allFiles);
            docx.save("blob", (data) => saveAs(data, "MergedDocuments.docx"));

            // for (let i = 0; i < allFiles.length; i++) {
            //     saveAs(allFiles[i].processedDoc, allFiles[i].fileName);
            //     if (i % 10 === 0) {
            //         await pause(1050);
            //     }
            // }
            console.log(`🚀 ~ file: App.jsx ~ line 104 ~ Done`);
        });
    };
}

async function mergeDocFiles(e) {
    try {
        let files = e.target.files || [];
        let allFiles = [];
        const filePromises = Array.from(files).map((file) => {
            return new Promise((resolve, reject) => {
                let reader = new FileReader();
                reader.readAsBinaryString(file);
                reader.onload = (evt) => {
                    const bstr = evt.target.result;
                    allFiles.push(bstr);
                    resolve(bstr);
                };
            });
        });
        await Promise.all(filePromises);
        let docx = new DocxMerger({}, allFiles);
        docx.save("blob", (data) => saveAs(data, "output.docx"));
    } catch (error) {
        console.log(`🚀 ------------------------------------------------------------`);
        console.log(`🚀 ~ file: App.jsx ~ line 123 ~ mergeDocFiles ~ error`, error);
        console.log(`🚀 ------------------------------------------------------------`);
    }
}
