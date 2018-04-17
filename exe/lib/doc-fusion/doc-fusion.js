"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const jsZip = require('jszip');
const docxtemplater = require('docxtemplater');
const _ = require("lodash");
const fs = require("fs");
const util = require("util");
const path = require("path");
class DocFusion {
    constructor(fusionOptions) {
        this.fusionOptions = fusionOptions;
    }
    static mergeOptions(o1, o2) {
        const o3 = _.clone(o1);
        _.merge(o3, o2);
        return o3;
    }
    async generateDoc(options, data) {
        const write = util.promisify(fs.writeFile);
        const read = util.promisify(fs.readFile);
        const exist = util.promisify(fs.exists);
        const finalOptions = DocFusion.mergeOptions(this.fusionOptions, options);
        const inputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.inputFileName);
        let outputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.outputFileName);
        if (!await exist(finalOptions.generationDirectory)) {
            throw Error('Directory does not exist!');
        }
        else {
            if (await !exist(inputFileName)) {
                throw Error(inputFileName + ': File does not exist!');
            }
            else {
                const content = await read(inputFileName, 'binary');
                console.log('CONTENT: ', content);
                const zip = new jsZip(content);
                const doc = new docxtemplater();
                doc.loadZip(zip);
                doc.setData(data);
                doc.render();
                const buf = doc.getZip().generate({ type: 'nodebuffer' });
                if (outputFileName === '' || outputFileName === undefined) {
                    outputFileName = this.generateRndmName(DocFusion.cpt);
                }
                await write(outputFileName, buf);
                console.log('Done!');
                return outputFileName;
            }
        }
    }
    generateRndmName(compteur) {
        return 'file_fusionned' + (compteur + 1) + '.docx';
    }
}
DocFusion.cpt = 0;
exports.DocFusion = DocFusion;
//# sourceMappingURL=doc-fusion.js.map