"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const _ = require("lodash");
const fs = require("fs");
const path = require("path");
const jsZip = require('jszip');
const Docxtemplater = require('docxtemplater');
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
        const finalOptions = DocFusion.mergeOptions(this.fusionOptions, options);
        const inputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.inputFileName);
        const outputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.outputFileName);
        const content = fs
            .readFileSync(inputFileName, 'binary');
        const zip = new jsZip(content);
        const doc = new Docxtemplater();
        doc.loadZip(zip);
        doc.setData(data);
        doc.render();
        const buf = doc.getZip()
            .generate({ type: 'nodebuffer' });
        fs.writeFileSync(outputFileName, buf);
        return outputFileName;
    }
}
exports.DocFusion = DocFusion;
//# sourceMappingURL=doc-fusion.js.map