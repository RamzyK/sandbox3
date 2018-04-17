import { IDocFusionOptions } from './fusion-options';
import * as _ from 'lodash';
import * as fs from 'fs';
import * as util from 'util';
import * as path from 'path';

// tslint:disable-next-line:no-var-requires
const jsZip = require('jszip');

// tslint:disable-next-line:variable-name
const Docxtemplater
    // tslint:disable-next-line:no-var-requires
    = require('docxtemplater');

export class DocFusion {
    private static mergeOptions(o1: IDocFusionOptions, o2: IDocFusionOptions) {
        const o3 = _.clone(o1);
        _.merge(o3, o2);
        return o3;
    }
    constructor(public readonly fusionOptions: IDocFusionOptions) {
    }

    public async generateDoc(options: IDocFusionOptions, data: any): Promise<string> {
        const finalOptions = DocFusion.mergeOptions(this.fusionOptions, options);
        const inputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.inputFileName);
        const outputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.outputFileName);
        const content = fs
            .readFileSync(inputFileName, 'binary');

        const zip = new jsZip(content);

        const doc = new Docxtemplater();
        doc.loadZip(zip);

        // set the templateVariables
        doc.setData(data);

        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render();
        const buf = doc.getZip()
            .generate({ type: 'nodebuffer' });

        // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
        fs.writeFileSync(outputFileName, buf);
        return outputFileName;
    }
}
