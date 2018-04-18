// tslint:disable-next-line:no-var-requires
const jsZip = require('jszip');

// tslint:disable-next-line:no-var-requires
const docxtemplater = require('docxtemplater');

import { IDocFusionOptions } from './fusion-options';
import * as _ from 'lodash';
import * as fs from 'fs';
import * as util from 'util';
import * as path from 'path';

export class DocFusion {
    private static cpt = 0;
    private static mergeOptions(o1: IDocFusionOptions, o2: IDocFusionOptions) {
        const o3 = _.clone(o1);
        _.merge(o3, o2);
        return o3;
    }
    constructor(public readonly fusionOptions: IDocFusionOptions) {

    }

    public async generateDoc(options: IDocFusionOptions, data: any): Promise<string> {
        const write = util.promisify(fs.writeFile);
        const read = util.promisify(fs.readFile);
        const exist = util.promisify(fs.exists);
        const finalOptions = DocFusion.mergeOptions(this.fusionOptions, options);
        const inputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.inputFileName);
        let outputFileName = path.resolve(finalOptions.generationDirectory, finalOptions.outputFileName);

        if (!await exist(finalOptions.generationDirectory)) {
            // tslint:disable-next-line:no-console
            console.log('Directory does not exist!');
        } else {
            if (await !exist(inputFileName)) {
                // tslint:disable-next-line:no-console
                console.log('File does not exist!');
            } else {
                const content = await read(inputFileName, 'binary');
                // tslint:disable-next-line:no-console
                // const file = 'file.txt';
                const zip = new jsZip(content);
                const doc = new docxtemplater();

                doc.loadZip(zip);

                // set the templateVariables
                doc.setData(data);

                // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
                doc.render();
                const buf = doc.getZip().generate({ type: 'nodebuffer' });

                if (outputFileName === '' || outputFileName === undefined) {
                    outputFileName = this.generateRndmName(DocFusion.cpt);
                    // tslint:disable-next-line:no-console
                    console.log('File named!');
                }
                // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
                await write(outputFileName, buf);
                // tslint:disable-next-line:no-console
                console.log('Fusion done!');

                return outputFileName;
            }
        }
    }

    public  generateRndmName(compteur: number): string {
        return 'file_fusionned' + (compteur + 1) + '.docx';
    }
}

let optionfusion = {
    generationDirectory: 'C:\\Users\\raker\\Desktop',
    inputFileName: 'Document.docx',
    outputFileName: 'doc_output.docx'};

let dataDoc = {
    last_name: 'Kermad',
    first_name: 'Ramzy',
    phone: '06-06-06-06-06',
    description: 'Stage dev Javascript/ Typescript'};

let d: DocFusion = new DocFusion({});

let fusionned = d.generateDoc(optionfusion, dataDoc);

// tslint:disable-next-line:no-console
console.log('doc fusionned');
