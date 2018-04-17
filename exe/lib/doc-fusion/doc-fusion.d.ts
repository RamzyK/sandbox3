import { IDocFusionOptions } from './fusion-options';
export declare class DocFusion {
    readonly fusionOptions: IDocFusionOptions;
    private static cpt;
    private static mergeOptions(o1, o2);
    constructor(fusionOptions: IDocFusionOptions);
    generateDoc(options: IDocFusionOptions, data: any): Promise<string>;
    generateRndmName(compteur: number): string;
}
