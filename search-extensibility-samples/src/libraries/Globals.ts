export enum Language {
    English = 'en-US',
    French = 'fr-FR'
}

export class Globals {
    private static _language: Language;
    public static jobOpportunityPageUrl: string;
    private static _jobTypes: string[];

    public static getLanguage(): Language {
        return this._language;
    }

    public static setLanguage(lang: string): void {
        if (lang) {
            lang = lang.toLowerCase();
            if (lang === 'en') {
                this._language = Language.English;
            }
            else if (lang === 'fr') {
                this._language = Language.French;
            }
            else {
                this._language = Language.English;
            }
        }
    }

    public static getJobTypes(): string[] {
        return this._jobTypes;
    }

    public static setJobTypes(jobTypes: string[]): void {
        if (jobTypes)
            this._jobTypes = jobTypes;
    }
}