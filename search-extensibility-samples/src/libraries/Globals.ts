export enum Language {
    English = 'en-US',
    French = 'fr-FR'
}

export class Globals {
    private static _language: string;
    public static jobOpportunityPageUrl: string;

    public static getLanguage(): string {
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
}