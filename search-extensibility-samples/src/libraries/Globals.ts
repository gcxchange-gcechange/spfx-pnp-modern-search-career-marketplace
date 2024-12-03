export enum Language {
    English = 'en',
    French = 'fr'
}

export class Globals {
    private static _language: string;

    public static getLanguage(): string {
        return this._language;
    }

    public static setLanguage(lang: string): void {
        if (lang) {
            lang = lang.toLowerCase();
            if (lang === Language.English || lang === Language.French) {
                this._language = lang;
                return;
            }
            this._language = 'en';
        }
    }
}