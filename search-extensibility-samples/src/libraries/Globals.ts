export enum Language {
    English = 'en-US',
    French = 'fr-FR'
}

export class Globals {
    public static jobOpportunityPageUrl: string;
    public static userDisplayName: string;
    public static userEmail: string;
    public static searchQuery: string;
    public static tenant: string;
    private static _language: Language;
    private static _jobTypes: string[];
    private static _newsSearchLayout: boolean;

    public static getLanguage(): Language {
        return this._language;
    }

    public static setLanguage(lang: string): void {
        if (lang) {
            lang = lang.toLowerCase();
            if (lang === Language.English || lang === 'en') {
                this._language = Language.English;
            }
            else if (lang === Language.French || lang === 'fr') {
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

    public static getNewsSearchLayout(): boolean {
        return this._newsSearchLayout;
    }

    public static setNewsSearchLayout(state: boolean): void {
        this._newsSearchLayout = state;
    }
}