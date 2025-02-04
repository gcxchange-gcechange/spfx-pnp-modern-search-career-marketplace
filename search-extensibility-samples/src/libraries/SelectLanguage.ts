import { Language } from './Globals';
const english = require("./myCompanyLibrary/loc/en-us.js")
const french = require("./myCompanyLibrary/loc/fr-fr.js")

export function SelectLanguage(lang: Language): IMyCompanyLibraryLibraryStrings  {
    switch (lang) {
        case Language.English: {
            return english;
        }
        case Language.French: {
            return french;
        }
        default: {
            return english;
        }
    }
}