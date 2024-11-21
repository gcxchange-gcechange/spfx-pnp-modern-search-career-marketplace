import * as strings from 'MyCompanyLibraryLibraryStrings';
const english = require("./myCompanyLibrary/loc/en-us.js")
const french = require("./myCompanyLibrary/loc/fr-fr.js")

export function SelectLanguage(lang: string): IMyCompanyLibraryLibraryStrings  {
    switch (lang.toLowerCase()) {
        case "en": {
            return english;
        }
        case "fr": {
            return french;
        }
        default: {
            return strings;
        }
    }
}