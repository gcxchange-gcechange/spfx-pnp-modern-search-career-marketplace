/* eslint-disable no-constant-condition */
import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { useTheme, Link } from '@fluentui/react';
import { SelectLanguage } from './SelectLanguage';
import './CustomComponent.css';
import { Globals, Language } from './Globals';
import * as DOMPurify from 'dompurify';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {
    path?: string;
    applicationDeadlineDate?: string;
    cityEn?: string;
    cityFr?: string;
    classificationLevel?: string;
    classificationCodeEn?: string;
    classificationCodeFr?: string;
    contactEmail?: string;
    contactName?: string;
    contactObjectId?: string;
    durationEn?: string;
    durationFr?: string;
    jobDescriptionEn?: string;
    jobDescriptionFr?: string;
    jobTitleEn?: string;
    jobTitleFr?: string;
    jobType?: string;
    durationQuantity?: string;
    jobTypeTermSetGuid?: string;
    searchQuery?: string;
    applyEmail?: string;
}

interface ISearchMatch {
    startIndex: number;
}

const JobCardComponent: React.FC<ICustomComponentProps> = (props) => {

    const theme = useTheme();
    const strings = SelectLanguage(Globals.getLanguage());
    const lang = Globals.getLanguage();
    const jobId = props.path && props.path.split('ID=').length == 2  ? props.path.split('ID=')[1] : 'null';
    const jobUrl = `${Globals.jobOpportunityPageUrl}${jobId}`;
    let hightlightMatchesTitle: ISearchMatch[] = [];
    let hightlightMatchesDesc: ISearchMatch[] = [];

    // Translate the JobType terms
    const jobTypeIds = getTermIds(props.jobType);
    if (jobTypeIds && Globals.getJobTypes()) {
        const jobTypeLabels: string[] = [];
        for (let i = 0; i < jobTypeIds.length; i++) {
            jobTypeLabels.push(getJobTypeLabel(jobTypeIds[i], lang));
        }
        props.jobType = jobTypeLabels.join(', ');
    } else {
        console.warn('Unable to translate JobType... defaulting to display language.');
    }
    
    const getContactNameInitials = () => {
        if (props.contactName) {
            let nameSplit = props.contactName.toUpperCase().split(' ');
            return nameSplit[0] ? nameSplit[0][0] + (nameSplit[1] ? nameSplit[1][0] : '') : 'NA';
        }
        return 'NA';
    };

    const getApplicationDeadlineDate = () => {
        if (!props.applicationDeadlineDate) return 'N/A';

        const utcDate = new Date(props.applicationDeadlineDate);
        const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

        return utcDate.toLocaleString('en-US', { timeZone: userTimeZone });
    }

    // Fallback to default language incase we can't get the translations
    const termLabelDefaultLanguage = (value: string) => {
        try {
            if (value){
                let terms = [];
                let split = value.split(';GTSet');
                for (let i = 0; i < split.length - (split.length > 1 ? 1 : 0); i++) {
                    const parts = split[i].split('|');
                    terms.push(parts[parts.length - 1]);
                }
                return terms.join(', ');
            }
            return value;
        }
        catch (e) {
            console.log(e);
            return value;
        }
    }

    function getTermIds(terms: string): string[] | null {
        if (terms) {
            const termIds: string[] = [];

            const termsSplit = terms.split(/;(?=GP0)/g);
            for (let i = 0; i < termsSplit.length; i++) {
                const match = termsSplit[i].match(/#([0-9a-fA-F-]+)/);
                const termId = match ? match[1] : null;
                if (termId)
                    termIds.push(termId);
            }

            return termIds;
        }

        return null;
    }

    function getJobTypeLabel(termId: string, language: Language): string {
        try {
            const jobTypes: any[] = Globals.getJobTypes();
            for (let i = 0; i < jobTypes.length; i++) {
                if (jobTypes[i].id === termId) {
                    for (let n = 0; n < jobTypes[i].labels.length; n++) {
                        if (jobTypes[i].labels[n].languageTag === language) {
                            return jobTypes[i].labels[n].name;
                        }
                    }
                }
            }
        } catch (e) {
            console.error(`Unable to get JobType label for ${termId} - ${e}`)
        }
        return 'N/A';
    }

    interface IHighlightText {
        text: string;
        startIndicies: number[];
    }

    function highlightText(origText: string): IHighlightText {
        let retVal: IHighlightText = { text: origText, startIndicies: [] };

        try {
            const searchhWords = props.searchQuery.split('path:')[0].replace(/[*]/g, "").trim().split(/\s+/).filter(Boolean);

            if (searchhWords.length === 0)
                return retVal;

            const lowerOrigText = origText.toLowerCase();
            const matchIndices: Array<{ start: number; end: number }> = [];

            searchhWords.forEach(word => {
                const lowerWord = word.toLowerCase();
                let startIndex = 0;

                while (true) {
                    const index = lowerOrigText.indexOf(lowerWord, startIndex);
                    if (index === -1) 
                        break;

                    const isWordStart = index === 0 || !/[a-z0-9]/i.test(lowerOrigText[index - 1]);

                    // Only match when it starts with the word, since that's how our pnp search works.
                    if (isWordStart) {
                        matchIndices.push({ start: index, end: index + word.length });
                        retVal.startIndicies.push(index);
                    }

                    startIndex = index + 1;
                }
            });

            // Insert tags from right to left to avoid index shift
            matchIndices.sort((a, b) => b.start - a.start);

            matchIndices.forEach(({ start, end }) => {
                retVal.text = retVal.text.slice(0, end) + '</mark>' + retVal.text.slice(end);
                retVal.text = retVal.text.slice(0, start) + '<mark>' + retVal.text.slice(start);
            });
        } catch (e) {
            console.error(e);
        }

        return retVal;
    }

    const isExpired = ():boolean => {
        if (props.applicationDeadlineDate) {
            if (new Date() >= new Date(`${props.applicationDeadlineDate.toString()} UTC`))
                return true;
            else
                return false;
        }
        return true;
    }

    const expired = isExpired();

    const transformedTitle = highlightText(lang === Language.French ? props.jobTitleFr : props.jobTitleEn);
    transformedTitle.startIndicies = transformedTitle.startIndicies.filter(match => match <= 25);

    const transformedDescription = highlightText(lang === Language.French ? props.jobDescriptionFr : props.jobDescriptionEn);
    transformedDescription.startIndicies = transformedDescription.startIndicies.filter(match => match <= 199);

    return (
        <Link 
            href={jobUrl} 
            target='_blank'
            className='noLinkStyle'
            id={'jobView-'+ jobId}
        >
            <div 
                className={expired ? 'jobcard expiredJobCard' : 'jobcard'}
                style={{
                    border: `1px solid ${theme.palette.themePrimary}`,
                }}
            >
                {expired && 
                    <div className='expiredBanner'>
                        <p role="status" aria-live="polite">{strings.jobExpired}</p>
                    </div>
                }
                <div className="card-content">
                    <h3 style={{
                            color: `${theme.palette.themePrimary}`,
                            overflow: 'hidden',
                            maxWidth: '350px'
                        }}
                    >
                        <span dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(transformedTitle.text) }} />
                    </h3>
                    <div className="sub">
                        { props.searchQuery && props.searchQuery.indexOf('* path:') !== 0 && hightlightMatchesTitle.length === 0 && hightlightMatchesDesc.length === 0 &&
                            <div className="searchTermFound">
                                <mark><b>{strings.searchTermFound}</b></mark>
                            </div>
                        }
                        <div>
                            <b>{strings.classificationLevel}</b>: {lang === Language.French ? props.classificationCodeFr : props.classificationCodeEn}-{props.classificationLevel}
                        </div>
                        <div>
                            <b>{strings.opportunityType}</b>: {termLabelDefaultLanguage(props.jobType)}
                        </div>
                        <div>
                            <b>{strings.duration}</b>: {!props.durationQuantity ? strings.undetermined : `${props.durationQuantity} ${lang === Language.French ? props.durationFr : props.durationEn}`}
                        </div>
                    </div>
                    <div className="description">
                        <b>{strings.description}</b>: <span dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(transformedDescription.text) }} /> 
                    </div>
                    <div className="sub">
                        <div>
                            <b>{strings.location}</b>: {lang === Language.French ? (props.cityFr ? props.cityFr : strings.remote) : (props.cityEn ? props.cityEn : strings.remote)}
                        </div>
                        <div>
                            <b>{strings.deadline}</b>: {getApplicationDeadlineDate()}
                        </div>
                    </div>
                    <div className="contact">
                        <div className="profile">
                            <div>
                                {getContactNameInitials()}
                            </div>
                        </div>
                        <div className="info">
                            <div>{props.contactName}</div>
                            <div>{props.applyEmail? props.applyEmail : props.contactEmail}</div>
                        </div>
                    </div>
                </div>
            </div>
        </Link>
    );
};

export class MyCustomComponentWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {
        let props = {} as any;

        const getAttr = (name: string) => this.getAttribute(name) ?? undefined;

        props.path = getAttr("path") || "";
        props.applicationDeadlineDate = getAttr("application-deadline-date");
        props.cityEn = getAttr("city-en");
        props.cityFr = getAttr("city-fr");
        props.classificationLevel = getAttr("classification-level");
        props.classificationCodeEn = getAttr("classification-code-en");
        props.classificationCodeFr = getAttr("classification-code-fr");
        props.contactEmail = getAttr("contact-email");
        props.contactName = getAttr("contact-name");
        props.contactObjectId = getAttr("contact-object-id");
        props.durationEn = getAttr("duration-en");
        props.durationFr = getAttr("duration-fr");
        props.jobDescriptionEn = getAttr("job-description-en");
        props.jobDescriptionFr = getAttr("job-description-fr");
        props.jobTitleEn = getAttr("job-title-en");
        props.jobTitleFr = getAttr("job-title-fr");
        props.jobType = getAttr("job-type");
        props.durationQuantity = getAttr("duration-quantity");
        props.jobTypeTermSetGuid = getAttr("job-type-term-set-guid");
        props.searchQuery = getAttr("search-query");
        props.applyEmail = getAttr("apply-email");

        if (props.applicationDeadlineDate) {
            const d = new Date(props.applicationDeadlineDate);
            props.applicationDeadlineDate = isNaN(d.getTime()) ? undefined : d;
        }

        const JobCard = <JobCardComponent {...props} />;
        ReactDOM.render(JobCard, this);
    }    

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}