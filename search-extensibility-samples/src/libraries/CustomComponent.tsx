import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PrimaryButton, DefaultButton, useTheme } from '@fluentui/react';
import { SelectLanguage } from './SelectLanguage';
import './CustomComponent.css';
import { Globals, Language } from './Globals';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {
    path?: string;
    applicationDeadlineDate?: Date;
    cityEn?: string;
    cityFr?: string;
    classificationLevel?: string;
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
}

const JobCardComponent: React.FC<ICustomComponentProps> = (props) => {

    const theme = useTheme();
    const strings = SelectLanguage(Globals.getLanguage());
    const lang = Globals.getLanguage();

    // Translate the JobType terms
    const jobTypeIds = getTermIds(props.jobType);
    if (jobTypeIds) {
        const jobTypeLabels: string[] = [];
        for (let i = 0; i < jobTypeIds.length; i++) {
            jobTypeLabels.push(getJobTypeLabel(jobTypeIds[i], lang));
        }
        props.jobType = jobTypeLabels.join(', ');
    }
    
    const getContactNameInitials = () => {
        if (props.contactName) {
            let nameSplit = props.contactName.toUpperCase().split(' ');
            return nameSplit[0] ? nameSplit[0][0] + (nameSplit[1] ? nameSplit[1][0] : '') : 'NA';
        }
        return 'NA';
    };

    const handleViewClick = () => {
        if (props.path && props.path.split('ID=').length == 2) 
            window.open(`${Globals.jobOpportunityPageUrl}${props.path.split('ID=')[1]}`, '_blank');
    };

    const handleApplyClick = () => {
        window.location.href = `mailto:${props.contactEmail}?subject=${lang === Language.French ? props.jobTitleFr : props.jobTitleEn}`;
    };

    const getApplicationDeadlineDate = () => {
        if (props.applicationDeadlineDate) {
            const utcDate = new Date(`${props.applicationDeadlineDate.toString()} UTC`);
            const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone; 
            return utcDate.toLocaleString('en-US', { timeZone: userTimeZone });
        }
        return 'N/A';
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

    return (
        <div 
            className="jobcard" 
            style={{
                border: `1px solid ${theme.palette.themePrimary}`,
            }}
        >
            <div className="card-content">
                <h3 style={{
                        color: `${theme.palette.themePrimary}`,
                    }}
                >
                    {lang === Language.French ? props.jobTitleFr : props.jobTitleEn}
                </h3>
                <div className="sub">
                    <div>{strings.classificationLevel}: {props.classificationLevel}</div>
                    <div>{strings.opportunityType}: {termLabelDefaultLanguage(props.jobType)}</div>
                    <div>{strings.duration}: {props.durationQuantity} {lang === Language.French ? props.durationFr : props.durationEn}</div>
                </div>
                <div className="description">
                    {lang === Language.French ? props.jobDescriptionFr : props.jobDescriptionEn}
                </div>
                <div className="sub bold">
                    <div>{strings.location}: {lang === Language.French ? props.cityFr : props.cityEn}</div>
                    <div>{strings.deadline}: {getApplicationDeadlineDate()}</div>
                </div>
                <div className="contact">
                    <div className="profile">
                        <div>
                            {getContactNameInitials()}
                        </div>
                    </div>
                    <div className="info">
                        <div>{props.contactName}</div>
                        <div>{props.contactEmail}</div>
                    </div>
                </div>
                <div className="actions">
                    <DefaultButton 
                        text={strings.view} 
                        aria-label={strings.viewAria}
                        onClick={handleViewClick}
                    />
                    <PrimaryButton 
                        text={strings.apply}
                        aria-label={strings.applyAria}
                        onClick={handleApplyClick} 
                    />
                </div>
            </div>
        </div>
    );
};

export class MyCustomComponentWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        const JobCard = <JobCardComponent {...props} />;
        ReactDOM.render(JobCard, this);
    }    

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}