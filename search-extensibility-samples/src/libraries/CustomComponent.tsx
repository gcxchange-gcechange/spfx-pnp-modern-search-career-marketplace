import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PrimaryButton, DefaultButton, useTheme, Link } from '@fluentui/react';
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
}

const JobCardComponent: React.FC<ICustomComponentProps> = (props) => {

    const theme = useTheme();
    const strings = SelectLanguage(Globals.getLanguage());
    const lang = Globals.getLanguage();
    const jobId = props.path && props.path.split('ID=').length == 2  ? props.path.split('ID=')[1] : 'null';
    const jobUrl = `${Globals.jobOpportunityPageUrl}${jobId}`;

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

    const formatName = (displayName: string): string => {
        let formattedName = displayName;

        if (displayName) 
            formattedName = displayName.split(',', 2).reverse().join(' ').replace(',', '');

        return formattedName;
    }

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

    const mailApplyBodyEn = encodeURIComponent(`Hello ${formatName(props.contactName)},\n\nI hope this message finds you well. My name is ${Globals.userDisplayName}, and I am interested in the career opportunity you posted on the GCXchange Career Marketplace. Please find my resumé attached for your review.\n\nI would appreciate the opportunity to discuss how my skills align with your needs.\nThank you for your time and consideration.\n\nBest regards,\n${formatName(Globals.userDisplayName)}`);
    const mailApplyBodyFr = encodeURIComponent(`Bonjour ${formatName(props.contactName)},\n\nJ\’espère que vous allez bien. Mon nom est ${Globals.userDisplayName} et l\’offre d\’emploi que vous avez publiée dans le Carrefour d\’emploi sur GCÉchange m\’intéresse. Vous trouverez ci joint mon curriculum vitæ.\n\nMes compétences semblent correspondre à vos besoins et j\’aimerais en discuter avec vous.\nJe vous remercie de prendre le temps de considérer ma candidature.\n\nCordialement,\n${formatName(Globals.userDisplayName)}`);

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
                        overflow: 'hidden',
                        maxWidth: '350px'
                    }}
                >
                    <a href="#" onClick={(e) => {
                        e.preventDefault();
                        window.open(`${Globals.jobOpportunityPageUrl}${props.path.split('ID=')[1]}`, '_blank', 'noopener,noreferrer');
                    }}>
                        {lang === Language.French ? props.jobTitleFr : props.jobTitleEn}
                    </a>
                </h3>
                <div className="sub">
                    <div>
                        <b>{strings.classificationLevel}</b>: {lang === Language.French ? props.classificationCodeFr : props.classificationCodeEn}-{props.classificationLevel}
                    </div>
                    <div>
                        <b>{strings.opportunityType}</b>: {termLabelDefaultLanguage(props.jobType)}
                    </div>
                    <div>
                        <b>{strings.duration}</b>: {props.durationQuantity} {lang === Language.French ? props.durationFr : props.durationEn}
                    </div>
                </div>
                <div className="description">
                    <b>{strings.description}</b>: {lang === Language.French ? props.jobDescriptionFr : props.jobDescriptionEn}
                </div>
                <div className="sub">
                    <div>
                        <b>{strings.location}</b>: {lang === Language.French ? props.cityFr : props.cityEn}
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
                        <div>{props.contactEmail}</div>
                    </div>
                </div>
                <div className="actions">
                    <Link 
                        href={jobUrl} 
                        target='_blank'
                    >
                        <DefaultButton 
                            id={'jobView-'+ jobId}
                            aria-label={strings.viewAria + (lang === Language.French ? props.jobTitleFr : props.jobTitleEn)}
                            text={strings.view}
                        />
                    </Link>
                    <Link 
                        href={`mailto:${props.contactEmail}?subject=${lang === Language.French ? `Intérêt pour l'opportunité ${props.jobTitleFr}` : `Interested in the ${props.jobTitleEn} opportunity`}&body=${lang === Language.French ? mailApplyBodyFr : mailApplyBodyEn}&JobOpportunityId=${jobId}`}
                        target='_blank'
                    >
                        <PrimaryButton 
                            id={'jobApply-'+ jobId}
                            aria-label={strings.applyAria + (lang === Language.French ? props.jobTitleFr : props.jobTitleEn)}
                            text={strings.apply}
                        />
                    </Link>
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