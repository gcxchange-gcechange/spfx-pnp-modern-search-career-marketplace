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

    async function translateTerm(term: string, termSetGuid: string, elementId: string) {
        if (term && props.jobTypeTermSetGuid && lang !== Language.English) {
            const termsSplit = term.split(/;(?=GP0)/g);
    
            const fetchPromises = termsSplit.map(async (term) => {
                const match = term.match(/#([0-9a-fA-F-]+)/);
                const termId = match ? match[1] : null;
    
                if (!termId) return null;
    
                try {
                    const response = await fetch(`/_api/v2.1/termstore/sets/${termSetGuid}/terms/${termId}`, {
                        method: 'GET',
                        headers: { 'Accept': 'application/json;odata=verbose' }
                    });
    
                    if (!response.ok) throw new Error(`Failed to fetch term: ${termId}`);
    
                    const data = await response.json();
                    const translatedLabel = data.labels.find((label: { languageTag: string; }) => label.languageTag === lang)?.name || null;
                    
                    return translatedLabel;
                } catch (error) {
                    console.error("Error fetching term:", error);
                    return null;
                }
            });
    
            const translatedTerms = await Promise.all(fetchPromises);
            const finalTranslation = translatedTerms.filter(label => label !== null).join(", ");
    
            const template = document.getElementById(elementId);
            if (template) {
                template.innerText = finalTranslation;
            } else {
                console.error(`Couldn't find element with ID ${elementId}`);
            }
        }
    }

    translateTerm(props.jobType, props.jobTypeTermSetGuid, `jobType-${props.path}`);

    
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

    const termLabel = (value: string) => {
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
                    <div>{strings.opportunityType}: <span id={`jobType-${props.path}`}>{termLabel(props.jobType)}</span></div>
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