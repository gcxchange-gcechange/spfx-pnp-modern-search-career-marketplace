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
}

const JobCardComponent: React.FC<ICustomComponentProps> = (props) => {

    const theme = useTheme();
    const strings = SelectLanguage(Globals.getLanguage());
    const lang = Globals.getLanguage();
    
    const getContactNameInitials = () => {
        if (props.contactName) {
            let nameSplit = props.contactName.toUpperCase().split(' ');
            return nameSplit[0] ? nameSplit[0][0] + (nameSplit[1] ? nameSplit[1][0] : '') : 'NA';
        }
        return 'NA';
    };

    const getJobType = () => {
        return props.jobType ? props.jobType.replace(';', ', ') : 'N/A';
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
                    <div>{strings.opportunityType}: {getJobType()}</div>
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