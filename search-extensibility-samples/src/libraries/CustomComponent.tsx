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
    approvedStaffing?: boolean;
    assetSkills?: string;
    city?: string;
    classificationCode?: string;
    classificationLevel?: string;
    contactEmail?: string;
    contactName?: string;
    contactObjectId?: string;
    duration?: string;
    department?: string;
    essentialSkills?: string;
    jobDescriptionEn?: string;
    jobDescriptionFr?: string;
    jobTitleEn?: string;
    jobTitleFr?: string;
    jobType?: string;
    languageRequirement?: string;
    location?: string;
    numberOfOpportunities?: string;
    programArea?: string;
    securityClearance?: string;
    workArrangement?: string;
    workSchedule?: string;
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
            window.open(`https://devgcx.sharepoint.com/sites/CM-test/SitePages/Job-Opportunity.aspx?JobOpportunityId=${props.path.split('ID=')[1]}`, '_blank');
    };

    const handleApplyClick = () => {
        window.location.href = `mailto:${props.contactEmail}?subject=${lang === Language.French ? props.jobTitleFr : props.jobTitleEn}`;
    };

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
                    <div>{strings.duration}: {props.duration}</div>
                </div>
                <div className="description">
                    {lang === Language.French ? props.jobDescriptionFr : props.jobDescriptionEn}
                </div>
                <div className="sub bold">
                    <div>{strings.location}: {props.location}</div>
                    <div>{strings.deadline}: {props.applicationDeadlineDate ? props.applicationDeadlineDate : 'None'}</div>
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