import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PrimaryButton, DefaultButton, useTheme } from '@fluentui/react';
import { SelectLanguage } from './SelectLanguage';
import './CustomComponent.css';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {
    path?: string;
    applicationdeadlinedate?: Date;
    approvedstaffing?: boolean;
    assetskills?: string;
    city?: string;
    classificationcode?: string;
    classificationlevel?: string;
    contactemail?: string;
    contactname?: string;
    contactobjectid?: string;
    duration?: string;
    essentialskills?: string;
    jobdescriptionen?: string;
    jobdescriptionfr?: string;
    jobtitleen?: string;
    jobtitlefr?: string;
    jobtype?: string;
    languagerequirement?: string;
    location?: string;
    numberofopportunities?: string;
    programarea?: string;
    securityclearance?: string;
    workarrangement?: string;
    workschedule?: string;
    selectedlanguage?: string;
}

const JobCardComponent: React.FC<ICustomComponentProps> = (props) => {

    if (props.selectedlanguage)
        props.selectedlanguage = props.selectedlanguage.toLowerCase();

    const theme = useTheme();
    const strings = SelectLanguage(props.selectedlanguage);
    
    const getContactNameInitials = () => {
        if (props.contactname) {
            let nameSplit = props.contactname.toUpperCase().split(' ');
            return nameSplit[0] ? nameSplit[0][0] + (nameSplit[1] ? nameSplit[1][0] : '') : 'NA';
        }
        return 'NA';
    };

    const getJobType = () => {
        return props.jobtype ? props.jobtype.replace(';', ', ') : 'N/A';
    };

    const handleViewClick = () => {
        if (props.path && props.path.split('ID=').length == 2) 
            window.open(`https://devgcx.sharepoint.com/sites/CM-test/SitePages/Job-Opportunity.aspx?JobOpportunityId=${props.path.split('ID=')[1]}`, '_blank');
    };

    const handleApplyClick = () => {
        window.location.href = `mailto:${props.contactemail}?subject=${props.selectedlanguage == 'fr' ? props.jobtitlefr : props.jobtitleen}`;
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
                    {props.selectedlanguage == 'fr' ? props.jobtitlefr : props.jobtitleen}
                </h3>
                <div className="sub">
                    <div>{strings.classificationLevel}: {props.classificationlevel}</div>
                    <div>{strings.opportunityType}: {getJobType()}</div>
                    <div>{strings.duration}: {props.duration}</div>
                </div>
                <div className="description">
                    {props.selectedlanguage == 'fr' ? props.jobdescriptionfr : props.jobdescriptionen}
                </div>
                <div className="sub bold">
                    <div>{strings.location}: {props.location}</div>
                    <div>{strings.deadline}: {props.applicationdeadlinedate ? props.applicationdeadlinedate : 'None'}</div>
                </div>
                <div className="contact">
                    <div className="profile">
                        <div>
                            {getContactNameInitials()}
                        </div>
                    </div>
                    <div className="info">
                        <div>{props.contactname}</div>
                        <div>{props.contactemail}</div>
                    </div>
                </div>
                <div className="actions">
                    <DefaultButton 
                        text={strings.view} 
                        onClick={handleViewClick}
                    />
                    <PrimaryButton 
                        text={strings.apply}
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
        const customComponent = <JobCardComponent {...props} />;
        ReactDOM.render(customComponent, this);
    }    

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}