/* eslint-disable no-constant-condition */
import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { useTheme } from '@fluentui/react';
import { SelectLanguage } from './SelectLanguage';
import './JobCardSortComponent.css';
import { Globals } from './Globals';

export interface IJobCardSortComponentProps {
    
}

export enum JobSortSessionKeys {
  Initialized = 'gcx-cm-sortmod-init',
  Value = 'gcx-cm-sortmod-value'
}

const SortComponent: React.FC<IJobCardSortComponentProps> = (props) => {

    const theme = useTheme();
    const strings = SelectLanguage(Globals.getLanguage());
    const lang = Globals.getLanguage();

    React.useEffect(() => {
        sessionStorage.setItem(JobSortSessionKeys.Initialized, 'true');
        return () => {
            sessionStorage.removeItem(JobSortSessionKeys.Initialized);
        };
    }, []);

    function sort(value: string) {
        sessionStorage.setItem(JobSortSessionKeys.Value, value);
    }

    return (
        <select id="cm-sortDropdown" onChange={e => sort(e.target.value)}>
            <option value="sortlist:'Created:descending'">
                { strings.sortbyDateDescending }
            </option>
            <option value="sortlist:'Created:ascending'">
                { strings.sortByDateAscending }
            </option>
            <option value="sortlist:'CM-ApplicationDeadlineDate:descending'">
                { strings.sortByDeadlineDescending }
            </option>
            <option value="sortlist:'CM-ApplicationDeadlineDate:ascending'">
                { strings.sortbyDeadlineAscending }
            </option>
        </select>
    );
};

export class JobCardSortComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        const sortComp = <SortComponent {...props} />;
        ReactDOM.render(sortComp, this);
    }    

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}