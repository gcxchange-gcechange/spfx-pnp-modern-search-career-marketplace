/* eslint-disable no-constant-condition */
import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
//import { useTheme } from '@fluentui/react';
//import { SelectLanguage } from '../SelectLanguage';
import './NewsArticle.css';
import { Globals, Language } from '../Globals';

export interface INewsArticleProps {
    path?: string;
}

const NewsArticleComponent: React.FC<INewsArticleProps> = (props) => {

    //const theme = useTheme();
    //const strings = SelectLanguage(Globals.getLanguage());
    const lang = Globals.getLanguage();

    return (
        <div>
            {lang === Language.French ? '' : 'TODO: Add the HTML for the job card here.'}
        </div>
    );
};

export class NewsArticleWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        const NewsArticleCard = <NewsArticleComponent {...props} />;
        ReactDOM.render(NewsArticleCard, this);
    }    

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}