/* eslint-disable no-constant-condition */
import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
//import { useTheme } from '@fluentui/react';
//import { SelectLanguage } from '../SelectLanguage';
import './NewsArticle.css';
import { Link } from '@fluentui/react';
//import { Globals, Language } from '../Globals';

export interface INewsArticleProps {
    path?: string;                      // Link to the news post
    title?: string;                     // Title
    hitHighlightedSummary?: string;     // Summarry
    pictureThumbnailUrl?: string;       // Article picture (thumbnail)
    siteTitle?: string;                 // Title of the site 
    siteLogo?: string;                  // Logo of the site
    siteUrl?: string;                   // Url of the site
    createdBy?: string;                 // Author
    created?: string;                   // Creation date string (UTC)
}

const NewsArticleComponent: React.FC<INewsArticleProps> = (props) => {

    //const theme = useTheme();
    //const strings = SelectLanguage(Globals.getLanguage());
    //const lang = Globals.getLanguage();

    const date = new Date(props.created);
    const createdDate = props.created ? (`${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`) : 'N/A';

    return (
        <div className='gcx-news-card'>
            <div className='news-card-header'>
                <img className='thumbnail' src={props.pictureThumbnailUrl} />
                <div className='header-text'>
                    <div>
                        <Link href={props.siteUrl}>
                            {props.siteTitle}
                        </Link>
                    </div>
                    <div>
                        <Link href={props.path}>
                            {props.title}
                        </Link>
                    </div>
                </div>
            </div>
            <div className='news-card-content'>
                <span>{props.hitHighlightedSummary}</span>
            </div>
            <div className='news-card-author'>
                <span>{props.createdBy}</span>
                <span>{createdDate}</span>
            </div>
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