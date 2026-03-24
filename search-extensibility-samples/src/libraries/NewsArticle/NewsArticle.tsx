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
    viewCount?: number;                 // View Count
    viewCountLifetime?: number;         // View Count Lifetime
    lastModifiedTime?: Date;            // Last Modified Time
    pictureURL?: string;                // User profile picture
    author?: string;                    // News article author
    authorOwsuser?: string              // New article author account information
    friendlyLastModifiedTime?: string;
}

const NewsArticleComponent: React.FC<INewsArticleProps> = (props) => {

    //const theme = useTheme();
    //const strings = SelectLanguage(Globals.getLanguage());
    //const lang = Globals.getLanguage();

    const date = new Date(props.created);
    const createdDate = props.created ? (`${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`) : 'N/A';
    const email = props.authorOwsuser.substring(0, props.authorOwsuser.indexOf('|')).trim();

    console.log("props.pictureURL: ", props.pictureURL);
    console.log("props.author: ", props.author);
    console.log("props.createdBy: ", props.createdBy);
    console.log("props.AuthorOWSUSER: ", props.authorOwsuser);
    console.log("email: ", email);
    console.log("props.lastModifiedTime", props.lastModifiedTime);
    console.log("props.friendlyLastModifiedTime", props.friendlyLastModifiedTime);

    // const getContactNameInitials = () => {
    //     if (props.author) {
    //         let nameSplit = props.author.toUpperCase().split(' ');
    //         return nameSplit[0] ? nameSplit[0][0] + (nameSplit[1] ? nameSplit[1][0] : '') : 'NA';
    //     }
    //     return 'NA';
    // };

    return (
        <div className='gcx-news-card'>
            <div className='oliver-test'>
                <div className='picture-test'>
                    <img className='thumbnail' src={props.pictureThumbnailUrl} />
                </div>
                <div className='details-test'>
                    <Link href={props.siteUrl}>{props.siteTitle}</Link><br />
                    <h3><Link href={props.path}>{props.title}</Link></h3>
                    
                    <div className='ellipsis-text'>
                        {props.hitHighlightedSummary}
                    </div>

                    <div className='news-card-author'>
                        <div className='user-icon'>
                            <img className='profile' src={String.prototype.concat("https://devgcx.sharepoint.com/_layouts/15/userphoto.aspx?size=S&accountname=", email)} />
                        </div>

                        <div className='news-meta-info'>
                            {props.author}&nbsp;{props.friendlyLastModifiedTime}<br />
                            {props.viewCount ? props.viewCount : "0"} Views
                        </div>
                    </div>
                </div>
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