/* eslint-disable no-constant-condition */
import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import './NewsArticle.css';
import { Link } from '@fluentui/react';
import { Globals } from "../Globals";
import { SelectLanguage } from "./../SelectLanguage";

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
    const strings = SelectLanguage(Globals.getLanguage());
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

    // Unable to get the elipsis using CSS was giving <ddd/> insted of ...
    const stripHtml = (html: string) => {
        const temp = document.createElement('div');
        temp.innerHTML = html;
        return temp.textContent || temp.innerText || '';
    };

    const truncateText = (text: string, maxLength: number) => {
        const cleanText = stripHtml(text);
        if (cleanText.length <= maxLength) return cleanText;
        const trimmed = cleanText.substring(0, maxLength);
        return trimmed.substring(0, trimmed.lastIndexOf(' '));
    };

    return ( 
        <div className='gcx-news-card'>
            <div className='newsArticle-cardImage'>
                {props.pictureThumbnailUrl ? (
                    <img src={props.pictureThumbnailUrl} alt="thumbnail" />
                    ) : (
                <div className="newsArticle-cardImage-Default" />
                )}
            </div>
            <div className='newsArticle-cardContent'>
                <div className='newsArticle-cardTitle'>
                    <Link style={{fontSize: 'smaller', fontWeight: '500' }} href={props.siteUrl}>{props.siteTitle}</Link>
                    <h3><Link style={{color: 'black'}}  href={props.path}>{props.title}</Link></h3>
                </div>
                <p >
                    {truncateText(props.hitHighlightedSummary,266)} ...
                </p>

                <div className='newsArticle-cardAuthor'>
                    <img className='news-article-profile' src={String.prototype.concat("https://devgcx.sharepoint.com/_layouts/15/userphoto.aspx?size=S&accountname=", email)} />
                    <p>{props.author}&nbsp;{props.friendlyLastModifiedTime} <br/>
                    {props.viewCount ? props.viewCount : "0"} {strings.views}</p>
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