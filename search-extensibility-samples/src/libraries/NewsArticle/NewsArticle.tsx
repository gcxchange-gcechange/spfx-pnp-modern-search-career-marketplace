/* eslint-disable no-constant-condition */
import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import './NewsArticle.css';
import { Link } from '@fluentui/react';
import { Globals, Language } from "../Globals";

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
    description?: string;
}

const NewsArticleComponent: React.FC<INewsArticleProps> = (props) => {
    const split = props.authorOwsuser.split(' | ');
    const email = split[0];
    const author = split[1];

    // console.log("props.pictureURL: ", props.pictureURL);
    // console.log("props.author: ", props.author);
    // console.log("props.createdBy: ", props.createdBy);
    // console.log("props.AuthorOWSUSER: ", props.authorOwsuser);
    // console.log("email: ", email);
    // console.log("props.lastModifiedTime", props.lastModifiedTime);
    // console.log("props.friendlyLastModifiedTime", props.friendlyLastModifiedTime);

    // Unable to get the elipsis using CSS was giving <ddd/> instead of ...
    const stripHtml = (html: string) => {
        const temp = document.createElement('div');
        temp.innerHTML = html;
        return temp.textContent || temp.innerText || '';
    };

    const truncateText = (text: string, maxLength: number) => {
        const cleanText = stripHtml(text);
        if (cleanText.length <= maxLength) 
            return cleanText;
        const trimmed = cleanText.substring(0, maxLength);
        return trimmed.substring(0, trimmed.lastIndexOf(' '));
    };

    const formatCreatedDate= (): string => {
        const date = new Date(props.created);
        const seconds = Math.floor((Date.now() - date.getTime()) / 1000);

        const intervals = [
            { labelEn: 'year', labelFr: 'année', seconds: 31536000 },
            { labelEn: 'month', labelFr: 'mois', seconds: 2592000 },
            { labelEn: 'week', labelFr: 'semaine', seconds: 604800 },
            { labelEn: 'day', labelFr: 'jour', seconds: 86400 },
            { labelEn: 'hour', labelFr: 'heure', seconds: 3600 },
            { labelEn: 'minute', labelFr: 'minute', seconds: 60 }
        ];

        for (const interval of intervals) {
            const count = Math.floor(seconds / interval.seconds);

            if (count >= 1) {

                if (Globals.getLanguage() === Language.French)
                    return `a publié il y a ${count} ${interval.labelFr}${count > 1 ? (interval.labelEn !== 'month' ? 's' : '') : ''}`;

                return `posted ${count} ${interval.labelEn}${count > 1 ? 's' : ''} ago`;
            }
        }

        if (Globals.getLanguage() === Language.French)
            return `Publié à l'instant`;

        return `posted just now`;
    }

    return ( 
        <div className='gcx-news-card'>
            <div className='newsArticle-cardImage'>
                {props.pictureThumbnailUrl ? (
                    <Link href={props.path}>
                        <img src={props.pictureThumbnailUrl} alt="thumbnail" />
                    </Link>
                    ) : (
                <Link href={props.path}>
                    <div className="newsArticle-cardImage-Default" />
                </Link>
                )}
            </div>
            <div className='newsArticle-cardContent'>
                <div className='newsArticle-cardTitle'>
                    <h3>
                        <Link style={{color: 'black'}} href={props.path}>
                            {truncateText(props.title, 100)}
                        </Link>
                    </h3>
                </div>
                <div className='newsArticle-cardDescription'>
                    <div>{props.description || truncateText(props.hitHighlightedSummary, 8675309)}</div>
                </div>

                <div className='newsArticle-cardAuthor'>
                    <img className='news-article-profile' src={`${Globals.tenant}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(email)}`} />
                    <p>
                        {author}&nbsp;{formatCreatedDate()}
                    </p>
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