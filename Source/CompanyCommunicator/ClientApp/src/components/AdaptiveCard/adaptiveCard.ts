// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TFunction } from "i18next";
import * as AdaptiveCards from "adaptivecards";
import MarkdownIt from "markdown-it";

// Static method to render markdown on the adaptive card
AdaptiveCards.AdaptiveCard.onProcessMarkdown = function (text, result) {
    var md = new MarkdownIt();
    // Teams only supports a subset of markdown as per https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-format?tabs=adaptive-md%2Cconnector-html#formatting-cards-with-markdown
    md.disable(['image', 'table', 'heading',
        'hr', 'code', 'reference',
        'lheading', 'html_block', 'fence',
        'blockquote', 'strikethrough']);
    // renders the text
    result.outputHtml = md.render(text);
    result.didProcess = true;
}

export const getInitAdaptiveCard = (t: TFunction) => {
    const titleTextAsString = t("TitleText");

    return (
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "weight": "Bolder",
                    "text": titleTextAsString,
                    "size": "ExtraLarge",
                    "wrap": true
                },
                {
                    "type": "Image",
                    "spacing": "Default",
                    "url": "",
                    "msTeams": {
                        "allowExpand": true
                    },
                    "selectAction": {
                        "type": "Action.OpenUrl",
                        "title": "Image",
                        "url": ""
                      },
                    "size": "Stretch",
                    "width": "300px",
                    "altText": ""
                },
                {
                    "type": "Image",
                    "spacing": "Default",
                    "url": "",
                    "width": "80px",
                    "height": "80px",
                    "altText": "",
                    "horizontalAlignment": "left"
                },
                {
                    "type": "TextBlock",
                    "text": "",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "text": "",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": ""
                }
            ],
            "msteams": {
                "width": "Full"
            },
            "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.4"
        }
    );
}

export const getCardTitle = (card: any) => {
    return card.body[0].text;
}

export const setCardTitle = (card: any, title: string) => {
    card.body[0].text = title;
}

export const getCardImageLink = (card: any) => {
    return card.body[1].url;
}

export const getCardImageselectActionLink = (card: any) => {
    return card.body[1].selectAction.url;
}

export const setCardImageLink = (card: any, imageLink?: string) => {
    card.body[1].url = imageLink;
    card.body[1].selectAction.url = imageLink;
}

export const getCardPDFImage = (card: any) => {
    return card.body[2].url;
}

export const setCardPDFImage = (card: any, imageLink?: string) => {
    card.body[2].url = imageLink;
}

export const getCardPdfName = (card: any) => {
    return card.body[3].text;
}

export const setCardPdfName = (card: any, link?: string) => {
    card.body[3].text = link;
}


export const getCardSummary = (card: any) => {
    return card.body[4].text;
}

export const setCardSummary = (card: any, summary?: string) => {
    card.body[4].text = summary;
}

export const getCardAuthor = (card: any) => {
    return card.body[5].text;
}

export const setCardAuthor = (card: any, author?: string) => {
    card.body[5].text = author;
}

export const getCardBtnTitle = (card: any) => {
    return card.actions[0].title;
}

export const getCardBtnLink = (card: any) => {
    return card.actions[0].url;
}

// set the values collection with buttons to the card actions
export const setCardBtns = (card: any, values: any[]) => {
    if (values !== null) {
        card.actions = values;
    } else {
        delete card.actions;
    }
}


// https://teams.microsoft.com/l/stage/5ca3098c-6c7d-4080-9bc0-eefc4fc26f00/0?context={"contentUrl":"https://5dt3sqbkwreas.blob.core.windows.net:443/pdffiles/011ce496-ced4-44a1-b76c-dfe7f1baf4c1_800x600.jpg","websiteUrl":"https://5dt3sqbkwreas.blob.core.windows.net:443/pdffiles/011ce496-ced4-44a1-b76c-dfe7f1baf4c1_800x600.jpg","name":"Image"}