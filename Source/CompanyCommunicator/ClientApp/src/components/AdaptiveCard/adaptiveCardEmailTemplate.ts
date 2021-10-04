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


export const getInitAdaptiveCardEmailTemplate = (t: TFunction) => {
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
                        "type": "TextBlock",
                        "wrap": true,
                        "text": ""
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "text": ""
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
                "version": "1.2"
            }
        );

}

export const getCardTitleEmailTemplate = (card: any) => {
    return card.body[0].text;
}

export const setCardTitleEmailTemplate = (card: any, title: string) => {
    card.body[0].text = title;
}

export const getCardImageLinkEmailTemplate = (card: any) => {
    return card.body[1].url;
}

export const getCardImageselectActionLinkEmailTemplate = (card: any) => {
    return card.body[1].selectAction.url;
}

export const setCardImageLinkEmailTemplate = (card: any, imageLink?: string) => {
    card.body[1].url = imageLink;
    card.body[1].selectAction.url = imageLink;
}

export const getCardFileNameEmailTemplate = (card: any) => {
    return card.body[2].text;
}

export const setCardFileNameEmailTemplate = (card: any, link?: string) => {
    card.body[2].text = link;
}


export const getCardSummaryEmailTemplate = (card: any) => {
    return card.body[3].text;
}

export const setCardSummaryEmailTemplate = (card: any, summary?: string) => {
    card.body[3].text = summary;
}

export const getCardAuthorEmailTemplate = (card: any) => {
    return card.body[4].text;
}

export const setCardAuthorEmailTemplate = (card: any, author?: string) => {
    card.body[4].text = author;
}




