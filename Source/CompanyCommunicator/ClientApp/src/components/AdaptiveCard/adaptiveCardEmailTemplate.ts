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




export const setCardEmailTemplate = (card: any, imageLink?: string, title?: any, author?: any, emailLink?:any, summary?:any) => {
    var jsonArray = new Array();
    jsonArray.push(`{
            "type": "TextBlock",
            "weight": "Bolder",
            "text": "${title}",
            "size": "ExtraLarge",
            "wrap": true
        }`)
        if(imageLink !== ""){
            jsonArray.push(`{
                "type": "Image",
                "spacing": "Default",
                "url": "${imageLink}",
                "msTeams": {
                    "allowExpand": true
                },
                "selectAction": {
                    "type": "Action.OpenUrl",
                    "title": "Image",
                    "url": "${imageLink}"
                  },
                "size": "Stretch",
                "width": "300px",
                "altText": ""
            }`)
        }
        
        jsonArray.push(`{
            "type": "TextBlock",
            "wrap": true,
            "size": "Small",
            "weight": "Lighter",
            "text": "${emailLink}"
        }`)
        jsonArray.push(`{
            "type": "TextBlock",
            "wrap": true,
            "size": "Small",
            "weight": "Lighter",
            "text": "${summary}"
        }`)
        jsonArray.push(`{
            "type": "TextBlock",
            "wrap": true,
            "size": "Small",
            "weight": "Lighter",
            "text": "${author}"
        }`)

        var jsonQuestionCard = "[" + jsonArray.join() + "]";
        card.body = JSON.parse(jsonQuestionCard);
        return card.body[0];

}






