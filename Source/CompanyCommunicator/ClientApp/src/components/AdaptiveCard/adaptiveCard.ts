// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TFunction } from "i18next";

export const getInitAdaptiveCard = (t: TFunction) => {
    const titleTextAsString = t("TitleText");
    return (
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "Container",
                    "bleed": true,
                    "backgroundImage": {
                        "url": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAkAAAAJCAIAAABv85FHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAEnQAABJ0Ad5mH3gAAAASSURBVBhXY7gs6IoLDW85QVcAcSpd+tJRPKIAAAAASUVORK5CYII=",
                        "fillMode": "repeat"
                    },
                    "id": "headercontainer",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "HEADER",
                            "weight": "bolder",
                            "horizontalAlignment": "center",
                            "size": "default",
                            "color": "light",
                            "wrap": true
                        }
                    ]
                },
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
                    "size": "Stretch",
                    "width": "400px",
                    "altText": ""
                },
                {
                    "type": "TextBlock",
                    "text": "",
                    "weight": "",
                    "fontType": "",
                    "size": "Medium",
                    "horizontalAlignment": "left",
                    "color":"default",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": ""
                },
                {
                    "type": "ActionSet",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Watch Video",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "ActionSet",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 1",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "ActionSet",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 2",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "ActionSet",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 3",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "ActionSet",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 4",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "ActionSet",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 5",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "",
                    "text": ""
                },
                {
                    "type": "",
                    "id": "Reaction",
                    "style": "expanded",
                    "isMultiSelect": false,
                    "choices": [
                        {
                            "title": "Extremely satisfied",
                            "value": "1"
                        },
                        {
                            "title": "Somewhat satisfied",
                            "value": "2"
                        },
                        {
                            "title": "Neither satisfied nor dissatisfied",
                            "value": "3"
                        },
                        {
                            "title": "Somewhat dissatisfied",
                            "value": "4"
                        },
                        {
                            "title": "Extremely dissatisfied",
                            "value": "5"
                        }
                    ]
                },
                {
                    "type": "",
                    "text": ""
                },
                {
                    "type": "",
                    "id": "FreeTextSurvey",
                    "placeholder": "Enter Text Here",
                    "maxLength": 500,
                    "isMultiline": true
                },
                {
                    "type": "",
                    "text": ""
                },
                {
                    "type": "",
                    "id": "YesNo",
                    "style": "expanded",
                    "isMultiSelect": false,
                    "value": "1",
                    "choices": [
                        {
                            "title": "Yes",
                            "value": "Yes"
                        },
                        {
                            "title": "No",
                            "value": "No"
                        }
                    ]
                },
                {
                    "type": "",
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Submit",
                        }
                    ]
                },
                {
                    "type": "",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Open Survey",
                            "url": ""
                        }
                    ]
                },
            ],
            "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.2"
        }
    );
}

export const getCardTitle = (card: any) => {
    return card.body[1].text;
}

export const setCardTitle = (card: any, title: string) => {
    card.body[1].text = title;
}

export const getCardImageLink = (card: any) => {
    return card.body[2].url;
}

export const setCardImageLink = (card: any, imageLink?: string) => {
    card.body[2].url = imageLink;
}

export const getCardSummary = (card: any) => {
    return card.body[3].text;
}

export const setCardSummary = (card: any, summary?: string) => {
    card.body[3].text = summary;
}

export const getCardAuthor = (card: any) => {
    return card.body[4].text;
}

export const setCardAuthor = (card: any, author?: string) => {
    card.body[4].text = author;
}

export const getCardBtnTitle = (card: any) => {
    return card.actions[1].title;
}

export const getCardBtnLink = (card: any) => {
    return card.actions[1].url;
}

export const setCardBtn = (card: any, buttonTitle?: string, buttonLink?: string) => {
    if (buttonTitle && buttonLink) {
        card.body[6].type = "ActionSet";
        card.body[6].actions[0].title = buttonTitle;
        card.body[6].actions[0].url = buttonLink;
    } else {
        card.body[6].type = "";
    }
}

export const setCardBtn2 = (card: any, buttonTitle2?: string, buttonLink2?: string) => {
    if (buttonTitle2 && buttonLink2) {
        card.body[7].type = "ActionSet";
        card.body[7].actions[0].title = buttonTitle2;
        card.body[7].actions[0].url = buttonLink2;
    } else {
        card.body[7].type = "";
    }
}

export const setCardBtn3 = (card: any, buttonTitle3?: string, buttonLink3?: string) => {
    if (buttonTitle3 && buttonLink3) {
        card.body[8].type = "ActionSet";
        card.body[8].actions[0].title = buttonTitle3;
        card.body[8].actions[0].url = buttonLink3;
    } else {
        card.body[8].type = "";
    }
}

export const setCardBtn4 = (card: any, buttonTitle4?: string, buttonLink4?: string) => {
    if (buttonTitle4 && buttonLink4) {
        card.body[9].type = "ActionSet";
        card.body[9].actions[0].title = buttonTitle4;
        card.body[9].actions[0].url = buttonLink4;
    } else {
        card.body[9].type = "";
    }
}

export const setCardBtn5 = (card: any, buttonTitle5?: string, buttonLink5?: string) => {
    if (buttonTitle5 && buttonLink5) {
        card.body[10].type = "ActionSet";
        card.body[10].actions[0].title = buttonTitle5;
        card.body[10].actions[0].url = buttonLink5;
    } else {
        card.body[10].type = "";
    }
}

export const setVideoBtn = (card: any, videoLink?: string) => {
    if (videoLink) {
        card.body[5].type = "ActionSet";
        card.body[5].actions[0].url = videoLink;
    } else {
        card.body[5].type = "";
    }
}
export const setSenderTemplate = (card: any, senderTemplate?: string) => {
    if (senderTemplate) {
        card.body[0].type = "Container";
        card.body[0].items[0].text = senderTemplate;
    } else {
        card.body[0].type = "";
    }
}
export const setBoldSummary = (card: any, boldSummary?: string) => {
    if (boldSummary === "Bold") {
        card.body[3].weight = "Bolder";
    } else {
        card.body[3].weight = "Default";
    }
}
export const setFontSummary = (card: any, fontSummary?: string) => {
    if (fontSummary === "Monospace") {
        card.body[3].fontType = "Monospace";
    } else {
        card.body[3].fontType = "Default";
    }
}
export const setAlignmentSummary = (card: any, alignment?: string) => {
    if (alignment === "left") {
        card.body[3].horizontalAlignment = "left";
    } else if (alignment === "center") {
        card.body[3].horizontalAlignment = "center";
    } else if (alignment === "right") {
        card.body[3].horizontalAlignment = "right";
    }
}
export const setFontSizeSummary = (card: any, fontSizeSummary?: string) => {
    if (fontSizeSummary === "default") {
        card.body[3].size = "default";
    } else if (fontSizeSummary === "small") {
        card.body[3].size = "small";
    } else if (fontSizeSummary === "medium") {
        card.body[3].size = "medium";
    } else if (fontSizeSummary === "large") {
        card.body[3].size = "large";
    } else if (fontSizeSummary === "extraLarge") {
        card.body[3].size = "extraLarge";
    }
}
export const setFontColorSummary = (card: any, fontColorSummary?: string) => {
    if (fontColorSummary === "default") {
        card.body[3].color = "default";
    } else if (fontColorSummary === "accent") {
        card.body[3].color = "accent";
    } else if (fontColorSummary === "good") {
        card.body[3].color = "good";
    } else if (fontColorSummary === "warning") {
        card.body[3].color = "warning";
    } else if (fontColorSummary === "attention") {
        card.body[3].color = "attention";
    } else if (fontColorSummary === "light") {
        card.body[3].color = "light";
    } else if (fontColorSummary === "dark") {
        card.body[3].color = "dark";
    }
}
export const setReaction = (card: any, isReaction?: boolean) => {
    if (isReaction === true) {
        card.body[11].type = "TextBlock";
        card.body[12].type = "Input.ChoiceSet";
    } else {
        card.body[11].type = "";
        card.body[12].type = "";
    }
}
export const SetReactionQuestion = (card: any, reactionQuestion?: string) => {
    card.body[11].text = reactionQuestion;
}
export const SetFreeText = (card: any, isFreeText?: boolean) => {
    if (isFreeText === true) {
        card.body[13].type = "TextBlock";
        card.body[14].type = "Input.Text";
    } else {
        card.body[13].type = "";
        card.body[14].type = "";
    }
}
export const SetFreeTextQuestion = (card: any, freeTextQuestion?: string) => {
    card.body[13].text = freeTextQuestion;
}
export const setYesNo = (card: any, isYesNo?: boolean) => {
    if (isYesNo === true) {
        card.body[15].type = "TextBlock";
        card.body[16].type = "Input.ChoiceSet";
    } else {
        card.body[15].type = "";
        card.body[16].type = "";
    }
}
export const setYesNoQuestion = (card: any, yesNoQuestion?: string) => {
    card.body[15].text = yesNoQuestion;
}
export const setSubmitBtn = (card: any, isSubmit?: boolean) => {
    if (isSubmit === true) {
        card.body[17].type = "ActionSet";
    } else {
        card.body[17].type = "";
    }
}
export const setToggleLinktoSurvey = (card: any, toggleLinktoSurvey?: boolean) => {
    if (toggleLinktoSurvey === true) {
        card.body[18].type = "ActionSet";
    } else {
        card.body[18].type = "";
    }
}
export const setLinktoSurvey = (card: any, linkToSurvey?: string) => {
        card.body[18].actions[0].url = linkToSurvey;
}