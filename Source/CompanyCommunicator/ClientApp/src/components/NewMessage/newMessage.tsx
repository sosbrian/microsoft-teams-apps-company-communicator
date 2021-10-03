// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as AdaptiveCards from "adaptivecards";
import { Button, Loader, Dropdown, Text, Flex, Input, TextArea, RadioGroup, FlexItem } from '@fluentui/react-northstar'
import * as microsoftTeams from "@microsoft/teams-js";

import './newMessage.scss';
import './teamTheme.scss';
import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, searchGroups, getGroups, verifyGroupAccess, getUsers } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn, setCardBtn2, setCardBtn3, setCardBtn4, setCardBtn5, setVideoBtn, setSenderTemplate, setBoldSummary, setFontSummary, setAlignmentSummary, setFontSizeSummary, setFontColorSummary, setReaction, SetReactionQuestion, SetFreeText, SetFreeTextQuestion, setYesNo, setYesNoQuestion, setSubmitBtn, setToggleLinktoSurvey, setLinktoSurvey
} from '../AdaptiveCard/adaptiveCard';
import { getBaseUrl } from '../../configVariables';
import { ImageUtil } from '../../utility/imageutility';
import { TFunction } from "i18next";
//Using material-ui PublishIcon
import IconButton from '@material-ui/core/IconButton';
import PublishIcon from '@material-ui/icons/Publish';
import { faAlignLeft, faAlignCenter, faAlignRight, faBold, faFont, faUpload } from '@fortawesome/free-solid-svg-icons';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { Switch, FormLabel, FormControl, FormGroup, FormControlLabel, FormHelperText, Grid } from '@material-ui/core';
//import Button as Button1 } from '@mui/material/Button';
import { ToggleButton, ToggleButtonGroup } from '@material-ui/lab';
//import Paper from '@material-ui/lab';
import FormatBoldIcon from '@material-ui/icons/FormatBold';
import FontDownloadIcon from '@material-ui/icons/FontDownload';
import FormatAlignLeftIcon from '@material-ui/icons/FormatAlignLeft';
import FormatAlignCenterIcon from '@material-ui/icons/FormatAlignCenter';
import FormatAlignRightIcon from '@material-ui/icons/FormatAlignRight';
import { CSVReader } from 'react-papaparse';
import UploadIcon from '@mui/icons-material/Upload';
import { styled } from '@material-ui/core/styles';
import { forEachChild } from 'typescript';


type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: {
        id: string,
    },
}

export interface IDraftMessage {
    template?: string,
    senderTemplate: string,
    id?: string,
    title: string,
    imageLink?: string,
    videoLink?: string,
    alignment?: string,
    summary?: string,
    boldSummary?: string,
    fontSummary?: string,
    fontSizeSummary?: string,
    fontColorSummary?: string,
    author: string,
    buttonTitle?: string,
    buttonLink?: string,
    buttonTitle2?: string,
    buttonLink2?: string,
    buttonTitle3?: string,
    buttonLink3?: string,
    buttonTitle4?: string,
    buttonLink4?: string,
    buttonTitle5?: string,
    buttonLink5?: string,
    surReaction?: boolean,
    reactionQuestion?: string,
    surFreeText?: boolean,
    freeTextQuestion?: string,
    surYesNo?: boolean,
    yesNoQuestion?: string,
    surLinkToSurvey?: boolean,
    linkToSurvey?: string,
    teams: any[],
    rosters: any[],
    groups: any[],
    uploadedList: any[],
    uploadedListName: string,
    exclusionList: string,
    allUsers: boolean
}

export interface formState {
    template: string,
    senderTemplate: string,
    title: string,
    summary?: string,
    boldSummary?: string,
    fontSummary?: string,
    fontSizeSummary?: string,
    fontColorSummary?: string,
    alignment?: string,
    btnLink?: string,
    btnLink2?: string,
    btnLink3?: string,
    btnLink4?: string,
    btnLink5?: string,
    imageLink?: string,
    videoLink?: string,
    btnTitle?: string,
    btnTitle2?: string,
    btnTitle3?: string,
    btnTitle4?: string,
    btnTitle5?: string,
    surReaction?: boolean,
    reactionQuestion?: string,
    surFreeText?: boolean,
    freeTextQuestion?: string,
    surYesNo?: boolean,
    yesNoQuestion?: string,
    surLinkToSurvey?: boolean,
    linkToSurvey?: string,
    author: string,
    card?: any,
    page: string,
    teamsOptionSelected: boolean,
    rostersOptionSelected: boolean,
    allUsersOptionSelected: boolean,
    groupsOptionSelected: boolean,
    uploadOptionSelected: boolean,
    teams?: any[],
    groups?: any[],
    exists?: boolean,
    messageId: string,
    loader: boolean,
    groupAccess: boolean,
    loading: boolean,
    noResultMessage: string,
    unstablePinned?: boolean,
    selectedTeamsNum: number,
    selectedRostersNum: number,
    selectedGroupsNum: number,
    selectedRadioBtn: string,
    selectedTeams: dropdownItem[],
    selectedRosters: dropdownItem[],
    selectedGroups: dropdownItem[],
    uploadedList: any[],
    uploadedListName: string,
    errorImageUrlMessage: string,
    errorVideoUrlMessage: string,
    errorButtonUrlMessage: string,
    errorButtonUrlMessage2: string,
    errorButtonUrlMessage3: string,
    errorButtonUrlMessage4: string,
    errorButtonUrlMessage5: string,
    errorLinkToSurveyUrlMessage: string,
    selectedFile: string,
    selectedFileName: string,
    exclusionList: string,
    resetCSVReader: boolean,
}

export interface INewMessageProps extends RouteComponentProps, WithTranslation {
    getDraftMessagesList?: any;
}

class NewMessage extends React.Component<INewMessageProps, formState> {
    readonly localize: TFunction;
    private card: any;

    constructor(props: INewMessageProps) {
        super(props);
        initializeIcons();
        this.localize = this.props.t;
        //this.card = getInitAdaptiveCard(this.localize);
        this.card =
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
                            "size": "large",
                            "color": "light",
                            "wrap": true
                        },
                    ]
                },
                {
                    "type": "TextBlock",
                    "weight": "Bolder",
                    "text": "Title",
                    "size": "ExtraLarge",
                    "wrap": true
                },
                {
                    "type": "Image",
                    "spacing": "Default",
                    "url": getBaseUrl() + "/image/imagePlaceholder.png",
                    "size": "Stretch",
                    "width": "400px",
                    "altText": "Testing"
                },
                {
                    "type": "TextBlock",
                    "text": "Summary",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": "- Author"
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
                    "type": "",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 2",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 3",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Button title 4",
                            "url": "https://adaptivecards.io"
                        }
                    ]
                },
                {
                    "type": "",
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
        ;
        //this.setDefaultCard(this.card);
        //
        this.state = {
            template: "",
            senderTemplate: "",
            title: "",
            summary: "",
            alignment: "left",
            fontSummary: "",
            fontSizeSummary: "",
            fontColorSummary: "",
            boldSummary: "",
            author: "",
            btnLink: "",
            btnLink2: "",
            btnLink3: "",
            btnLink4: "",
            btnLink5: "",
            imageLink: "",
            videoLink: "",
            btnTitle: "",
            btnTitle2: "",
            btnTitle3: "",
            btnTitle4: "",
            btnTitle5: "",
            surReaction: false,
            reactionQuestion: "",
            surFreeText: false,
            freeTextQuestion: "",
            surYesNo: false,
            yesNoQuestion: "",
            surLinkToSurvey: false,
            linkToSurvey: "",
            card: this.card,
            page: "TemplateSelection",
            //page: "TemplateSelection",AudienceSelection
            teamsOptionSelected: true,
            rostersOptionSelected: false,
            allUsersOptionSelected: false,
            groupsOptionSelected: false,
            uploadOptionSelected: false,
            messageId: "",
            loader: true,
            groupAccess: false,
            loading: false,
            noResultMessage: "",
            unstablePinned: true,
            selectedTeamsNum: 0,
            selectedRostersNum: 0,
            selectedGroupsNum: 0,
            selectedRadioBtn: "teams",
            selectedTeams: [],
            selectedRosters: [],
            selectedGroups: [],
            uploadedList: [],
            uploadedListName: "",
            errorImageUrlMessage: "",
            errorVideoUrlMessage: "",
            errorButtonUrlMessage: "",
            errorButtonUrlMessage2: "",
            errorButtonUrlMessage3: "",
            errorButtonUrlMessage4: "",
            errorButtonUrlMessage5: "",
            errorLinkToSurveyUrlMessage: "",
            selectedFile: "",
            selectedFileName: "",
            exclusionList: "",
            resetCSVReader: false
        }
    }

    public async componentDidMount() {
        microsoftTeams.initialize();
        //- Handle the Esc key
        document.addEventListener("keydown", this.escFunction, false);
        let params = this.props.match.params;
        this.setGroupAccess();
        this.getTeamList().then(() => {
            if ('id' in params) {
                let id = params['id'];
                this.getItem(id).then(() => {
                    const selectedTeams = this.makeDropdownItemList(this.state.selectedTeams, this.state.teams);
                    const selectedRosters = this.makeDropdownItemList(this.state.selectedRosters, this.state.teams);
                    this.setState({
                        exists: true,
                        messageId: id,
                        selectedTeams: selectedTeams,
                        selectedRosters: selectedRosters,
                    })
                });
                this.getGroupData(id).then(() => {
                    const selectedGroups = this.makeDropdownItems(this.state.groups);
                    this.setState({
                        selectedGroups: selectedGroups
                    })
                });
            } else {
                this.setState({
                    exists: false,
                    loader: false
                }, () => {
                    let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                    adaptiveCard.parse(this.state.card);
                    let renderedCard = adaptiveCard.render();
                    document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                    if (this.state.btnLink) {
                        let link = this.state.btnLink;
                        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                    }
                })
            }
        });
    }

    private makeDropdownItems = (items: any[] | undefined) => {
        const resultedTeams: dropdownItem[] = [];
        if (items) {
            items.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id
                    },

                });
            });
        }
        return resultedTeams;
    }

    private makeDropdownItemList = (items: any[], fromItems: any[] | undefined) => {
        const dropdownItemList: dropdownItem[] = [];
        items.forEach(element =>
            dropdownItemList.push(
                typeof element !== "string" ? element : {
                    key: fromItems!.find(x => x.id === element).id,
                    header: fromItems!.find(x => x.id === element).name,
                    image: ImageUtil.makeInitialImage(fromItems!.find(x => x.id === element).name),
                    team: {
                        id: element
                    }
                })
        );
        return dropdownItemList;
    }

    public setDefaultCard = (card: any) => {
        if (this.state.template === "Announcements with image") {
            this.setState({
                card:
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
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 1",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle",
                                    "url": "https://adaptivecards.io/"
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
            }, () => {
                this.updateCard();
            });

        } else if (this.state.template === "Announcement with video") {
            this.setState({
                card:
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
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 2",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Watch Video",
                                    "url": "https://adaptivecards.io/"
                                }
                            ]
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle",
                                    "url": "https://adaptivecards.io/"
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
            }, () => {
                this.updateCard();
            });


        } else if (this.state.template === "Announcement in plain text") {
            this.setState({
                card:
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
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 3",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle",
                                    "url": "https://adaptivecards.io/"
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
            }, () => {
                this.updateCard();
            });


        } else if (this.state.template === "Event promotion with sessions to register/join") {
            this.setState({
                card:
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
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 4",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle1",
                                    "url": "https://adaptivecards.io/"
                                }
                            ]
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle2",
                                    "url": "https://adaptivecards.io/"
                                }
                            ]
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle3",
                                    "url": "https://adaptivecards.io/"
                                }
                            ]
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle4",
                                    "url": "https://adaptivecards.io/"
                                }
                            ]
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle5",
                                    "url": "https://adaptivecards.io/"
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
            }, () => {
                this.updateCard();
            });


        }
    }

    private getTeamList = async () => {
        try {
            const response = await getTeams();
            this.setState({
                teams: response.data
            });
        } catch (error) {
            return error;
        }
    }

    private getGroupItems() {
        if (this.state.groups) {
            return this.makeDropdownItems(this.state.groups);
        }
        const dropdownItems: dropdownItem[] = [];
        return dropdownItems;
    }

    private setGroupAccess = async () => {
        await verifyGroupAccess().then(() => {
            this.setState({
                groupAccess: true
            });
        }).catch((error) => {
            const errorStatus = error.response.status;
            if (errorStatus === 403) {
                this.setState({
                    groupAccess: false
                });
            }
            else {
                throw error;
            }
        });
    }

    private getGroupData = async (id: number) => {
        try {
            console.log("id" + id)
            const response = await getGroups(id);
            this.setState({
                groups: response.data
            });
            
        }
        catch (error) {
            console.log("error" + error)
            return error;
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getDraftNotification(id);
            const draftMessageDetail = response.data;
            console.log("response: " + response);
            let selectedRadioButton = "teams";
            if (draftMessageDetail.rosters.length > 0) {
                selectedRadioButton = "rosters";
            }
            else if (draftMessageDetail.groups.length > 0) {
                selectedRadioButton = "groups";
            }
            else if (draftMessageDetail.allUsers) {
                selectedRadioButton = "allUsers";
            }
            else if (draftMessageDetail.uploadedList) {
                selectedRadioButton = "uploadList"
            }
            this.setState({
                teamsOptionSelected: draftMessageDetail.teams.length > 0,
                selectedTeamsNum: draftMessageDetail.teams.length,
                rostersOptionSelected: draftMessageDetail.rosters.length > 0,
                selectedRostersNum: draftMessageDetail.rosters.length,
                groupsOptionSelected: draftMessageDetail.groups.length > 0,
                uploadOptionSelected: draftMessageDetail.uploadedList.length > 0,
                selectedGroupsNum: draftMessageDetail.groups.length,
                selectedRadioBtn: selectedRadioButton,
                selectedTeams: draftMessageDetail.teams,
                selectedRosters: draftMessageDetail.rosters,
                selectedGroups: draftMessageDetail.groups,
                uploadedList: draftMessageDetail.uploadedList,
                uploadedListName: draftMessageDetail.uploadedListName,
                exclusionList: draftMessageDetail.exclusionList
            });

            setSenderTemplate(this.card, draftMessageDetail.senderTemplate);
            setCardTitle(this.card, draftMessageDetail.title);
            setCardImageLink(this.card, draftMessageDetail.imageLink);
            setCardSummary(this.card, draftMessageDetail.summary);
            setCardAuthor(this.card, draftMessageDetail.author);
            setCardBtn(this.card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink);
            setCardBtn2(this.card, draftMessageDetail.buttonTitle2, draftMessageDetail.buttonLink2);
            setCardBtn3(this.card, draftMessageDetail.buttonTitle3, draftMessageDetail.buttonLink3);
            setCardBtn4(this.card, draftMessageDetail.buttonTitle4, draftMessageDetail.buttonLink4);
            setCardBtn5(this.card, draftMessageDetail.buttonTitle5, draftMessageDetail.buttonLink5);
            setVideoBtn(this.card, draftMessageDetail.videoLink);
            setBoldSummary(this.card, draftMessageDetail.boldSummary);
            setFontSummary(this.card, draftMessageDetail.fontSummary);
            setAlignmentSummary(this.card, draftMessageDetail.alignment);
            setFontSizeSummary(this.card, draftMessageDetail.fontSizeSummary);
            setFontColorSummary(this.card, draftMessageDetail.fontColorSummary);
            setReaction(this.card, draftMessageDetail.surReaction);
            SetReactionQuestion(this.card, draftMessageDetail.reactionQuestion);
            SetFreeText(this.card, draftMessageDetail.surFreeText);
            SetFreeTextQuestion(this.card, draftMessageDetail.freeTextQuestion);
            setYesNo(this.card, draftMessageDetail.surYesNo);
            setYesNoQuestion(this.card, draftMessageDetail.yesNoQuestion);
            setToggleLinktoSurvey(this.card, draftMessageDetail.surLinkToSurvey);
            setLinktoSurvey(this.card, draftMessageDetail.linktoSurvey);

            if (draftMessageDetail.surReaction || draftMessageDetail.surFreeText || draftMessageDetail.surYesNo) {
                setSubmitBtn(this.card, true);
            }

            this.setState({
                page: "CardCreation",
                template: draftMessageDetail.template,
                senderTemplate: draftMessageDetail.senderTemplate,
                title: draftMessageDetail.title,
                summary: draftMessageDetail.summary,
                alignment: draftMessageDetail.alignment,
                fontSummary: draftMessageDetail.fontSummary,
                fontSizeSummary: draftMessageDetail.fontSizeSummary,
                fontColorSummary: draftMessageDetail.fontColorSummary,
                boldSummary: draftMessageDetail.boldSummary,
                btnLink: draftMessageDetail.buttonLink,
                imageLink: draftMessageDetail.imageLink,
                videoLink: draftMessageDetail.videoLink,
                btnTitle: draftMessageDetail.buttonTitle,
                btnTitle2: draftMessageDetail.buttonTitle2,
                btnTitle3: draftMessageDetail.buttonTitle3,
                btnTitle4: draftMessageDetail.buttonTitle4,
                btnTitle5: draftMessageDetail.buttonTitle5,
                btnLink2: draftMessageDetail.buttonLink2,
                btnLink3: draftMessageDetail.buttonLink3,
                btnLink4: draftMessageDetail.buttonLink4,
                btnLink5: draftMessageDetail.buttonLink5,
                surReaction: draftMessageDetail.surReaction,
                reactionQuestion: draftMessageDetail.reactionQuestion,
                surFreeText: draftMessageDetail.surFreeText,
                freeTextQuestion: draftMessageDetail.freeTextQuestion,
                surYesNo: draftMessageDetail.surYesNo,
                yesNoQuestion: draftMessageDetail.yesNoQuestion,
                surLinkToSurvey: draftMessageDetail.surLinkToSurvey,
                linkToSurvey: draftMessageDetail.linkToSurvey,
                author: draftMessageDetail.author,
                allUsersOptionSelected: draftMessageDetail.allUsers,
                loader: false
            }, () => {
                this.updateCard();
            });
        } catch (error) {
            return error;
        }
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    public render(): JSX.Element {
        const isUploadedList = this.state.selectedFileName;
        let uploadList: {} | null | undefined;
        if (isUploadedList) {
            uploadList = this.state.selectedFileName;
        } else {
            uploadList = "No file chosen";
        }
        const hiddenDiv = {
            display: 'block',
        };
        const templates = [
            {
                key: "1", header: "Announcements with image"
            }, {
                key: "2", header: "Announcement with video"
            }, {
                key: "3", header: "Announcement in plain text"
            }, {
                key: "4", header: "Event promotion with sessions to register/join"
            }
        ];
        const FontSize = ["Default", "Small", "Medium", "Large", "ExtraLarge" ];
        const Color = ["Default", "Dark", "Light", "Accent", "Good", "Warning", "Attention"];

            if (this.state.loader) {
                return (
                    <div className="Loader">
                        <Loader />
                    </div>
                );
            } else {
                if (this.state.page === "TemplateSelection") {
                    return (
                        <div className="taskModule">
                            <Flex column className="formContainer1" vAlign="stretch" gap="gap.small">
                                <Flex className="scrollableContent" vAlign="center" hAlign="start">
                                    <Flex.Item>
                                        <Flex column className="formContentContainer">
                                            <Dropdown
                                                className="Templates"
                                                placeholder="Choose a Template"
                                                value={this.state.template}
                                                items={templates}
                                                onChange={this.onTemplateChange}
                                                checkable
                                            />
                                        </Flex>
                                    </Flex.Item>
                                    <br/>
                                    <Flex.Item>
                                        <Flex column>
                                            
                                        </Flex>
                                    </Flex.Item>
                                    <Flex.Item>
                                        <div className="adaptiveCardContainer" style={{ width: "436px"}}>
                                        </div>
                                    </Flex.Item>
                                </Flex>

                                <Flex className="footerContainer" vAlign="end" hAlign="end">
                                    <Flex className="buttonContainer">
                                        <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                    </Flex>
                                </Flex>

                            </Flex>
                        </div>
                    );
                } else if (this.state.page === "CardCreation") {
                    if (this.state.template === "Announcements with image") {
                        return (
                            <div className="taskModule">
                                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                    <Flex className="scrollableContent">
                                        <Flex.Item size="size.half">
                                            <Flex column className="formContentContainer">
                                                <Input className="senderTemplates"
                                                    value={this.state.senderTemplate}
                                                    label={this.localize("Header")}
                                                    placeholder={this.localize("HEADER")}
                                                    onChange={this.onSenderTemplateChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    value={this.state.title}
                                                    label={this.localize("TitleText")}
                                                    placeholder={this.localize("PlaceHolderTitle")}
                                                    onChange={this.onTitleChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Flex vAlign="end">
                                                    <Input fluid className="inputField"
                                                        value={this.state.imageLink}
                                                        label={this.localize("ImageURL")}
                                                        placeholder={this.localize("ImageURL")}
                                                        onChange={this.onImageLinkChanged}
                                                        error={!(this.state.errorImageUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                </Flex>
                                                <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />
                                                <div className="textArea">
                                                    <Text content={this.localize("Summary")} />
                                                    <Flex vAlign="center" hAlign="start">
                                                        <ToggleButtonGroup size="small" value={this.state.boldSummary}>
                                                            <ToggleButton value="Bold" onClick={this.onBoldSummaryChange}>
                                                                <FormatBoldIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <ToggleButtonGroup size="small" value={this.state.fontSummary} >
                                                            <ToggleButton value="Monospace" onClick={this.onFontSummaryChange}>
                                                                <FontDownloadIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <Dropdown className="testDrop"
                                                            items={FontSize}
                                                            placeholder="FontSize"
                                                            onChange={this.onFontSizeChange}
                                                        />
                                                        <Dropdown className="testDrop"
                                                            items={Color}
                                                            placeholder="Color"
                                                            onChange={this.onFontColorChange}
                                                        />
                                                        <Grid item>
                                                            <ToggleButtonGroup value={this.state.alignment} size="small" onChange={this.onAlignmentChange} exclusive>
                                                                <ToggleButton value="left">
                                                                    <FormatAlignLeftIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="center">
                                                                    <FormatAlignCenterIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="right">
                                                                    <FormatAlignRightIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                        </Grid>
                                                    </Flex>
                                                    <TextArea
                                                        autoFocus
                                                        placeholder={this.localize("Summary")}
                                                        value={this.state.summary}
                                                        onChange={this.onSummaryChanged}
                                                        fluid />
                                                </div>

                                                <Input className="inputField"
                                                    value={this.state.author}
                                                    label={this.localize("Author")}
                                                    placeholder={this.localize("Author")}
                                                    onChange={this.onAuthorChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle}
                                                    label={this.localize("ButtonTitle")}
                                                    placeholder={this.localize("ButtonTitle")}
                                                    onChange={this.onBtnTitleChanged}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink}
                                                    label={this.localize("ButtonURL")}
                                                    placeholder={this.localize("ButtonURL")}
                                                    onChange={this.onBtnLinkChanged}
                                                    error={!(this.state.errorButtonUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                                
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surReaction}
                                                            onChange={this.onReactionChanged}
                                                        />
                                                    }
                                                    label="Reaction"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for reaction")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onReactionQuestionChanged}
                                                    value={this.state.reactionQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surFreeText}
                                                            onChange={this.onFreeTextChanged}
                                                        />
                                                    }
                                                    label="Free Text"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for for free text")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onFreeTextQuestionChanged}
                                                    value={this.state.freeTextQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surYesNo}
                                                            onChange={this.onYesNoChanged}
                                                        />
                                                    }
                                                    label="Yes/ No Question"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for Yes / No")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onYesNoQuestionChanged}
                                                    value={this.state.yesNoQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surLinkToSurvey}
                                                            onChange={this.onToggleLinktoSurveyChanged}
                                                        />
                                                    }
                                                    label="Link to Survey"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Survey URL")}
                                                    placeholder={this.localize("Enter URL")}
                                                    onChange={this.onLinktoSurveyChanged}
                                                    value={this.state.linkToSurvey}
                                                    error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />
                                            </Flex>
                                        </Flex.Item>
                                        <Flex.Item size="size.half">
                                            <div className="adaptiveCardContainer">
                                            </div>
                                        </Flex.Item>
                                    </Flex>

                                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                                        <Flex className="buttonContainer">
                                            <Flex.Item push>
                                                <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                            </Flex.Item>
                                            <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                        </Flex>
                                    </Flex>

                                </Flex>
                            </div>
                        );
                    } else if (this.state.template === "Announcement with video") {
                        return (
                            <div className="taskModule">
                                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                    <Flex className="scrollableContent">
                                        <Flex.Item size="size.half">
                                            <Flex column className="formContentContainer">
                                                {/*<Dropdown*/}
                                                {/*    className="senderTemplates"*/}
                                                {/*    placeholder="Choose a Sender Template"*/}
                                                {/*    value={this.state.senderTemplate}*/}
                                                {/*    items={senderTemplates}*/}
                                                {/*    checkable*/}
                                                {/*    onChange={this.onSenderTemplateChanged}*/}
                                                {/*/>*/}
                                                <Input className="senderTemplates"
                                                    value={this.state.senderTemplate}
                                                    label={this.localize("Header")}
                                                    placeholder={this.localize("HEADER")}
                                                    onChange={this.onSenderTemplateChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    value={this.state.title}
                                                    label={this.localize("TitleText")}
                                                    placeholder={this.localize("PlaceHolderTitle")}
                                                    onChange={this.onTitleChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <div className="test2" style={hiddenDiv}>
                                                    <Input fluid className="inputField"
                                                        value={this.state.imageLink}
                                                        label={this.localize("ImageURL")}
                                                        placeholder={this.localize("ImageURL")}
                                                        onChange={this.onImageLinkChanged}
                                                        error={!(this.state.errorImageUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                </div>
                                                <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />

                                                <div className="textArea">
                                                    <Text content={this.localize("Summary")} />
                                                    <Flex vAlign="center" hAlign="start">
                                                        <ToggleButtonGroup size="small" value={this.state.boldSummary}>
                                                            <ToggleButton value="Bold" onClick={this.onBoldSummaryChange}>
                                                                <FormatBoldIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <ToggleButtonGroup size="small" value={this.state.fontSummary} >
                                                            <ToggleButton value="Monospace" onClick={this.onFontSummaryChange}>
                                                                <FontDownloadIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <Dropdown className="testDrop"
                                                            items={FontSize}
                                                            placeholder="FontSize"
                                                            onChange={this.onFontSizeChange}
                                                        />
                                                        <Dropdown className="testDrop"
                                                            items={Color}
                                                            placeholder="Color"
                                                            onChange={this.onFontColorChange}
                                                        />
                                                        <Grid item>
                                                            <ToggleButtonGroup value={this.state.alignment} size="small" onChange={this.onAlignmentChange} exclusive>
                                                                <ToggleButton value="left">
                                                                    <FormatAlignLeftIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="center">
                                                                    <FormatAlignCenterIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="right">
                                                                    <FormatAlignRightIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                        </Grid>
                                                    </Flex>
                                                    <TextArea
                                                        autoFocus
                                                        placeholder={this.localize("Summary")}
                                                        value={this.state.summary}
                                                        onChange={this.onSummaryChanged}
                                                        fluid />
                                                </div>

                                                <Input className="inputField"
                                                    value={this.state.author}
                                                    label={this.localize("Author")}
                                                    placeholder={this.localize("Author")}
                                                    onChange={this.onAuthorChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input fluid className="inputField"
                                                    value={this.state.videoLink}
                                                    label={this.localize("Video URL")}
                                                    placeholder={this.localize("Video URL")}
                                                    onChange={this.onVideoLinkChanged}
                                                    error={!(this.state.errorVideoUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorVideoUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorVideoUrlMessage} />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle}
                                                    label={this.localize("ButtonTitle")}
                                                    placeholder={this.localize("ButtonTitle")}
                                                    onChange={this.onBtnTitleChanged}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink}
                                                    label={this.localize("ButtonURL")}
                                                    placeholder={this.localize("ButtonURL")}
                                                    onChange={this.onBtnLinkChanged}
                                                    error={!(this.state.errorButtonUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surReaction}
                                                            onChange={this.onReactionChanged}
                                                        />
                                                    }
                                                    label="Reaction"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for reaction")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onReactionQuestionChanged}
                                                    value={this.state.reactionQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surFreeText}
                                                            onChange={this.onFreeTextChanged}
                                                        />
                                                    }
                                                    label="Free Text"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for for free text")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onFreeTextQuestionChanged}
                                                    value={this.state.freeTextQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surYesNo}
                                                            onChange={this.onYesNoChanged}
                                                        />
                                                    }
                                                    label="Yes/ No Question"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for Yes / No")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onYesNoQuestionChanged}
                                                    value={this.state.yesNoQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surLinkToSurvey}
                                                            onChange={this.onToggleLinktoSurveyChanged}
                                                        />
                                                    }
                                                    label="Link to Survey"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Survey URL")}
                                                    placeholder={this.localize("Enter URL")}
                                                    onChange={this.onLinktoSurveyChanged}
                                                    value={this.state.linkToSurvey}
                                                    error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />
                                            </Flex>
                                        </Flex.Item>
                                        <Flex.Item size="size.half">
                                            <div className="adaptiveCardContainer">
                                            </div>
                                        </Flex.Item>
                                    </Flex>

                                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                                        <Flex className="buttonContainer">
                                            <Flex.Item push>
                                                <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                            </Flex.Item>
                                            <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                        </Flex>
                                    </Flex>

                                </Flex>
                            </div>
                        );
                    } else if (this.state.template === "Announcement in plain text") {
                        return (
                            <div className="taskModule">
                                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                    <Flex className="scrollableContent">
                                        <Flex.Item size="size.half">
                                            <Flex column className="formContentContainer">
                                                {/*<Dropdown*/}
                                                {/*    className="senderTemplates"*/}
                                                {/*    placeholder="Choose a Sender Template"*/}
                                                {/*    value={this.state.senderTemplate}*/}
                                                {/*    items={senderTemplates}*/}
                                                {/*    checkable*/}
                                                {/*    onChange={this.onSenderTemplateChanged}*/}
                                                {/*/>*/}
                                                <Input className="senderTemplates"
                                                    value={this.state.senderTemplate}
                                                    label={this.localize("Header")}
                                                    placeholder={this.localize("HEADER")}
                                                    onChange={this.onSenderTemplateChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    value={this.state.title}
                                                    label={this.localize("TitleText")}
                                                    placeholder={this.localize("PlaceHolderTitle")}
                                                    onChange={this.onTitleChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />

                                                <div className="textArea">
                                                    <Text content={this.localize("Summary")} />
                                                    <Flex vAlign="center" hAlign="start">
                                                        <ToggleButtonGroup size="small" value={this.state.boldSummary}>
                                                            <ToggleButton value="Bold" onClick={this.onBoldSummaryChange}>
                                                                <FormatBoldIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <ToggleButtonGroup size="small" value={this.state.fontSummary} >
                                                            <ToggleButton value="Monospace" onClick={this.onFontSummaryChange}>
                                                                <FontDownloadIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <Dropdown className="testDrop"
                                                            items={FontSize}
                                                            placeholder="FontSize"
                                                            onChange={this.onFontSizeChange}
                                                        />
                                                        <Dropdown className="testDrop"
                                                            items={Color}
                                                            placeholder="Color"
                                                            onChange={this.onFontColorChange}
                                                        />
                                                        <Grid item>
                                                            <ToggleButtonGroup value={this.state.alignment} size="small" onChange={this.onAlignmentChange} exclusive>
                                                                <ToggleButton value="left">
                                                                    <FormatAlignLeftIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="center">
                                                                    <FormatAlignCenterIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="right">
                                                                    <FormatAlignRightIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                        </Grid>
                                                    </Flex>
                                                    <TextArea
                                                        autoFocus
                                                        placeholder={this.localize("Summary")}
                                                        value={this.state.summary}
                                                        onChange={this.onSummaryChanged}
                                                        fluid />
                                                </div>

                                                <Input className="inputField"
                                                    value={this.state.author}
                                                    label={this.localize("Author")}
                                                    placeholder={this.localize("Author")}
                                                    onChange={this.onAuthorChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle}
                                                    label={this.localize("ButtonTitle")}
                                                    placeholder={this.localize("ButtonTitle")}
                                                    onChange={this.onBtnTitleChanged}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink}
                                                    label={this.localize("ButtonURL")}
                                                    placeholder={this.localize("ButtonURL")}
                                                    onChange={this.onBtnLinkChanged}
                                                    error={!(this.state.errorButtonUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surReaction}
                                                            onChange={this.onReactionChanged}
                                                        />
                                                    }
                                                    label="Reaction"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for reaction")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onReactionQuestionChanged}
                                                    value={this.state.reactionQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surFreeText}
                                                            onChange={this.onFreeTextChanged}
                                                        />
                                                    }
                                                    label="Free Text"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for for free text")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onFreeTextQuestionChanged}
                                                    value={this.state.freeTextQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surYesNo}
                                                            onChange={this.onYesNoChanged}
                                                        />
                                                    }
                                                    label="Yes/ No Question"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for Yes / No")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onYesNoQuestionChanged}
                                                    value={this.state.yesNoQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surLinkToSurvey}
                                                            onChange={this.onToggleLinktoSurveyChanged}
                                                        />
                                                    }
                                                    label="Link to Survey"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Survey URL")}
                                                    placeholder={this.localize("Enter URL")}
                                                    onChange={this.onLinktoSurveyChanged}
                                                    value={this.state.linkToSurvey}
                                                    error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />
                                            </Flex>
                                        </Flex.Item>
                                        <Flex.Item size="size.half">
                                            <div className="adaptiveCardContainer">
                                            </div>
                                        </Flex.Item>
                                    </Flex>

                                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                                        <Flex className="buttonContainer">
                                            <Flex.Item push>
                                                <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                            </Flex.Item>
                                            <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                        </Flex>
                                    </Flex>

                                </Flex>
                            </div>
                        );
                    } else if (this.state.template === "Event promotion with sessions to register/join") {
                        return (
                            <div className="taskModule">
                                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                    <Flex className="scrollableContent">
                                        <Flex.Item size="size.half">
                                            <Flex column className="formContentContainer">
                                                {/*<Dropdown*/}
                                                {/*    className="senderTemplates"*/}
                                                {/*    placeholder="Choose a Sender Template"*/}
                                                {/*    value={this.state.senderTemplate}*/}
                                                {/*    items={senderTemplates}*/}
                                                {/*    checkable*/}
                                                {/*    onChange={this.onSenderTemplateChanged}*/}
                                                {/*/>*/}
                                                <Input className="senderTemplates"
                                                    value={this.state.senderTemplate}
                                                    label={this.localize("Header")}
                                                    placeholder={this.localize("HEADER")}
                                                    onChange={this.onSenderTemplateChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    value={this.state.title}
                                                    label={this.localize("TitleText")}
                                                    placeholder={this.localize("PlaceHolderTitle")}
                                                    onChange={this.onTitleChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <div className="test2" style={hiddenDiv}>
                                                    <Input fluid className="inputField"
                                                        value={this.state.imageLink}
                                                        label={this.localize("ImageURL")}
                                                        placeholder={this.localize("ImageURL")}
                                                        onChange={this.onImageLinkChanged}
                                                        error={!(this.state.errorImageUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                </div>
                                                <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />

                                                <div className="textArea">
                                                    <Text content={this.localize("Summary")} />
                                                    <Flex vAlign="center" hAlign="start">
                                                        <ToggleButtonGroup size="small" value={this.state.boldSummary}>
                                                            <ToggleButton value="Bold" onClick={this.onBoldSummaryChange}>
                                                                <FormatBoldIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <ToggleButtonGroup size="small" value={this.state.fontSummary} >
                                                            <ToggleButton value="Monospace" onClick={this.onFontSummaryChange}>
                                                                <FontDownloadIcon />
                                                            </ToggleButton>
                                                        </ToggleButtonGroup>
                                                        <Dropdown className="testDrop"
                                                            items={FontSize}
                                                            placeholder="FontSize"
                                                            onChange={this.onFontSizeChange}
                                                        />
                                                        <Dropdown className="testDrop"
                                                            items={Color}
                                                            placeholder="Color"
                                                            onChange={this.onFontColorChange}
                                                        />
                                                        <Grid item>
                                                            <ToggleButtonGroup value={this.state.alignment} size="small" onChange={this.onAlignmentChange} exclusive>
                                                                <ToggleButton value="left">
                                                                    <FormatAlignLeftIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="center">
                                                                    <FormatAlignCenterIcon />
                                                                </ToggleButton>
                                                                <ToggleButton value="right">
                                                                    <FormatAlignRightIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                        </Grid>
                                                    </Flex>
                                                    <TextArea
                                                        autoFocus
                                                        placeholder={this.localize("Summary")}
                                                        value={this.state.summary}
                                                        onChange={this.onSummaryChanged}
                                                        fluid />
                                                </div>

                                                <Input className="inputField"
                                                    value={this.state.author}
                                                    label={this.localize("Author")}
                                                    placeholder={this.localize("Author")}
                                                    onChange={this.onAuthorChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle}
                                                    label={this.localize("ButtonTitle")}
                                                    placeholder={this.localize("ButtonTitle")}
                                                    onChange={this.onBtnTitleChanged}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink}
                                                    label={this.localize("ButtonURL")}
                                                    placeholder={this.localize("ButtonURL")}
                                                    onChange={this.onBtnLinkChanged}
                                                    error={!(this.state.errorButtonUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />

                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle2}
                                                    label={this.localize("Button title 2")}
                                                    placeholder={this.localize("Button title 2")}
                                                    onChange={this.onBtnTitleChanged2}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink2}
                                                    label={this.localize("Button URL 2")}
                                                    placeholder={this.localize("Button URL 2")}
                                                    onChange={this.onBtnLinkChanged2}
                                                    error={!(this.state.errorButtonUrlMessage2 === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage2 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage2} />

                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle3}
                                                    label={this.localize("Button title 3")}
                                                    placeholder={this.localize("Button title 3")}
                                                    onChange={this.onBtnTitleChanged3}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink3}
                                                    label={this.localize("Button URL 3")}
                                                    placeholder={this.localize("Button URL 3")}
                                                    onChange={this.onBtnLinkChanged3}
                                                    error={!(this.state.errorButtonUrlMessage3 === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage3 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage3} />

                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle4}
                                                    label={this.localize("Button title 4")}
                                                    placeholder={this.localize("Button title 4")}
                                                    onChange={this.onBtnTitleChanged4}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink4}
                                                    label={this.localize("Button URL 4")}
                                                    placeholder={this.localize("Button URL 4")}
                                                    onChange={this.onBtnLinkChanged4}
                                                    error={!(this.state.errorButtonUrlMessage4 === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage4 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage4} />

                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnTitle5}
                                                    label={this.localize("Button title 5")}
                                                    placeholder={this.localize("Button title 5")}
                                                    onChange={this.onBtnTitleChanged5}
                                                    autoComplete="off"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    value={this.state.btnLink5}
                                                    label={this.localize("Button URL 5")}
                                                    placeholder={this.localize("Button URL 5")}
                                                    onChange={this.onBtnLinkChanged5}
                                                    error={!(this.state.errorButtonUrlMessage5 === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorButtonUrlMessage5 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage5} />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surReaction}
                                                            onChange={this.onReactionChanged}
                                                        />
                                                    }
                                                    label="Reaction"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for reaction")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onReactionQuestionChanged}
                                                    value={this.state.reactionQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surFreeText}
                                                            onChange={this.onFreeTextChanged}
                                                        />
                                                    }
                                                    label="Free Text"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for for free text")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onFreeTextQuestionChanged}
                                                    value={this.state.freeTextQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surYesNo}
                                                            onChange={this.onYesNoChanged}
                                                        />
                                                    }
                                                    label="Yes/ No Question"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Question for Yes / No")}
                                                    placeholder={this.localize("Enter the question")}
                                                    onChange={this.onYesNoQuestionChanged}
                                                    value={this.state.yesNoQuestion}
                                                    autoComplete="off"
                                                />
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            color="primary"
                                                            checked={this.state.surLinkToSurvey}
                                                            onChange={this.onToggleLinktoSurveyChanged}
                                                        />
                                                    }
                                                    label="Link to Survey"
                                                />
                                                <Input className="inputField"
                                                    fluid
                                                    label={this.localize("Survey URL")}
                                                    placeholder={this.localize("Enter URL")}
                                                    onChange={this.onLinktoSurveyChanged}
                                                    value={this.state.linkToSurvey}
                                                    error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />
                                            </Flex>
                                        </Flex.Item>
                                        <Flex.Item size="size.half">
                                            <div className="adaptiveCardContainer">
                                            </div>
                                        </Flex.Item>
                                    </Flex>

                                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                                        <Flex className="buttonContainer">
                                            <Flex.Item push>
                                                <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                            </Flex.Item>
                                            <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                        </Flex>
                                    </Flex>

                                </Flex>
                            </div>
                        );
                    }
                    return (
                        <div></div>
                        );
                }
                else if (this.state.page === "AudienceSelection") {
                    return (
                        <div className="taskModule">
                            <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                <Flex className="scrollableContent">
                                    <Flex.Item size="size.half">
                                        <Flex column className="formContentContainer">
                                            <h3>{this.localize("SendHeadingText")}</h3>
                                            <RadioGroup
                                                className="radioBtns"
                                                checkedValue={this.state.selectedRadioBtn}
                                                onCheckedValueChange={this.onGroupSelected}
                                                vertical={true}
                                                items={[
                                                    {
                                                        name: "teams",
                                                        key: "teams",
                                                        value: "teams",
                                                        label: this.localize("SendToGeneralChannel"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <Dropdown
                                                                        fluid
                                                                        hidden={!this.state.teamsOptionSelected}
                                                                        placeholder={this.localize("SendToGeneralChannelPlaceHolder")}
                                                                        search
                                                                        multiple
                                                                        items={this.getItems()}
                                                                        value={this.state.selectedTeams}
                                                                        onChange={this.onTeamsChange}
                                                                        noResultsMessage={this.localize("NoMatchMessage")}
                                                                    />
                                                                </Flex>
                                                            )
                                                        },
                                                    },
                                                    {
                                                        name: "rosters",
                                                        key: "rosters",
                                                        value: "rosters",
                                                        label: this.localize("SendToRosters"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <Dropdown
                                                                        fluid
                                                                        hidden={!this.state.rostersOptionSelected}
                                                                        placeholder={this.localize("SendToRostersPlaceHolder")}
                                                                        search
                                                                        multiple
                                                                        items={this.getItems()}
                                                                        value={this.state.selectedRosters}
                                                                        onChange={this.onRostersChange}
                                                                        unstable_pinned={this.state.unstablePinned}
                                                                        noResultsMessage={this.localize("NoMatchMessage")}
                                                                    />
                                                                </Flex>
                                                            )
                                                        },
                                                    },
                                                    {
                                                        name: "allUsers",
                                                        key: "allUsers",
                                                        value: "allUsers",
                                                        label: this.localize("SendToAllUsers"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <div className={this.state.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                                                        <div className="noteText">
                                                                            <Text error content={this.localize("SendToAllUsersNote")} />
                                                                        </div>
                                                                    </div>
                                                                </Flex>
                                                            )
                                                        },
                                                    },
                                                    {
                                                        name: "groups",
                                                        key: "groups",
                                                        value: "groups",
                                                        label: this.localize("SendToGroups"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <div className={this.state.groupsOptionSelected && !this.state.groupAccess ? "" : "hide"}>
                                                                        <div className="noteText">
                                                                            <Text error content={this.localize("SendToGroupsPermissionNote")} />
                                                                        </div>
                                                                    </div>
                                                                    <Dropdown
                                                                        fluid
                                                                        className="hideToggle"
                                                                        hidden={!this.state.groupsOptionSelected || !this.state.groupAccess}
                                                                        placeholder={this.localize("SendToGroupsPlaceHolder")}
                                                                        search={this.onGroupSearch}
                                                                        multiple
                                                                        loading={this.state.loading}
                                                                        loadingMessage={this.localize("LoadingText")}
                                                                        items={this.getGroupItems()}
                                                                        value={this.state.selectedGroups}
                                                                        onSearchQueryChange={this.onGroupSearchQueryChange}
                                                                        onChange={this.onGroupsChange}
                                                                        noResultsMessage={this.state.noResultMessage}
                                                                        unstable_pinned={this.state.unstablePinned}
                                                                    />
                                                                    <div className={this.state.groupsOptionSelected && this.state.groupAccess ? "" : "hide"}>
                                                                        <div className="noteText">
                                                                            <Text error content={this.localize("SendToGroupsNote")} />
                                                                        </div>
                                                                    </div>
                                                                </Flex>
                                                            )
                                                        },
                                                    },
                                                    {
                                                        name: "uploadList",
                                                        key: "uploadList",
                                                        value: "uploadList",
                                                        //label: this.localize("SendToUploadList"),
                                                        label: "Send to the list of members of upload file",
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <div
                                                                        hidden={!this.state.uploadedListName}
                                                                    >
                                                                        Filename: {this.state.uploadedListName}
                                                                    </div>
                                                                    <Flex
                                                                        hidden={!this.state.uploadOptionSelected}
                                                                    >
                                                                        <CSVReader
                                                                            isReset={this.state.resetCSVReader}
                                                                            onDrop={this.handleOnDrop}
                                                                            onError={this.handleOnError}
                                                                            addRemoveButton
                                                                            removeButtonColor='#659cef'
                                                                            onRemoveFile={this.handleOnRemoveFile}
                                                                        >
                                                                            <span>{this.renderFileName()}</span>
                                                                            {/*<span>Drop CSV file here or click to upload.</span>*/}
                                                                        </CSVReader>
                                                                    </Flex>
                                                                </Flex>
                                                            )
                                                        },
                                                    }
                                                ]}
                                            >

                                            </RadioGroup>
                                            <Text content={this.localize("Add users to be exclued from communication")} />
                                            {/*<Dropdown*/}
                                            {/*    fluid*/}
                                            {/*    search*/}
                                            {/*    multiple*/}
                                            {/*    placeholder="Search for a person"*/}
                                            {/*    a11ySelectedItemsMessage="Press Delete or Backspace to remove"*/}
                                            {/*/>*/}
                                            {/*noResultsMessage="We couldn't find any matches."*/}
                                            <Input className="inputField"
                                                fluid
                                                value={this.state.exclusionList}
                                                label="Exclusion List (Use ; to separate email)"
                                                placeholder="Type in Email to exclude from this message"
                                                onChange={this.onExclusionListChanged}
                                                autoComplete="off"
                                            />
                                            {/*<div style={{ paddingLeft: '5px' }}>*/}
                                                <FormControlLabel
                                                    control={
                                                        <Switch
                                                            disabled
                                                            color="primary"
                                                        />
                                                    }
                                                    label="Send email to members"
                                                />
                                            {/*</div>*/}
                                            
                                        </Flex>
                                        
                                    </Flex.Item>
                                    <Flex.Item size="size.half">
                                        <div className="adaptiveCardContainer">
                                        </div>
                                    </Flex.Item>
                                </Flex>
                                <Flex className="footerContainer" vAlign="end" hAlign="end">
                                    <Flex className="buttonContainer" gap="gap.small">
                                        <Flex.Item push>
                                            <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                        </Flex.Item>
                                        <Button content={this.localize("SaveAsDraft")} disabled={this.isSaveBtnDisabled()} id="saveBtn" onClick={this.onSave} primary />
                                    </Flex>
                                </Flex>
                            </Flex>
                        </div>
                    );
                } else {
                    return (<div>Error</div>);
                }
            }
        }
    private onGroupSelected = (event: any, data: any) => {
        this.setState({
            selectedRadioBtn: data.value,
            teamsOptionSelected: data.value === 'teams',
            rostersOptionSelected: data.value === 'rosters',
            groupsOptionSelected: data.value === 'groups',
            uploadOptionSelected: data.value === 'uploadList',
            allUsersOptionSelected: data.value === 'allUsers',
            selectedTeams: data.value === 'teams' ? this.state.selectedTeams : [],
            selectedTeamsNum: data.value === 'teams' ? this.state.selectedTeamsNum : 0,
            selectedRosters: data.value === 'rosters' ? this.state.selectedRosters : [],
            selectedRostersNum: data.value === 'rosters' ? this.state.selectedRostersNum : 0,
            selectedGroups: data.value === 'groups' ? this.state.selectedGroups : [],
            selectedGroupsNum: data.value === 'groups' ? this.state.selectedGroupsNum : 0,
            uploadedList: data.value === 'uploadList' ? this.state.uploadedList : [],
            uploadedListName: data.value === 'uploadList' ? this.state.uploadedListName : "",
            resetCSVReader: data.value === 'uploadList' ? false : true
        });
    }

    private isSaveBtnDisabled = () => {
        const teamsSelectionIsValid = (this.state.teamsOptionSelected && (this.state.selectedTeamsNum !== 0)) || (!this.state.teamsOptionSelected);
        const rostersSelectionIsValid = (this.state.rostersOptionSelected && (this.state.selectedRostersNum !== 0)) || (!this.state.rostersOptionSelected);
        const groupsSelectionIsValid = (this.state.groupsOptionSelected && (this.state.selectedGroupsNum !== 0)) || (!this.state.groupsOptionSelected);
        const uploadSelectionIsValid = (this.state.uploadOptionSelected && (this.state.uploadedList.length !== 0) || (!this.state.uploadOptionSelected));
        const nothingSelected = (!this.state.uploadOptionSelected) && (!this.state.teamsOptionSelected) && (!this.state.rostersOptionSelected) && (!this.state.groupsOptionSelected) && (!this.state.allUsersOptionSelected);
        return (!uploadSelectionIsValid || !teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected)
        //return (!teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected)
    }

    private isNextBtnDisabled = () => {
        const title = this.state.title;
        const btnTitle = this.state.btnTitle;
        const btnLink = this.state.btnLink;
        if (this.state.page === "TemplateSelection") {
            return !((this.state.template != ""));
        } else if (this.state.page === "CardCreation") {
            return !(title && ((btnTitle && btnLink) || (!btnTitle && !btnLink)) && (this.state.errorImageUrlMessage === "") && (this.state.errorButtonUrlMessage === ""));
        }
    }

    private getItems = () => {
        const resultedTeams: dropdownItem[] = [];
        if (this.state.teams) {
            let remainingUserTeams = this.state.teams;
            if (this.state.selectedRadioBtn !== "allUsers") {
                if (this.state.selectedRadioBtn === "teams") {
                    this.state.teams.filter(x => this.state.selectedTeams.findIndex(y => y.team.id === x.id) < 0);
                }
                else if (this.state.selectedRadioBtn === "rosters") {
                    this.state.teams.filter(x => this.state.selectedRosters.findIndex(y => y.team.id === x.id) < 0);
                }
            }
            remainingUserTeams.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id
                    }
                });
            });
        }
        return resultedTeams;
    }

    private static MAX_SELECTED_TEAMS_NUM: number = 20;

    private onTeamsChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedTeams: itemsData.value,
            selectedTeamsNum: itemsData.value.length,
            selectedRosters: [],
            selectedRostersNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0,
            uploadedList: [],
            uploadedListName: "",
        })
    }

    private onTemplateChange = (event: any, itemsData: any) => {
        if (itemsData.value.key === "1") {
            this.setState({
                template: itemsData.value.header,
                card:
                {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "Container",
                            "bleed": true,
                            "backgroundImage": {
                                //"url": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAATSURBVChTY7gs6IoHjUpjQYKuAHs0dAUXB9EuAAAAAElFTkSuQmCC",
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
                                    "size": "medium",
                                    "color": "light",
                                    "wrap": true
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 1",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        ,
                        {
                            "type": "",
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
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 2",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 3",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 4",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
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
            }, () => {
                this.updateCard();
            });

        } else if (itemsData.value.key === "2") {
            this.setState({
                template: itemsData.value.header,
                card:
                {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "Container",
                            "bleed": true,
                            "backgroundImage": {
                                //"url": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAATSURBVChTY7gs6IoHjUpjQYKuAHs0dAUXB9EuAAAAAElFTkSuQmCC",
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
                                    "size": "medium",
                                    "color": "light",
                                    "wrap": true
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 2",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        ,
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
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 2",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 3",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 4",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
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
            }, () => {
                this.updateCard();
            });


        } else if (itemsData.value.key === "3") {
            this.setState({
                template: itemsData.value.header,
                card:
                {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "Container",
                            "bleed": true,
                            "backgroundImage": {
                                //"url": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAATSURBVChTY7gs6IoHjUpjQYKuAHs0dAUXB9EuAAAAAElFTkSuQmCC",
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
                                    "size": "medium",
                                    "color": "light",
                                    "wrap": true
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 3",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        ,
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Watch Video",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 1",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 2",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 3",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Button title 4",
                                    "url": "https://adaptivecards.io"
                                }
                            ]
                        },
                        {
                            "type": "",
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
            }, () => {
                this.updateCard();
            });


        } else if (itemsData.value.key === "4") {
            this.setState({
                template: itemsData.value.header,
                card:
                {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "Container",
                            "bleed": true,
                            "backgroundImage": {
                                //"url": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAATSURBVChTY7gs6IoHjUpjQYKuAHs0dAUXB9EuAAAAAElFTkSuQmCC",
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
                                    "size": "medium",
                                    "color": "light",
                                    "wrap": true
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Template 4",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        ,
                        {
                            "type": "",
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
            }, () => {
                this.updateCard();
            });
        }
    }

    private onRostersChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedRosters: itemsData.value,
            selectedRostersNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0,
            uploadedList: [],
            uploadedListName: "",
            resetCSVReader: true
        })
    }

    private onGroupsChange = (event: any, itemsData: any) => {
        this.setState({
            selectedGroups: itemsData.value,
            selectedGroupsNum: itemsData.value.length,
            groups: [],
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedRosters: [],
            selectedRostersNum: 0,
            uploadedList: [],
            resetCSVReader: true
        })
    }

    private onGroupSearch = (itemList: any, searchQuery: string) => {
        const result = itemList.filter(
            (item: { header: string; content: string; }) => (item.header && item.header.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1) ||
                (item.content && item.content.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1),
        )
        return result;
    }

    private onGroupSearchQueryChange = async (event: any, itemsData: any) => {

        if (!itemsData.searchQuery) {
            this.setState({
                groups: [],
                noResultMessage: "",
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
            this.setState({
                loading: false,
                noResultMessage: this.localize("NoMatchMessage"),
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
            // handle event trigger on item select.
            const result = itemsData.items && itemsData.items.find(
                (item: { header: string; }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
            )
            if (result) {
                return;
            }

            this.setState({
                loading: true,
                noResultMessage: "",
            });

            try {
                const query = encodeURIComponent(itemsData.searchQuery);
                const response = await searchGroups(query);
                this.setState({
                    groups: response.data,
                    loading: false,
                    noResultMessage: this.localize("NoMatchMessage")
                });
            }
            catch (error) {
                return error;
            }
        }
    }

    private onSave = () => {
        const selectedTeams: string[] = [];
        const selctedRosters: string[] = [];
        const selectedGroups: string[] = [];
        this.state.selectedTeams.forEach(x => selectedTeams.push(x.team.id));
        this.state.selectedRosters.forEach(x => selctedRosters.push(x.team.id));
        this.state.selectedGroups.forEach(x => selectedGroups.push(x.team.id));

        const draftMessage: IDraftMessage = {
            senderTemplate: this.state.senderTemplate,
            template: this.state.template,
            id: this.state.messageId,
            title: this.state.title,
            imageLink: this.state.imageLink,
            videoLink: this.state.videoLink,
            summary: this.state.summary,
            alignment: this.state.alignment,
            boldSummary: this.state.boldSummary,
            fontSummary: this.state.fontSummary,
            fontSizeSummary: this.state.fontSizeSummary,
            fontColorSummary: this.state.fontColorSummary,
            author: this.state.author,
            buttonTitle: this.state.btnTitle,
            buttonLink: this.state.btnLink,
            buttonTitle2: this.state.btnTitle2,
            buttonLink2: this.state.btnLink2,
            buttonTitle3: this.state.btnTitle3,
            buttonLink3: this.state.btnLink3,
            buttonTitle4: this.state.btnTitle4,
            buttonLink4: this.state.btnLink4,
            buttonTitle5: this.state.btnTitle5,
            buttonLink5: this.state.btnLink5,
            surReaction: this.state.surReaction,
            reactionQuestion: this.state.reactionQuestion,
            surFreeText: this.state.surFreeText,
            freeTextQuestion: this.state.freeTextQuestion,
            surYesNo: this.state.surYesNo,
            yesNoQuestion: this.state.yesNoQuestion,
            surLinkToSurvey: this.state.surLinkToSurvey,
            linkToSurvey: this.state.linkToSurvey,
            teams: selectedTeams,
            rosters: selctedRosters,
            //rosters: ["19:e0dRe1TRSqFur-yWALdrniodjxblg_TeTb-L6Bn0kXQ1@thread.tacv2"],
            groups: selectedGroups,
            uploadedList: this.state.uploadedList,
            uploadedListName: this.state.uploadedListName,
            allUsers: this.state.allUsersOptionSelected,
            exclusionList: this.state.exclusionList
        };
        if (this.state.exists) {
            this.editDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        } else {
            this.postDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        }
    }

    private editDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            await updateDraftNotification(draftMessage);
        } catch (error) {
            return error;
        }
    }

    private postDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            await createDraftNotification(draftMessage);
        } catch (error) {
            throw error;
        }
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    private onNext = (event: any) => {
        if (this.state.page === "TemplateSelection") {
            this.setState({
                page: "CardCreation"
            }, () => {
                this.updateCard();
            });
        } else if (this.state.page === "CardCreation") {
            this.setState({
                page: "AudienceSelection"
            }, () => {
                this.updateCard();
            });
        }
    }

    private onBack = (event: any) => {
        if (this.state.page === "AudienceSelection") {
            this.setState({
                page: "CardCreation"
            }, () => {
                this.updateCard();
            });
        } else if (this.state.page === "CardCreation") {
            this.setState({
                alignment: "left",
                page: "TemplateSelection",
                title: "",
                summary: "",
                fontSummary: "",
                fontSizeSummary: "",
                fontColorSummary: "",
                boldSummary: "",
                author: "",
                btnLink: "",
                btnLink2: "",
                btnLink3: "",
                btnLink4: "",
                btnLink5: "",
                imageLink: "",
                videoLink: "",
                btnTitle: "",
                btnTitle2: "",
                btnTitle3: "",
                btnTitle4: "",
                btnTitle5: "",
                template: "",
                senderTemplate: "",
                surReaction: false,
                reactionQuestion: "",
                surFreeText: false,
                freeTextQuestion: "",
                surYesNo: false,
                yesNoQuestion: "",
                surLinkToSurvey: false,
                linkToSurvey: "",
                card:
                {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "Container",
                            "bleed": true,
                            "backgroundImage": {
                                //"url": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAATSURBVChTY7gs6IoHjUpjQYKuAHs0dAUXB9EuAAAAAElFTkSuQmCC",
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
                                    "size": "medium",
                                    "color": "light",
                                    "wrap": true
                                },
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "Title",
                            "size": "ExtraLarge",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "spacing": "Default",
                            "url": getBaseUrl() + "/image/imagePlaceholder.png",
                            "size": "Stretch",
                            "width": "400px",
                            "altText": "Testing"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Summary",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "wrap": true,
                            "size": "Small",
                            "weight": "Lighter",
                            "text": "- Author"
                        },
                        {
                            "type": "ActionSet",
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "ButtonTitle",
                                    "url": "https://adaptivecards.io/"
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
            }, () => {
                this.updateCard();
            });
        }
    }

    private onSenderTemplateChanged = (event: any, itemsData: any) => {
        let showDefaultCard = (!itemsData.value && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, itemsData.value);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            senderTemplate: itemsData.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onTitleChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !event.target.value && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, event.target.value);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            title: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onImageLinkChanged = (event: any) => {
        let url = event.target.value.toLowerCase();
        if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,"))))) {
            this.setState({
                errorImageUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorImageUrlMessage: ""
            });
        }

        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !event.target.value && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, event.target.value);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            imageLink: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onVideoLinkChanged = (event: any) => {
        let url = event.target.value.toLowerCase();
        if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,"))))) {
            this.setState({
                errorVideoUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorVideoUrlMessage: ""
            });
        }

        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !event.target.value && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, event.target.value);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            videoLink: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onBoldSummaryChange = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (this.state.boldSummary === "Bold") {
            this.setState({
                boldSummary: "",
                card: this.card
            }, () => {
                setBoldSummary(this.card, this.state.boldSummary);
                this.updateCard();
            });
        } else {
            this.setState({
                boldSummary: "Bold",
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                setBoldSummary(this.card, this.state.boldSummary);
                this.updateCard();
            });
        }
        
    }

    private onFontSummaryChange = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (this.state.fontSummary === "Monospace") {
            this.setState({
                fontSummary: "",
                card: this.card
            }, () => {
                setFontSummary(this.card, this.state.fontSummary);
                this.updateCard();
            });
        } else  {
            this.setState({
                fontSummary: "Monospace",
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                setFontSummary(this.card, this.state.fontSummary);
                this.updateCard();
            });
        }
    }

    private onFontSizeChange = (event: any, itemsData: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (itemsData.value === "Default") {
            this.setState({
                fontSizeSummary: "default",
                card: this.card
            }, () => {
                setFontSizeSummary(this.card, this.state.fontSizeSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Small") {
            this.setState({
                fontSizeSummary: "small",
                card: this.card
            }, () => {
                setFontSizeSummary(this.card, this.state.fontSizeSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Medium") {
            this.setState({
                fontSizeSummary: "medium",
                card: this.card
            }, () => {
                setFontSizeSummary(this.card, this.state.fontSizeSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Large") {
            this.setState({
                fontSizeSummary: "large",
                card: this.card
            }, () => {
                setFontSizeSummary(this.card, this.state.fontSizeSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "ExtraLarge") {
            this.setState({
                fontSizeSummary: "extraLarge",
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                setFontSizeSummary(this.card, this.state.fontSizeSummary);
                this.updateCard();
            });
        }
    }

    private onFontColorChange = (event: any, itemsData: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (itemsData.value === "Default") {
            this.setState({
                fontColorSummary: "default",
                card: this.card
            }, () => {
                setFontColorSummary(this.card, this.state.fontColorSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Accent") {
            this.setState({
                fontColorSummary: "accent",
                card: this.card
            }, () => {
                setFontColorSummary(this.card, this.state.fontColorSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Good") {
            this.setState({
                fontColorSummary: "good",
                card: this.card
            }, () => {
                setFontColorSummary(this.card, this.state.fontColorSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Warning") {
            this.setState({
                fontColorSummary: "warning",
                card: this.card
            }, () => {
                setFontColorSummary(this.card, this.state.fontColorSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Attention") {
            this.setState({
                fontColorSummary: "attention",
                card: this.card
            }, () => {
                setFontColorSummary(this.card, this.state.fontColorSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Light") {
            this.setState({
                fontColorSummary: "light",
                card: this.card
            }, () => {
                setFontColorSummary(this.card, this.state.fontColorSummary);
                this.updateCard();
            });
        } else if (itemsData.value === "Dark") {
            this.setState({
                fontColorSummary: "dark",
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                setFontColorSummary(this.card, this.state.fontColorSummary);
                this.updateCard();
            });
        }
    }

    private onAlignmentChange = (event: any, onAlignmentChange: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (onAlignmentChange === "left") {
            this.setState({
                alignment: "left",
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                setAlignmentSummary(this.card, this.state.alignment);
                this.updateCard();
            });
        } else if (onAlignmentChange === "center") {
            this.setState({
                alignment: "center",
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                setAlignmentSummary(this.card, this.state.alignment);
                this.updateCard();
            });
        } else if (onAlignmentChange === "right") {
            this.setState({
                alignment: "right",
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                setAlignmentSummary(this.card, this.state.alignment);
                this.updateCard();
            });
        }
    }

    private onSummaryChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !event.target.value && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, event.target.value);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            summary: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onAuthorChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !event.target.value && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, event.target.value);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            author: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onBtnTitleChanged = (event: any) => {
        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnTitle2 && !this.state.btnLink && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (event.target.value && this.state.btnLink) {
            setCardBtn(this.card, event.target.value, this.state.btnLink);
            this.setState({
                btnTitle: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn(this.card, event.target.value, this.state.btnLink);
            this.setState({
                btnTitle: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnTitleChanged2 = (event: any) => {
        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnTitle && !this.state.btnLink && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (event.target.value && this.state.btnLink2) {
            setCardBtn2(this.card, event.target.value, this.state.btnLink2);
            this.setState({
                btnTitle2: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn2(this.card, event.target.value, this.state.btnLink2);
            this.setState({
                btnTitle2: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnTitleChanged3 = (event: any) => {
        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (event.target.value && this.state.btnLink3) {
            setCardBtn3(this.card, event.target.value, this.state.btnLink3);
            this.setState({
                btnTitle3: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn3(this.card, event.target.value, this.state.btnLink3);
            this.setState({
                btnTitle3: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnTitleChanged4 = (event: any) => {
        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (event.target.value && this.state.btnLink4) {
            setCardBtn4(this.card, event.target.value, this.state.btnLink4);
            this.setState({
                btnTitle4: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn4(this.card, event.target.value, this.state.btnLink4);
            this.setState({
                btnTitle4: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnTitleChanged5 = (event: any) => {
        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (event.target.value && this.state.btnLink5) {
            setCardBtn5(this.card, event.target.value, this.state.btnLink5);
            this.setState({
                btnTitle5: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn5(this.card, event.target.value, this.state.btnLink5);
            this.setState({
                btnTitle5: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
        }

        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnTitle2 && !this.state.btnLink2 && !event.target.value && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (this.state.btnTitle && event.target.value) {
            setCardBtn(this.card, this.state.btnTitle, event.target.value);
            this.setState({
                btnLink: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn(this.card, this.state.btnTitle, event.target.value);
            this.setState({
                btnLink: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged2 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage2: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage2: ""
            });
        }

        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnTitle2 && !this.state.btnLink && !event.target.value && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (this.state.btnTitle2 && event.target.value) {
            setCardBtn2(this.card, this.state.btnTitle2, event.target.value);
            this.setState({
                btnLink2: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn2(this.card, this.state.btnTitle2, event.target.value);
            this.setState({
                btnLink2: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged3 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage3: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage3: ""
            });
        }

        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnTitle2 && !this.state.btnLink && !this.state.btnLink2 && !event.target.value && !this.state.btnTitle3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (this.state.btnTitle3 && event.target.value) {
            setCardBtn3(this.card, this.state.btnTitle3, event.target.value);
            this.setState({
                btnLink3: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn3(this.card, this.state.btnTitle3, event.target.value);
            this.setState({
                btnLink3: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged4 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage4: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage4: ""
            });
        }

        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnTitle2 && !this.state.btnLink && !this.state.btnLink2 && !this.state.btnLink3 && !event.target.value && !this.state.btnTitle3 && !this.state.btnTitle4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (this.state.btnTitle4 && event.target.value) {
            setCardBtn4(this.card, this.state.btnTitle4, event.target.value);
            this.setState({
                btnLink4: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn4(this.card, this.state.btnTitle4, event.target.value);
            this.setState({
                btnLink4: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged5 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage5: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage5: ""
            });
        }

        const showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnTitle2 && !this.state.btnLink && !this.state.btnLink2 && !this.state.btnLink3 && !this.state.btnLink4 && !event.target.value && !this.state.btnTitle3 && !this.state.btnTitle4 && !this.state.btnTitle5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
		setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        if (this.state.btnTitle4 && event.target.value) {
            setCardBtn5(this.card, this.state.btnTitle5, event.target.value);
            this.setState({
                btnLink5: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            //delete this.card.actions;
            setCardBtn5(this.card, this.state.btnTitle5, event.target.value);
            this.setState({
                btnLink5: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onReactionChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            surReaction: !this.state.surReaction,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            setReaction(this.card, this.state.surReaction);
            if (this.state.surReaction === true || this.state.surFreeText === true || this.state.surYesNo === true) {
                setSubmitBtn(this.card, true);
            } else {
                setSubmitBtn(this.card, false);
            }
            this.updateCard();
        });
    }

    private onReactionQuestionChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !event.target.value && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, event.target.value);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            reactionQuestion: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onFreeTextChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            surFreeText: !this.state.surFreeText,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            SetFreeText(this.card, this.state.surFreeText);

            if (this.state.surReaction === true || this.state.surFreeText === true || this.state.surYesNo === true) {
                setSubmitBtn(this.card, true);
            } else {
                setSubmitBtn(this.card, false);
            }
            this.updateCard();
        });
    }

    private onFreeTextQuestionChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !event.target.value);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, event.target.value);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            freeTextQuestion: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onYesNoChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            surYesNo: !this.state.surYesNo,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            setYesNo(this.card, this.state.surYesNo);

            if (this.state.surReaction === true || this.state.surFreeText === true || this.state.surYesNo === true) {
                setSubmitBtn(this.card, true);
            } else {
                setSubmitBtn(this.card, false);
            }
            this.updateCard();
        });
    }

    private onYesNoQuestionChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion && !event.target.value);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, event.target.value);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            yesNoQuestion: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onToggleLinktoSurveyChanged = (event: any) => {
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion && !this.state.yesNoQuestion && !this.state.linkToSurvey);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, this.state.linkToSurvey);
        this.setState({
            surLinkToSurvey: !this.state.surLinkToSurvey,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            setToggleLinktoSurvey(this.card, this.state.surLinkToSurvey);
            this.updateCard();
        });
    }

    private onLinktoSurveyChanged = (event: any) => {
        let url = event.target.value.toLowerCase();
        if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,"))))) {
            this.setState({
                errorLinkToSurveyUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorLinkToSurveyUrlMessage: ""
            });
        }
        let showDefaultCard = (!this.state.senderTemplate && !this.state.title && !this.state.imageLink && !this.state.videoLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2 && !this.state.btnTitle3 && !this.state.btnLink3 && !this.state.btnTitle4 && !this.state.btnLink4 && !this.state.btnTitle5 && !this.state.btnLink5 && !this.state.reactionQuestion && !this.state.freeTextQuestion && !this.state.yesNoQuestion && !event.target.value);
        setSenderTemplate(this.card, this.state.senderTemplate);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        setCardBtn2(this.card, this.state.btnTitle2, this.state.btnLink2);
        setCardBtn3(this.card, this.state.btnTitle3, this.state.btnLink3);
        setCardBtn4(this.card, this.state.btnTitle4, this.state.btnLink4);
        setCardBtn5(this.card, this.state.btnTitle5, this.state.btnLink5);
        setVideoBtn(this.card, this.state.videoLink);
        SetReactionQuestion(this.card, this.state.reactionQuestion);
        SetFreeTextQuestion(this.card, this.state.freeTextQuestion);
        setYesNoQuestion(this.card, this.state.yesNoQuestion);
        setLinktoSurvey(this.card, event.target.value);
        this.setState({
            linkToSurvey: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private handleSelectedFile = (event: any) => {
        //Papa.parse([])
        if (event.target.files[0]) {
            this.setState({
                selectedFile: event.target.files[0],
                selectedFileName: event.target.files[0].name
            });
        }
        
        //console.log(event.target.files[0]);
    }

    private resetCSVReader = () => {

    }

    private handleOnDrop = (data: any, file: any) => {
        //console.log('---------------------------')
        for (let i = 0; i < data.length - 1; i++) {
            //console.log(JSON.stringify(data[i].data).replace(/["+/g, '').replace(/\[|\]|"/g, ""));
            this.setState({
                uploadedList: this.state.uploadedList.concat(JSON.stringify(data[i].data).replace(/['"]"+/g, '').replace(/\[|\]|"/g, "")),
                uploadedListName: file.name,
                selectedGroups: [],
                selectedGroupsNum: 0,
                groups: [],
                selectedTeams: [],
                selectedTeamsNum: 0,
                selectedRosters: [],
                resetCSVReader: false
            })
        }
        //console.log('---------------------------')
        //console.log(file.name)
    }

    private handleOnError = (err: any , file: any, inputElem:any , reason:any) => {
        console.log(err)
    }

    private handleOnRemoveFile = (data: any) => {
        //console.log('---------------------------')
        //console.log(data)
        //console.log('---------------------------')
    }

    private testGetUser = async () => {
        try {
            const response = await getUsers();
            //this.setState({
            //    teams: response.data
            //});
            //console.log(response.data)
        } catch (error) {
            return error;
        }
    }

    private testConsole = () => {
        console.log(this.state.exclusionList)
    }

    private onExclusionListChanged = (event: any) => {
        this.setState({
            exclusionList: event.target.value
        })
    }

    private renderFileName = () => {
        if (this.state.uploadedListName) {
            return (
                this.state.uploadedListName
                ); 
        } else {
            return (
                "Drop CSV file here or click to upload."
                );
        }
    }

    private updateCard = () => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(this.state.card);
        const renderedCard = adaptiveCard.render();
        const container = document.getElementsByClassName('adaptiveCardContainer')[0].firstChild;
        if (container != null) {
            container.replaceWith(renderedCard);
        } else {
            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
        }
        //const link = this.state.btnLink;
        //adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
        adaptiveCard.onExecuteAction = function (action) {
            if (action instanceof AdaptiveCards.OpenUrlAction) {
                window.open(action.url);
            }
            else if (action instanceof AdaptiveCards.SubmitAction) {
                // action for Action.Submit
            }
            else {
                // <unknown>
            }
        }
    }
}

const newMessageWithTranslation = withTranslation()(NewMessage);
export default newMessageWithTranslation;
