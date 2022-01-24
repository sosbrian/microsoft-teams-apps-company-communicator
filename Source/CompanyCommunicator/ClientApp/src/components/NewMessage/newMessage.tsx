// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as AdaptiveCards from "adaptivecards";
import { Button, Loader, Dropdown, Text, Flex, Input, TextArea, RadioGroup, FlexItem, Checkbox, Datepicker } from '@fluentui/react-northstar'
import { FilesUploadIcon } from '@fluentui/react-icons-northstar'
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
import { Switch, FormControlLabel, Grid } from '@material-ui/core';
import { ToggleButton, ToggleButtonGroup } from '@material-ui/lab';
import FormatBoldIcon from '@material-ui/icons/FormatBold';
import FontDownloadIcon from '@material-ui/icons/FontDownload';
import FormatAlignLeftIcon from '@material-ui/icons/FormatAlignLeft';
import FormatAlignCenterIcon from '@material-ui/icons/FormatAlignCenter';
import FormatAlignRightIcon from '@material-ui/icons/FormatAlignRight';
import { CSVReader } from 'react-papaparse';
import Resizer from 'react-image-file-resizer';
import { BlobServiceClient, ContainerClient } from '@azure/storage-blob';

//hours to be chosen when scheduling messages
const hours = ["00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
    "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23",
];

//minutes to be chosen when scheduling messages
const minutes = ["00", "05", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55",
];

//coeficient to round dates to the next 5 minutes
const coeff = 1000 * 60 * 5;

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
    priLanguage: string,
    secLanguage: string,
    senderTemplate: string,
    secSenderTemplate: string,
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
    secTitle: string,
    secImageLink?: string,
    secVideoLink?: string,
    secAlignment?: string,
    secSummary?: string,
    secBoldSummary?: string,
    secFontSummary?: string,
    secFontSizeSummary?: string,
    secFontColorSummary?: string,
    secAuthor: string,
    secButtonTitle?: string,
    secButtonLink?: string,
    secButtonTitle2?: string,
    secButtonLink2?: string,
    secButtonTitle3?: string,
    secButtonLink3?: string,
    secButtonTitle4?: string,
    secButtonLink4?: string,
    secButtonTitle5?: string,
    secButtonLink5?: string,
    secSurReaction?: boolean,
    secReactionQuestion?: string,
    secSurFreeText?: boolean,
    secFreeTextQuestion?: string,
    secSurYesNo?: boolean,
    secYesNoQuestion?: string,
    secSurLinkToSurvey?: boolean,
    secLinkToSurvey?: string,
    teams: any[],
    rosters: any[],
    groups: any[],
    uploadedList: any[],
    uploadedListName: string,
    emailOption: boolean,
    exclusionList: string,
    allUsers: boolean,
    isScheduled: boolean, // indicates if the message is scheduled
    ScheduledDate: Date, // stores the scheduled date
    isExpirySet: boolean, // indicates if the expiry date is set
    expiryDate: Date, // stores the expiry date
    isExpiredContentErased: boolean // indicates if the content is erased due to expiry date is in due course
}

export interface formState {
    template: string,
    senderTemplate: string,
    secSenderTemplate: string,
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
    secTitle: string,
    secImageLink?: string,
    secVideoLink?: string,
    secAlignment?: string,
    secSummary?: string,
    secBoldSummary?: string,
    secFontSummary?: string,
    secFontSizeSummary?: string,
    secFontColorSummary?: string,
    secAuthor: string,
    secbtnTitle?: string,
    secbtnLink?: string,
    secbtnTitle2?: string,
    secbtnLink2?: string,
    secbtnTitle3?: string,
    secbtnLink3?: string,
    secbtnTitle4?: string,
    secbtnLink4?: string,
    secbtnTitle5?: string,
    secbtnLink5?: string,
    secSurReaction?: boolean,
    secReactionQuestion?: string,
    secSurFreeText?: boolean,
    secFreeTextQuestion?: string,
    secSurYesNo?: boolean,
    secYesNoQuestion?: string,
    secSurLinkToSurvey?: boolean,
    secLinkToSurvey?: string,
    card?: any,
    page: string,
    language: string,
    priLanguage: string,
    secLanguage: string,
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
    emailOption: boolean,
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
    video: any,
    selectedSchedule: boolean, //status of the scheduler checkbox
    scheduledDate: string, //stores the scheduled date in string format
    DMY: Date, //scheduled date in date format
    DMYHour: string, //hour selected
    DMYMins: string, //mins selected
    futuredate: boolean, //if the date is in the future (valid schedule)
    isExpirySet: boolean, //status of "the expiry date is set" checkbox
    expiryDate: string, //stores the scheduled date in string format
    expiryDMY: Date, //scheduled date in date format
    expiryDMYHour: string, //hour selected
    expiryDMYMins: string, //mins selected
    expiryfuturedate: boolean //if the expiry date is in the future (valid schedule)
}

export interface INewMessageProps extends RouteComponentProps, WithTranslation {
    getDraftMessagesList?: any;
}

class NewMessage extends React.Component<INewMessageProps, formState> {
    readonly localize: TFunction;
    private card: any;
    fileInput: any;


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
        var TempDate = this.getRoundedDate(5, this.getDateObject()); //get the current date
        //this.setDefaultCard(this.card);
        this.state = {
            video: null,
            template: "",
            senderTemplate: "",
            secSenderTemplate: "",
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
            secTitle: "",
            secImageLink: "",
            secVideoLink: "",
            secAlignment: "left",
            secSummary: "",
            secBoldSummary: "",
            secFontSummary: "",
            secFontSizeSummary: "",
            secFontColorSummary: "",
            secAuthor: "",
            secbtnTitle: "",
            secbtnLink: "",
            secbtnTitle2: "",
            secbtnLink2: "",
            secbtnTitle3: "",
            secbtnLink3: "",
            secbtnTitle4: "",
            secbtnLink4: "",
            secbtnTitle5: "",
            secbtnLink5: "",
            secSurReaction: false,
            secReactionQuestion: "",
            secSurFreeText: false,
            secFreeTextQuestion: "",
            secSurYesNo: false,
            secYesNoQuestion: "",
            secSurLinkToSurvey: false,
            secLinkToSurvey: "",
            card: this.card,
            page: "TemplateSelection",
            language: "Primary",
            priLanguage: "English",
            secLanguage: "",
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
            emailOption: false,
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
            selectedSchedule: false, //scheduler option is disabled by default
            scheduledDate: TempDate.toUTCString(), //current date in UTC string format
            DMY: TempDate, //current date in Date format
            DMYHour: this.getDateHour(TempDate.toUTCString()), //initialize with the current hour (rounded up)
            DMYMins: this.getDateMins(TempDate.toUTCString()), //initialize with the current minute (rounded up)
            futuredate: false, //by default the date is not in the future
            isExpirySet: false, //expiry option is disabled by default
            expiryDate: TempDate.toUTCString(), //current date in UTC string format
            expiryDMY: TempDate, //current date in Date format
            expiryDMYHour: this.getDateHour(TempDate.toUTCString()), //initialize with the current hour (rounded up)
            expiryDMYMins: this.getDateMins(TempDate.toUTCString()), //initialize with the current minute (rounded up)
            expiryfuturedate: false, //by default the date is not in the future
            resetCSVReader: false
        }
        this.fileInput = React.createRef();
        this.handleImageSelection = this.handleImageSelection.bind(this);
        this.handleSecImageSelection = this.handleSecImageSelection.bind(this);
        this.uploadVideo = this.uploadVideo.bind(this);
        this.onVideoUpload = this.onVideoUpload.bind(this);
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
                console.log("id" + id + "params" + params);
                this.getItem(id).then(() => {
                    console.log("after getItem");
                    const selectedTeams = this.makeDropdownItemList(this.state.selectedTeams, this.state.teams);
                    const selectedRosters = this.makeDropdownItemList(this.state.selectedRosters, this.state.teams);
                    this.setState({
                        exists: true,
                        messageId: id,
                        selectedTeams: selectedTeams,
                        selectedRosters: selectedRosters,
                        selectedSchedule: this.state.selectedSchedule,
                        scheduledDate: this.state.scheduledDate,
                        DMY: this.getDateObject(this.state.scheduledDate),
                        DMYHour: this.getDateHour(this.state.scheduledDate),
                        DMYMins: this.getDateMins(this.state.scheduledDate),
                        isExpirySet: this.state.isExpirySet,
                        expiryDate: this.state.expiryDate,
                        expiryDMY: this.getDateObject(this.state.expiryDate),
                        expiryDMYHour: this.getDateHour(this.state.expiryDate),
                        expiryDMYMins: this.getDateMins(this.state.expiryDate),
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
                    console.log("this: " + this.state.card);
                    document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                    if (this.state.btnLink) {
                        let link = this.state.btnLink;
                        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                    }
                })
            }
        });
    }

    //function to handle the secondary language layout
    private switchLanguage = (event: any, itemsData: any) => {
        //console.log(event)
        if (itemsData.value === "Primary") {
            this.setState({
                language: "Primary"
            }, () => {
                console.log("Language: " + this.state.language);
            });
        } else if (itemsData.value === "Secondary") {
            this.setState({
                language: "Secondary"
            }, () => {
                console.log("Language: " + this.state.language);
            });
        }
    }

    //function to handle the secondary language layout
    private selectLanguage = (event: any, itemsData: any) => {
        if(this.state.language === "Primary"){
            this.setState({
                priLanguage: itemsData.value
            });
        } else if (this.state.language === "Secondary") {
            this.setState({
                secLanguage: itemsData.value
            });
        }
    }

    //function to handle the selection of the OS file upload box
    private handleImageSelection() {
        //get the first file selected
        const file = this.fileInput.current.files[0];
        if (file) { //if we have a file
            //resize the image to fit in the adaptivecard
            var cardsize = JSON.stringify(this.card).length;
            Resizer.imageFileResizer(file, 400, 400, 'JPEG', 80, 0,
                uri => {
                    if (uri.toString().length < 30720) {
                        //everything is ok with the image, lets set it on the card and update
                        setCardImageLink(this.card, uri.toString());
                        this.updateCard();
                        //lets set the state with the image value
                        this.setState({
                            imageLink: uri.toString()
                        }
                        );
                        console.log(this.state.imageLink)
                    } else {
                        var errormsg = this.localize("ErrorImageTooBig") + " " + this.localize("ErrorImageTooBigSize") + " " + (30720 - cardsize) + " bytes.";
                        //images bigger than 32K cannot be saved, set the error message to be presented
                        this.setState({
                            errorImageUrlMessage: errormsg
                        });
                        console.log(this.state.imageLink)
                    }

                },
                'base64'); //we need the image in base64
        }
    }

    //Function calling a click event on a hidden file input
    private handleUploadClick = (event: any) => {
        //reset the error message and the image link as the upload will reset them potentially
        this.setState({
            errorImageUrlMessage: "",
            imageLink: ""
        });
        //fire the fileinput click event and run the handleimageselection function
        this.fileInput.current.click();
    };

    //function to handle the selection of the OS file upload box
    private handleSecImageSelection() {
        //get the first file selected
        const file = this.fileInput.current.files[0];
        if (file) { //if we have a file
            //resize the image to fit in the adaptivecard
            var cardsize = JSON.stringify(this.card).length;
            Resizer.imageFileResizer(file, 400, 400, 'JPEG', 80, 0,
                uri => {
                    if (uri.toString().length < 30720) {
                        //lets set the state with the image value
                        this.setState({
                            secImageLink: uri.toString()
                        }
                        );
                        console.log(this.state.secImageLink)
                    } else {
                        var errormsg = this.localize("ErrorImageTooBig") + " " + this.localize("ErrorImageTooBigSize") + " " + (30720 - cardsize) + " bytes.";
                        //images bigger than 32K cannot be saved, set the error message to be presented
                        this.setState({
                            errorImageUrlMessage: errormsg
                        });
                        console.log(this.state.secImageLink)
                    }

                },
                'base64'); //we need the image in base64
        }
    }

    //Function calling a click event on a hidden file input
    private handleSecUploadClick = (event: any) => {
        //reset the error message and the image link as the upload will reset them potentially
        this.setState({
            errorImageUrlMessage: "",
            secImageLink: ""
        });
        //fire the fileinput click event and run the handleimageselection function
        this.fileInput.current.click();
    };

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
            console.log("teams: " + response);
            this.setState({
                teams: response.data
            });
        } catch (error) {
            return error;
        }
    }

    private onVideoUpload = (event: any) => {
        this.setState({
            video: event.target.files[0]
        });
    }
    
    private uploadVideo = async() => {
        let storageAccountName = "45u3yv4vigkqc";
        let sasToken = "sv=2020-08-04&ss=b&srt=sco&sp=rwlacix&se=2022-01-21T10:03:35Z&st=2022-01-19T02:03:35Z&spr=https&sig=h%2Frncoc%2F65mhdwxhwKUamZmJcAkV%2F18R%2F7cMSDm0o%2FU%3D";
        //let storageAccountName = "qabra5qtfyb2w";
        //let sasToken = "sv=2020-08-04&ss=b&srt=sco&sp=rwlacix&se=2023-01-19T19:24:27Z&st=2022-01-19T11:24:27Z&spr=https&sig=NRfFQM%2F4wF9EKGpjWYCwwee%2FgUKhxnDq0qMHptaFtBU%3D";

        var today = new Date();
        const newVideo = new File(
            [this.state.video],
            `${today.getFullYear()}${today.getMonth()}${today.getDate()}${today.getHours()}${today.getMinutes()}${today.getSeconds()}-${this.state.video.name}`,
            {type: this.state.video.type}
        );
    
        this.setState({
            video: newVideo
        }, () => {
            console.log(this.state.video)
        });
    
        const blobService = new BlobServiceClient(
            `https://${storageAccountName}.blob.core.windows.net/?${sasToken}`
        );
    
        const containerClient = blobService.getContainerClient('files');
        await containerClient.createIfNotExists({
            access: 'container',
        });
        
        const blobClient = containerClient.getBlockBlobClient(this.state.video.name);
    
        const options = { blobHTTPHeaders: { blobContentType: this.state.video.type }};
    
        await blobClient.uploadBrowserData(this.state.video, options);
    
        // let blobUrl = blobClient.Uri.AbsoluteUri;
        let blobUrl = blobClient.url;
        //setVideoBtn(this.card, blobUrl);
        //this.updateCard();
        console.log(blobUrl);
        if (this.state.language === "Primary"){
            this.setState({
                videoLink: blobUrl
            });
        } else if (this.state.language === "Secondary"){
            this.setState({
                secVideoLink: blobUrl
            });
        }
        
    }

    private printCon = (e: any) => {
        console.log("Pri: " + this.state.priLanguage)
        console.log("Sec: " + this.state.secLanguage)
    }

    private getGroupItems() {
        if (this.state.groups) {
            return this.makeDropdownItems(this.state.groups);
        }
        const dropdownItems: dropdownItem[] = [];
        return dropdownItems;
    }

    private setGroupAccess = async () => {
        console.log("group access");
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
            const response = await getGroups(id);
            this.setState({
                groups: response.data
            });
            
        }
        catch (error) {
            return error;
        }
    }

    private getItem = async (id: number) => {
        try {
            console.log("get item");
            const response = await getDraftNotification(id);
            console.log("get item response: " + response)
            const draftMessageDetail = response.data;
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
                emailOption: draftMessageDetail.emailOption,
                exclusionList: draftMessageDetail.exclusionList,
                selectedSchedule: draftMessageDetail.isScheduled,
                scheduledDate: draftMessageDetail.scheduledDate,
                isExpirySet: draftMessageDetail.isExpirySet,
                expiryDate: draftMessageDetail.expiryDate,
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
                priLanguage: draftMessageDetail.priLanguage,
                secLanguage: draftMessageDetail.secLanguage,
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
                secSenderTemplate: draftMessageDetail.secSenderTemplate,
                secTitle: draftMessageDetail.secTitle,
                secSummary: draftMessageDetail.secSummary,
                secAlignment: draftMessageDetail.secAlignment,
                secFontSummary: draftMessageDetail.secFontSummary,
                secFontSizeSummary: draftMessageDetail.secFontSizeSummary,
                secFontColorSummary: draftMessageDetail.secFontColorSummary,
                secBoldSummary: draftMessageDetail.secBoldSummary,
                secbtnLink: draftMessageDetail.secButtonLink,
                secImageLink: draftMessageDetail.secImageLink,
                secVideoLink: draftMessageDetail.secVideoLink,
                secbtnTitle: draftMessageDetail.secButtonTitle,
                secbtnTitle2: draftMessageDetail.secButtonTitle2,
                secbtnTitle3: draftMessageDetail.secButtonTitle3,
                secbtnTitle4: draftMessageDetail.secButtonTitle4,
                secbtnTitle5: draftMessageDetail.secButtonTitle5,
                secbtnLink2: draftMessageDetail.secButtonLink2,
                secbtnLink3: draftMessageDetail.secButtonLink3,
                secbtnLink4: draftMessageDetail.secButtonLink4,
                secbtnLink5: draftMessageDetail.secButtonLink5,
                secSurReaction: draftMessageDetail.secSurReaction,
                secReactionQuestion: draftMessageDetail.secReactionQuestion,
                secSurFreeText: draftMessageDetail.secSurFreeText,
                secFreeTextQuestion: draftMessageDetail.secFreeTextQuestion,
                secSurYesNo: draftMessageDetail.secSurYesNo,
                secYesNoQuestion: draftMessageDetail.secYesNoQuestion,
                secSurLinkToSurvey: draftMessageDetail.secSurLinkToSurvey,
                secLinkToSurvey: draftMessageDetail.secLinkToSurvey,
                secAuthor: draftMessageDetail.secAuthor,
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
        // const language = [
        //     {
        //         key: "eng", header: "English"
        //     }, {
        //         key: "bur", header: "Burmese"
        //     }, {
        //         key: "cht", header: "Traditional Chinese"
        //     }, {
        //         key: "chs", header: "Simplified Chinese"
        //     }, {
        //         key: "may", header: "Malay"
        //     }, {
        //         key: "fil", header: "Filipino"
        //     }, {
        //         key: "ind", header: "Indonesian"
        //     }, {
        //         key: "kor", header: "Korea"
        //     }, {
        //         key: "vie", header: "Vietnamese"
        //     }
        // ];
        const language = ["English", "Burmese", "Traditional Chinese", "Simplified Chinese", "Malay", "Filipino", "Indonesian", "Korea", "Vietnamese"];
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
                        if (this.state.language === "Primary") {
                            return (
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")} />
                                                    <Dropdown
                                                        value={this.state.priLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        className="testDrop" 
                                                    />
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
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Input fluid className="inputField"
                                                            value={this.state.imageLink}
                                                            label={this.localize("ImageURL")}
                                                            placeholder={this.localize("ImageURL")}
                                                            onChange={this.onImageLinkChanged}
                                                            error={!(this.state.errorImageUrlMessage === "")}
                                                            autoComplete="off"
                                                        />
                                                        <input type="file" accept="image/"
                                                            style={{ display: 'none' }}
                                                            onChange={this.handleImageSelection}
                                                            ref={this.fileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={this.handleUploadClick}
                                                                style={{ width: '10px' }}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title="UploadImage"
                                                            />
                                                        </Flex.Item>
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
                                                                value={this.state.fontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onFontColorChange}
                                                                value={this.state.fontColorSummary}
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

                                                    {/*<FormControlLabel*/}
                                                    {/*    control={*/}
                                                    {/*        <Switch*/}
                                                    {/*            color="primary"*/}
                                                    {/*            checked={this.state.selectedSchedule}*/}
                                                    {/*            onChange={this.onScheduleSelected}*/}
                                                    {/*        />*/}
                                                    {/*    }*/}
                                                    {/*    label={this.localize("ScheduledSend")}*/}
                                                    {/*/>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<FormControlLabel*/}
                                                    {/*    control={*/}
                                                    {/*        <Switch*/}
                                                    {/*            color="primary"*/}
                                                    {/*            checked={this.state.isExpirySet}*/}
                                                    {/*            onChange={this.onIsExpirySetSelected}*/}
                                                    {/*        />*/}
                                                    {/*    }*/}
                                                    {/*    label={this.localize("ExpirySend")}*/}
                                                    {/*/>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}



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
                        } else if (this.state.language === "Secondary") {
                            return (
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    {/*<Dropdown*/}
                                                    {/*    value={this.state.language}*/}
                                                    {/*    items={language}*/}
                                                    {/*    onChange={this.switchLanguage}*/}
                                                    {/*    checkable*/}
                                                    {/*/>*/}
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")} />
                                                    <Dropdown
                                                        value={this.state.secLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        className="testDrop" 
                                                    />
                                                    <Input className="senderTemplates"
                                                        value={this.state.secSenderTemplate}
                                                        label={this.localize("Secondary Header")}
                                                        placeholder={this.localize("HEADER")}
                                                        onChange={this.onSecSenderTemplateChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input className="inputField"
                                                        value={this.state.secTitle}
                                                        label={this.localize("Secondary Title")}
                                                        placeholder={this.localize("PlaceHolderTitle")}
                                                        onChange={this.onSecTitleChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Input fluid className="inputField"
                                                            value={this.state.secImageLink}
                                                            label={this.localize("Secondary Image URL")}
                                                            placeholder={this.localize("ImageURL")}
                                                            onChange={this.onSecImageLinkChanged}
                                                            error={!(this.state.errorImageUrlMessage === "")}
                                                            autoComplete="off"
                                                        />
                                                        <input type="file" accept="image/"
                                                            style={{ display: 'none' }}
                                                            onChange={this.handleSecImageSelection}
                                                            ref={this.fileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={this.handleSecUploadClick}
                                                                style={{ width: '10px' }}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title="UploadImage"
                                                            />
                                                        </Flex.Item>
                                                    </Flex>
                                                    <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />
                                                    <div className="textArea">
                                                        <Text content={this.localize("Secondary Summary")} />
                                                        <Flex vAlign="center" hAlign="start">
                                                            <ToggleButtonGroup size="small" value={this.state.secBoldSummary}>
                                                                <ToggleButton value="Bold" onClick={this.onSecBoldSummaryChange}>
                                                                    <FormatBoldIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <ToggleButtonGroup size="small" value={this.state.secFontSummary} >
                                                                <ToggleButton value="Monospace" onClick={this.onSecFontSummaryChange}>
                                                                    <FontDownloadIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <Dropdown className="testDrop"
                                                                items={FontSize}
                                                                placeholder="FontSize"
                                                                onChange={this.onSecFontSizeChange}
                                                                value={this.state.secFontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onSecFontColorChange}
                                                                value={this.state.secFontColorSummary}
                                                            />
                                                            <Grid item>
                                                                <ToggleButtonGroup value={this.state.secAlignment} size="small" onChange={this.onSecAlignmentChange} exclusive>
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
                                                            value={this.state.secSummary}
                                                            onChange={this.onSecSummaryChanged}
                                                            fluid />
                                                    </div>

                                                    <Input className="inputField"
                                                        value={this.state.secAuthor}
                                                        label={this.localize("Secondary Author")}
                                                        placeholder={this.localize("Author")}
                                                        onChange={this.onSecAuthorChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle}
                                                        label={this.localize("Secondary Button Title")}
                                                        placeholder={this.localize("ButtonTitle")}
                                                        onChange={this.onSecBtnTitleChanged}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink}
                                                        label={this.localize("Secondary Button URL")}
                                                        placeholder={this.localize("ButtonURL")}
                                                        onChange={this.onSecBtnLinkChanged}
                                                        error={!(this.state.errorButtonUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />

                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurReaction}
                                                                onChange={this.onSecReactionChanged}
                                                            />
                                                        }
                                                        label="Reaction"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for reaction")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecReactionQuestionChanged}
                                                        value={this.state.secReactionQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurFreeText}
                                                                onChange={this.onSecFreeTextChanged}
                                                            />
                                                        }
                                                        label="Free Text"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for for free text")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecFreeTextQuestionChanged}
                                                        value={this.state.secFreeTextQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurYesNo}
                                                                onChange={this.onSecYesNoChanged}
                                                            />
                                                        }
                                                        label="Yes/ No Question"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for Yes / No")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecYesNoQuestionChanged}
                                                        value={this.state.secYesNoQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurLinkToSurvey}
                                                                onChange={this.onSecToggleLinktoSurveyChanged}
                                                            />
                                                        }
                                                        label="Link to Survey"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Survey URL")}
                                                        placeholder={this.localize("Enter URL")}
                                                        onChange={this.onSecLinktoSurveyChanged}
                                                        value={this.state.secLinkToSurvey}
                                                        error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />



                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}







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
                    } else if (this.state.template === "Announcement with video") {
                        if (this.state.language === "Primary") { 
                            return(
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")} />
                                                    <Dropdown
                                                        value={this.state.priLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        className="testDrop"
                                                    />
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
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Input fluid className="inputField"
                                                            value={this.state.imageLink}
                                                            label={this.localize("ImageURL")}
                                                            placeholder={this.localize("ImageURL")}
                                                            onChange={this.onImageLinkChanged}
                                                            error={!(this.state.errorImageUrlMessage === "")}
                                                            autoComplete="off"
                                                        />
                                                        <input type="file" accept="image/"
                                                            style={{ display: 'none' }}
                                                            onChange={this.handleImageSelection}
                                                            ref={this.fileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={this.handleUploadClick}
                                                                style={{ width: '10px' }}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title="UploadImage"
                                                            />
                                                        </Flex.Item>
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
                                                                value={this.state.fontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onFontColorChange}
                                                                value={this.state.fontColorSummary}
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
                                                    <Flex gap="gap.smaller" vAlign="center" className="inputField">
                                                        <Input onChange={this.onVideoUpload} 
                                                            type="file"
                                                            accept="video/mp4"
                                                        />
                                                        <Button
                                                            onClick={this.uploadVideo}
                                                            content={this.localize("Upload")}
                                                        />
                                                    </Flex>
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



                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}


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
                        } else if (this.state.language === "Secondary") {
                            return(
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")} />
                                                    <Dropdown
                                                        value={this.state.secLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        className="testDrop"
                                                    />
                                                    <Input className="senderTemplates"
                                                        value={this.state.secSenderTemplate}
                                                        label={this.localize("Secondary Header")}
                                                        placeholder={this.localize("HEADER")}
                                                        onChange={this.onSecSenderTemplateChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input className="inputField"
                                                        value={this.state.secTitle}
                                                        label={this.localize("Secondary TitleText")}
                                                        placeholder={this.localize("PlaceHolderTitle")}
                                                        onChange={this.onSecTitleChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Input fluid className="inputField"
                                                            value={this.state.secImageLink}
                                                            label={this.localize("Secondary ImageURL")}
                                                            placeholder={this.localize("ImageURL")}
                                                            onChange={this.onSecImageLinkChanged}
                                                            error={!(this.state.errorImageUrlMessage === "")}
                                                            autoComplete="off"
                                                        />
                                                        <input type="file" accept="image/"
                                                            style={{ display: 'none' }}
                                                            onChange={this.handleSecImageSelection}
                                                            ref={this.fileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={this.handleSecUploadClick}
                                                                style={{ width: '10px' }}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title="UploadImage"
                                                            />
                                                        </Flex.Item>
                                                    </Flex>
                                                    <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />

                                                    <div className="textArea">
                                                        <Text content={this.localize("Secondary Summary")} />
                                                        <Flex vAlign="center" hAlign="start">
                                                            <ToggleButtonGroup size="small" value={this.state.secBoldSummary}>
                                                                <ToggleButton value="Bold" onClick={this.onSecBoldSummaryChange}>
                                                                    <FormatBoldIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <ToggleButtonGroup size="small" value={this.state.secFontSummary} >
                                                                <ToggleButton value="Monospace" onClick={this.onSecFontSummaryChange}>
                                                                    <FontDownloadIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <Dropdown className="testDrop"
                                                                items={FontSize}
                                                                placeholder="FontSize"
                                                                onChange={this.onSecFontSizeChange}
                                                                value={this.state.secFontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onSecFontColorChange}
                                                                value={this.state.secFontColorSummary}
                                                            />
                                                            <Grid item>
                                                                <ToggleButtonGroup value={this.state.secAlignment} size="small" onChange={this.onSecAlignmentChange} exclusive>
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
                                                            value={this.state.secSummary}
                                                            onChange={this.onSecSummaryChanged}
                                                            fluid />
                                                    </div>

                                                    <Input className="inputField"
                                                        value={this.state.secAuthor}
                                                        label={this.localize("Secondary Author")}
                                                        placeholder={this.localize("Author")}
                                                        onChange={this.onSecAuthorChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input fluid className="inputField"
                                                        value={this.state.secVideoLink}
                                                        label={this.localize("Secondary Video URL")}
                                                        placeholder={this.localize("Video URL")}
                                                        onChange={this.onSecVideoLinkChanged}
                                                        error={!(this.state.errorVideoUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Flex gap="gap.smaller" vAlign="center" className="inputField">
                                                        <Input onChange={this.onVideoUpload} 
                                                            type="file"
                                                            accept="video/mp4"
                                                        />
                                                        <Button
                                                            onClick={this.uploadVideo}
                                                            content={this.localize("Upload")}
                                                        />
                                                    </Flex>
                                                    <Text className={(this.state.errorVideoUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorVideoUrlMessage} />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle}
                                                        label={this.localize("Secondary ButtonTitle")}
                                                        placeholder={this.localize("ButtonTitle")}
                                                        onChange={this.onSecBtnTitleChanged}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink}
                                                        label={this.localize("Secondary ButtonURL")}
                                                        placeholder={this.localize("ButtonURL")}
                                                        onChange={this.onSecBtnLinkChanged}
                                                        error={!(this.state.errorButtonUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurReaction}
                                                                onChange={this.onSecReactionChanged}
                                                            />
                                                        }
                                                        label="Reaction"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for reaction")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecReactionQuestionChanged}
                                                        value={this.state.secReactionQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurFreeText}
                                                                onChange={this.onSecFreeTextChanged}
                                                            />
                                                        }
                                                        label="Free Text"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for for free text")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecFreeTextQuestionChanged}
                                                        value={this.state.secFreeTextQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurYesNo}
                                                                onChange={this.onSecYesNoChanged}
                                                            />
                                                        }
                                                        label="Yes/ No Question"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for Yes / No")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecYesNoQuestionChanged}
                                                        value={this.state.secYesNoQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurLinkToSurvey}
                                                                onChange={this.onSecToggleLinktoSurveyChanged}
                                                            />
                                                        }
                                                        label="Link to Survey"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Survey URL")}
                                                        placeholder={this.localize("Enter URL")}
                                                        onChange={this.onSecLinktoSurveyChanged}
                                                        value={this.state.secLinkToSurvey}
                                                        error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />



                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}



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
                    } else if (this.state.template === "Announcement in plain text") {
                        if (this.state.language === "Primary") {
                            return(
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")} />
                                                    <Dropdown
                                                        value={this.state.priLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        className="testDrop" 
                                                    />
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
                                                                value={this.state.fontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onFontColorChange}
                                                                value={this.state.fontColorSummary}
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



                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}



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
                        } else if (this.state.language === "Secondary") {
                            return(
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")} />
                                                    <Dropdown
                                                        value={this.state.secLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        className="testDrop" 
                                                    />
                                                    <Input className="senderTemplates"
                                                        value={this.state.secSenderTemplate}
                                                        label={this.localize("Secondary Header")}
                                                        placeholder={this.localize("HEADER")}
                                                        onChange={this.onSecSenderTemplateChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input className="inputField"
                                                        value={this.state.secTitle}
                                                        label={this.localize("Secondary TitleText")}
                                                        placeholder={this.localize("PlaceHolderTitle")}
                                                        onChange={this.onSecTitleChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <div className="textArea">
                                                        <Text content={this.localize("Secondary Summary")} />
                                                        <Flex vAlign="center" hAlign="start">
                                                            <ToggleButtonGroup size="small" value={this.state.secBoldSummary}>
                                                                <ToggleButton value="Bold" onClick={this.onSecBoldSummaryChange}>
                                                                    <FormatBoldIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <ToggleButtonGroup size="small" value={this.state.secFontSummary} >
                                                                <ToggleButton value="Monospace" onClick={this.onSecFontSummaryChange}>
                                                                    <FontDownloadIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <Dropdown className="testDrop"
                                                                items={FontSize}
                                                                placeholder="FontSize"
                                                                onChange={this.onSecFontSizeChange}
                                                                value={this.state.secFontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onSecFontColorChange}
                                                                value={this.state.secFontColorSummary}
                                                            />
                                                            <Grid item>
                                                                <ToggleButtonGroup value={this.state.secAlignment} size="small" onChange={this.onSecAlignmentChange} exclusive>
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
                                                            placeholder={this.localize("Secondary Summary")}
                                                            value={this.state.secSummary}
                                                            onChange={this.onSecSummaryChanged}
                                                            fluid />
                                                    </div>

                                                    <Input className="inputField"
                                                        value={this.state.secAuthor}
                                                        label={this.localize("Secondary Author")}
                                                        placeholder={this.localize("Author")}
                                                        onChange={this.onSecAuthorChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle}
                                                        label={this.localize("Secondary ButtonTitle")}
                                                        placeholder={this.localize("ButtonTitle")}
                                                        onChange={this.onSecBtnTitleChanged}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink}
                                                        label={this.localize("Secondary ButtonURL")}
                                                        placeholder={this.localize("ButtonURL")}
                                                        onChange={this.onSecBtnLinkChanged}
                                                        error={!(this.state.errorButtonUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurReaction}
                                                                onChange={this.onSecReactionChanged}
                                                            />
                                                        }
                                                        label="Reaction"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for reaction")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecReactionQuestionChanged}
                                                        value={this.state.secReactionQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurFreeText}
                                                                onChange={this.onSecFreeTextChanged}
                                                            />
                                                        }
                                                        label="Free Text"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for for free text")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecFreeTextQuestionChanged}
                                                        value={this.state.secFreeTextQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurYesNo}
                                                                onChange={this.onSecYesNoChanged}
                                                            />
                                                        }
                                                        label="Yes/ No Question"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for Yes / No")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecYesNoQuestionChanged}
                                                        value={this.state.secYesNoQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurLinkToSurvey}
                                                                onChange={this.onSecToggleLinktoSurveyChanged}
                                                            />
                                                        }
                                                        label="Secondary Link to Survey"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Survey URL")}
                                                        placeholder={this.localize("Enter URL")}
                                                        onChange={this.onSecLinktoSurveyChanged}
                                                        value={this.state.secLinkToSurvey}
                                                        error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />


                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}



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
                    } else if (this.state.template === "Event promotion with sessions to register/join") {
                        if (this.state.language === "Primary") {
                            return (
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")} />
                                                    <Dropdown
                                                        value={this.state.priLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        className="testDrop" 
                                                    />
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
                                                    {/*<Input fluid className="inputField"*/}
                                                    {/*    value={this.state.imageLink}*/}
                                                    {/*    label={this.localize("ImageURL")}*/}
                                                    {/*    placeholder={this.localize("ImageURL")}*/}
                                                    {/*    onChange={this.onImageLinkChanged}*/}
                                                    {/*    error={!(this.state.errorImageUrlMessage === "")}*/}
                                                    {/*    autoComplete="off"*/}
                                                    {/*/>*/}
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Input fluid className="inputField"
                                                            value={this.state.imageLink}
                                                            label={this.localize("ImageURL")}
                                                            placeholder={this.localize("ImageURL")}
                                                            onChange={this.onImageLinkChanged}
                                                            error={!(this.state.errorImageUrlMessage === "")}
                                                            autoComplete="off"
                                                        />
                                                        <input type="file" accept="image/"
                                                            style={{ display: 'none' }}
                                                            onChange={this.handleImageSelection}
                                                            ref={this.fileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={this.handleUploadClick}
                                                                style={{ width: '10px' }}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title="UploadImage"
                                                            />
                                                        </Flex.Item>
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
                                                                value={this.state.fontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onFontColorChange}
                                                                value={this.state.fontColorSummary}
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



                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}



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
                        } else if (this.state.language === "Secondary") {
                            return (
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Primary")}
                                                            value="Primary"
                                                        />
                                                        <Button
                                                            onClick={this.switchLanguage}
                                                            content={this.localize("Secondary")}
                                                            value="Secondary"
                                                        />
                                                    </Flex>
                                                    <Text content={this.localize("Language")}/>
                                                    <Dropdown
                                                        value={this.state.secLanguage}
                                                        items={language}
                                                        onChange={this.selectLanguage}
                                                        placeholder={this.state.secLanguage}
                                                        className="testDrop" 
                                                    />
                                                    <Input className="senderTemplates"
                                                        value={this.state.secSenderTemplate}
                                                        label={this.localize("Secondary Header")}
                                                        placeholder={this.localize("HEADER")}
                                                        onChange={this.onSecSenderTemplateChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input className="inputField"
                                                        value={this.state.secTitle}
                                                        label={this.localize("Secondary TitleText")}
                                                        placeholder={this.localize("PlaceHolderTitle")}
                                                        onChange={this.onSecTitleChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                        <Input fluid className="inputField"
                                                            value={this.state.secImageLink}
                                                            label={this.localize("Secondary ImageURL")}
                                                            placeholder={this.localize("ImageURL")}
                                                            onChange={this.onSecImageLinkChanged}
                                                            error={!(this.state.errorImageUrlMessage === "")}
                                                            autoComplete="off"
                                                        />
                                                        <input type="file" accept="image/"
                                                            style={{ display: 'none' }}
                                                            onChange={this.handleSecImageSelection}
                                                            ref={this.fileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={this.handleSecUploadClick}
                                                                style={{ width: '10px' }}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title="UploadImage"
                                                            />
                                                        </Flex.Item>
                                                    </Flex>
                                                    <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />
    
                                                    <div className="textArea">
                                                        <Text content={this.localize("Secondary Summary")} />
                                                        <Flex vAlign="center" hAlign="start">
                                                            <ToggleButtonGroup size="small" value={this.state.secBoldSummary}>
                                                                <ToggleButton value="Bold" onClick={this.onSecBoldSummaryChange}>
                                                                    <FormatBoldIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <ToggleButtonGroup size="small" value={this.state.secFontSummary} >
                                                                <ToggleButton value="Monospace" onClick={this.onSecFontSummaryChange}>
                                                                    <FontDownloadIcon />
                                                                </ToggleButton>
                                                            </ToggleButtonGroup>
                                                            <Dropdown className="testDrop"
                                                                items={FontSize}
                                                                placeholder="FontSize"
                                                                onChange={this.onSecFontSizeChange}
                                                                value={this.state.secFontSizeSummary}
                                                            />
                                                            <Dropdown className="testDrop"
                                                                items={Color}
                                                                placeholder="Color"
                                                                onChange={this.onSecFontColorChange}
                                                                value={this.state.secFontColorSummary}
                                                            />
                                                            <Grid item>
                                                                <ToggleButtonGroup value={this.state.secAlignment} size="small" onChange={this.onSecAlignmentChange} exclusive>
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
                                                            value={this.state.secSummary}
                                                            onChange={this.onSecSummaryChanged}
                                                            fluid />
                                                    </div>
    
                                                    <Input className="inputField"
                                                        value={this.state.secAuthor}
                                                        label={this.localize("Secondary Author")}
                                                        placeholder={this.localize("Author")}
                                                        onChange={this.onSecAuthorChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle}
                                                        label={this.localize("Secondary ButtonTitle")}
                                                        placeholder={this.localize("ButtonTitle")}
                                                        onChange={this.onSecBtnTitleChanged}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink}
                                                        label={this.localize("Secondary ButtonURL")}
                                                        placeholder={this.localize("ButtonURL")}
                                                        onChange={this.onSecBtnLinkChanged}
                                                        error={!(this.state.errorButtonUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
    
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle2}
                                                        label={this.localize("Secondary Button title 2")}
                                                        placeholder={this.localize("Button title 2")}
                                                        onChange={this.onSecBtnTitleChanged2}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink2}
                                                        label={this.localize("Secondary Secondary Button URL 2")}
                                                        placeholder={this.localize("Button URL 2")}
                                                        onChange={this.onSecBtnLinkChanged2}
                                                        error={!(this.state.errorButtonUrlMessage2 === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage2 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage2} />
    
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle3}
                                                        label={this.localize("Secondary Button title 3")}
                                                        placeholder={this.localize("Button title 3")}
                                                        onChange={this.onSecBtnTitleChanged3}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink3}
                                                        label={this.localize("Secondary Button URL 3")}
                                                        placeholder={this.localize("Button URL 3")}
                                                        onChange={this.onSecBtnLinkChanged3}
                                                        error={!(this.state.errorButtonUrlMessage3 === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage3 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage3} />
    
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle4}
                                                        label={this.localize("Secondary Button title 4")}
                                                        placeholder={this.localize("Button title 4")}
                                                        onChange={this.onSecBtnTitleChanged4}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink4}
                                                        label={this.localize("Secondary Button URL 4")}
                                                        placeholder={this.localize("Button URL 4")}
                                                        onChange={this.onSecBtnLinkChanged4}
                                                        error={!(this.state.errorButtonUrlMessage4 === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage4 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage4} />
    
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnTitle5}
                                                        label={this.localize("Secondary Button title 5")}
                                                        placeholder={this.localize("Button title 5")}
                                                        onChange={this.onSecBtnTitleChanged5}
                                                        autoComplete="off"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        value={this.state.secbtnLink5}
                                                        label={this.localize("Secondary Button URL 5")}
                                                        placeholder={this.localize("Button URL 5")}
                                                        onChange={this.onSecBtnLinkChanged5}
                                                        error={!(this.state.errorButtonUrlMessage5 === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorButtonUrlMessage5 === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage5} />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurReaction}
                                                                onChange={this.onSecReactionChanged}
                                                            />
                                                        }
                                                        label="Reaction"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for reaction")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecReactionQuestionChanged}
                                                        value={this.state.secReactionQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurFreeText}
                                                                onChange={this.onSecFreeTextChanged}
                                                            />
                                                        }
                                                        label="Free Text"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for for free text")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecFreeTextQuestionChanged}
                                                        value={this.state.secFreeTextQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurYesNo}
                                                                onChange={this.onSecYesNoChanged}
                                                            />
                                                        }
                                                        label="Yes/ No Question"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Question for Yes / No")}
                                                        placeholder={this.localize("Enter the question")}
                                                        onChange={this.onSecYesNoQuestionChanged}
                                                        value={this.state.secYesNoQuestion}
                                                        autoComplete="off"
                                                    />
                                                    <FormControlLabel
                                                        control={
                                                            <Switch
                                                                color="primary"
                                                                checked={this.state.secSurLinkToSurvey}
                                                                onChange={this.onSecToggleLinktoSurveyChanged}
                                                            />
                                                        }
                                                        label="Secondary Link to Survey"
                                                    />
                                                    <Input className="inputField"
                                                        fluid
                                                        label={this.localize("Secondary Survey URL")}
                                                        placeholder={this.localize("Enter URL")}
                                                        onChange={this.onSecLinktoSurveyChanged}
                                                        value={this.state.secLinkToSurvey}
                                                        error={!(this.state.errorLinkToSurveyUrlMessage === "")}
                                                        autoComplete="off"
                                                    />
                                                    <Text className={(this.state.errorLinkToSurveyUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorLinkToSurveyUrlMessage} />



                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="ScheduleCheckBox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onScheduleSelected}*/}
                                                    {/*        label={this.localize("ScheduledSend")}*/}
                                                    {/*        checked={this.state.selectedSchedule}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.selectedSchedule}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.selectedSchedule}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.scheduledDate)}*/}
                                                    {/*            onChange={this.handleMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}
                                                    {/*<Flex hAlign="start">*/}
                                                    {/*    <h3><Checkbox*/}
                                                    {/*        className="IsExpirySetCheckbox"*/}
                                                    {/*        labelPosition="start"*/}
                                                    {/*        onClick={this.onIsExpirySetSelected}*/}
                                                    {/*        label={this.localize("ExpirySend")}*/}
                                                    {/*        checked={this.state.isExpirySet}*/}
                                                    {/*        toggle*/}
                                                    {/*    /></h3>*/}
                                                    {/*</Flex>*/}
                                                    {/*<Flex gap="gap.smaller" className="DateTimeSelector">*/}
                                                    {/*    <Datepicker*/}
                                                    {/*        disabled={!this.state.isExpirySet}*/}
                                                    {/*        defaultSelectedDate={this.getDateObject(this.state.expiryDate)}*/}
                                                    {/*        minDate={new Date()}*/}
                                                    {/*        inputOnly*/}
                                                    {/*        onDateChange={this.handleExpiryDateChange}*/}
                                                    {/*    />*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="hour"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={hours}*/}
                                                    {/*            defaultValue={this.getDateHour(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryHourChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*    <Flex.Item shrink={true} size="1%">*/}
                                                    {/*        <Dropdown*/}
                                                    {/*            placeholder="mins"*/}
                                                    {/*            disabled={!this.state.isExpirySet}*/}
                                                    {/*            fluid={true}*/}
                                                    {/*            items={minutes}*/}
                                                    {/*            defaultValue={this.getDateMins(this.state.expiryDate)}*/}
                                                    {/*            onChange={this.handleExpiryMinsChange}*/}
                                                    {/*        />*/}
                                                    {/*    </Flex.Item>*/}
                                                    {/*</Flex>*/}
                                                    {/*<div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>*/}
                                                    {/*    <div className="noteText">*/}
                                                    {/*        <Text error content={this.localize('FutureDateError')} />*/}
                                                    {/*    </div>*/}
                                                    {/*</div>*/}



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
                                                <Button content={this.localize("Test")} onClick={this.printCon} />
                                            </Flex>
                                        </Flex>
    
                                    </Flex>
                                </div>
                            );
                        }
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
                                                            checked={this.state.emailOption}
                                                            value={this.state.emailOption}
                                                            onChange={this.emailOptionOnChanged}
                                                            color="primary"
                                                        />
                                                    }
                                                    label="Send email to members"
                                                />
                                            {/*</div>*/}
                                            <FormControlLabel
                                                control={
                                                    <Switch
                                                        color="primary"
                                                        checked={this.state.selectedSchedule}
                                                        onChange={this.onScheduleSelected}
                                                    />
                                                }
                                                label={this.localize("ScheduledSend")}
                                            />
                                            <Flex gap="gap.smaller" className="DateTimeSelector">
                                                <Datepicker
                                                    disabled={!this.state.selectedSchedule}
                                                    defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}
                                                    minDate={new Date()}
                                                    inputOnly
                                                    onDateChange={this.handleDateChange}
                                                />
                                                <Flex.Item shrink={true} size="1%">
                                                    <Dropdown
                                                        placeholder="hour"
                                                        disabled={!this.state.selectedSchedule}
                                                        fluid={true}
                                                        items={hours}
                                                        defaultValue={this.getDateHour(this.state.scheduledDate)}
                                                        onChange={this.handleHourChange}
                                                    />
                                                </Flex.Item>
                                                <Flex.Item shrink={true} size="1%">
                                                    <Dropdown
                                                        placeholder="mins"
                                                        disabled={!this.state.selectedSchedule}
                                                        fluid={true}
                                                        items={minutes}
                                                        defaultValue={this.getDateMins(this.state.scheduledDate)}
                                                        onChange={this.handleMinsChange}
                                                    />
                                                </Flex.Item>
                                            </Flex>
                                            <div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>
                                                <div className="noteText">
                                                    <Text error content={this.localize('FutureDateError')} />
                                                </div>
                                            </div>
                                            <FormControlLabel
                                                control={
                                                    <Switch
                                                        color="primary"
                                                        checked={this.state.isExpirySet}
                                                        onChange={this.onIsExpirySetSelected}
                                                    />
                                                }
                                                label={this.localize("ExpirySend")}
                                            />
                                            <Flex gap="gap.smaller" className="DateTimeSelector">
                                                <Datepicker
                                                    disabled={!this.state.isExpirySet}
                                                    defaultSelectedDate={this.getDateObject(this.state.expiryDate)}
                                                    minDate={new Date()}
                                                    inputOnly
                                                    onDateChange={this.handleExpiryDateChange}
                                                />
                                                <Flex.Item shrink={true} size="1%">
                                                    <Dropdown
                                                        placeholder="hour"
                                                        disabled={!this.state.isExpirySet}
                                                        fluid={true}
                                                        items={hours}
                                                        defaultValue={this.getDateHour(this.state.expiryDate)}
                                                        onChange={this.handleExpiryHourChange}
                                                    />
                                                </Flex.Item>
                                                <Flex.Item shrink={true} size="1%">
                                                    <Dropdown
                                                        placeholder="mins"
                                                        disabled={!this.state.isExpirySet}
                                                        fluid={true}
                                                        items={minutes}
                                                        defaultValue={this.getDateMins(this.state.expiryDate)}
                                                        onChange={this.handleExpiryMinsChange}
                                                    />
                                                </Flex.Item>
                                            </Flex>
                                            <div className={this.state.expiryfuturedate && this.state.isExpirySet ? "ErrorMessage" : "hide"}>
                                                <div className="noteText">
                                                    <Text error content={this.localize('FutureDateError')} />
                                                </div>
                                            </div>
                                            
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


    //get the next rounded up (ceil) date in minutes
    private getRoundedDate = (minutes: number, d = new Date()) => {

        let ms = 1000 * 60 * minutes; // convert minutes to ms
        let roundedDate = new Date(Math.ceil(d.getTime() / ms) * ms);

        return roundedDate
    }

    //get date object based on the string parameter
    private getDateObject = (datestring?: string) => {
        if (!datestring) {
            var TempDate = new Date(); //get current date
            TempDate.setTime(TempDate.getTime() + 86400000);
            return TempDate; //if date string is not provided, then return tomorrow rounded up next 5 minutes
        }
        return new Date(datestring); //if date string is provided, return current date object
    }

    //get the hour of the datestring
    private getDateHour = (datestring: string) => {
        if (!datestring) return "00";
        var thour = new Date(datestring).getHours().toString();
        return thour.padStart(2, "0");
    }

    //get the mins of the datestring
    private getDateMins = (datestring: string) => {
        if (!datestring) return "00";
        var tmins = new Date(datestring).getMinutes().toString();
        return tmins.padStart(2, "0");
    }

    //handles click on DatePicker to change the schedule date
    private handleDateChange = (e: any, v: any) => {
        var TempDate = v.value; //set the tempdate var with the value selected by the user
        TempDate.setMinutes(parseInt(this.state.DMYMins)); //set the minutes selected on minutes drop down 
        TempDate.setHours(parseInt(this.state.DMYHour)); //set the hour selected on hour drop down
        //set the state variables
        this.setState({
            scheduledDate: TempDate.toUTCString(), //updates the state string representation
            DMY: TempDate, //updates the date on the state
        });
    }

    //handles selection on the hour combo
    private handleHourChange = (e: any, v: any) => {
        var TempDate = this.state.DMY; //get the tempdate from the state
        TempDate.setHours(parseInt(v.value)); //set hour with the value select on the hour drop down
        //set state variables
        this.setState({
            scheduledDate: TempDate.toUTCString(), //updates the string representation 
            DMY: TempDate, //updates DMY
            DMYHour: v.value, //set the new hour value on the state
        });
    }

    //handles selection on the minutes combo
    private handleMinsChange = (e: any, v: any) => {
        var TempDate = this.state.DMY; //get the tempdate from the state
        TempDate.setMinutes(parseInt(v.value)); //set minutes with the value select on the minutes drop down
        //set state variables
        this.setState({
            scheduledDate: TempDate.toUTCString(), //updates the string representation 
            DMY: TempDate, //updates DMY
            DMYMins: v.value, //set the bew minutes on the state
        });
    }

    //handler for the Schedule Send checkbox
    private onScheduleSelected = () => {
        var TempDate = this.getRoundedDate(5, this.getDateObject()); //get the next day date rounded to the nearest hour/minute
        //set the state
        this.setState({
            selectedSchedule: !this.state.selectedSchedule,
            scheduledDate: TempDate.toUTCString(),
            DMY: TempDate
        });
    }

    //handles click on DatePicker to change the expiry date
    private handleExpiryDateChange = (e: any, v: any) => {
        var TempDate = v.value; //set the tempdate var with the value selected by the user
        TempDate.setMinutes(parseInt(this.state.expiryDMYMins)); //set the minutes selected on minutes drop down 
        TempDate.setHours(parseInt(this.state.expiryDMYHour)); //set the hour selected on hour drop down
        //set the state variables
        this.setState({
            expiryDate: TempDate.toUTCString(), //updates the state string representation
            expiryDMY: TempDate, //updates the date on the state
        });
    }

    //handles selection on the hour combo
    private handleExpiryHourChange = (e: any, v: any) => {
        var TempDate = this.state.expiryDMY; //get the tempdate from the state
        TempDate.setHours(parseInt(v.value)); //set hour with the value select on the hour drop down
        //set state variables
        this.setState({
            expiryDate: TempDate.toUTCString(), //updates the string representation 
            expiryDMY: TempDate, //updates expiryDMY
            expiryDMYHour: v.value, //set the new hour value on the state
        });
    }

    //handles selection on the minutes combo
    private handleExpiryMinsChange = (e: any, v: any) => {
        var TempDate = this.state.expiryDMY; //get the tempdate from the state
        TempDate.setMinutes(parseInt(v.value)); //set minutes with the value select on the minutes drop down
        //set state variables
        this.setState({
            expiryDate: TempDate.toUTCString(), //updates the string representation 
            expiryDMY: TempDate, //updates expiryDMY
            expiryDMYMins: v.value, //set the bew minutes on the state
        });
    }

    //handler for the Schedule Send checkbox
    private onIsExpirySetSelected = () => {
        var TempDate = this.getRoundedDate(5, this.getDateObject()); //get the next day date rounded to the nearest hour/minute
        //set the state
        this.setState({
            isExpirySet: !this.state.isExpirySet,
            expiryDate: TempDate.toUTCString(),
            expiryDMY: TempDate
        });
    }



    //called when the user clicks to schedule the message
    private onSchedule = () => {
        var Today = new Date(); //today date
        var Scheduled = new Date(this.state.DMY); //scheduled date

        var TempIsExpirySet = this.state.isExpirySet;
        var TempExpiryDMY = new Date(this.state.expiryDMY); //scheduled date

        //only allow the save when the scheduled date is 30 mins in the future, if that is the case calls the onSave function
        if (Scheduled.getTime() > Today.getTime() + 1800000) {
            if (TempIsExpirySet && TempExpiryDMY <= Scheduled) {
                this.setState({
                    expiryfuturedate: true
                })
            }
            else {
                this.onSave()
            }
        }
        else {
            //set the state to indicate future date error
            //if futuredate is true, an error message is shown right below the date selector
            this.setState({
                futuredate: true
            })
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
            secSenderTemplate: this.state.secSenderTemplate,
            priLanguage: this.state.priLanguage,
            secLanguage: this.state.secLanguage,
            template: this.state.template,
            id: this.state.messageId,
            title: this.state.title,
            secTitle: this.state.secTitle,
            imageLink: this.state.imageLink,
            secImageLink: this.state.secImageLink,
            videoLink: this.state.videoLink,
            secVideoLink: this.state.secVideoLink,
            summary: this.state.summary,
            secSummary: this.state.secSummary,
            alignment: this.state.alignment,
            secAlignment: this.state.secAlignment,
            boldSummary: this.state.boldSummary,
            secBoldSummary: this.state.secBoldSummary,
            fontSummary: this.state.fontSummary,
            secFontSummary: this.state.secFontSummary,
            fontSizeSummary: this.state.fontSizeSummary,
            secFontSizeSummary: this.state.secFontSizeSummary,
            fontColorSummary: this.state.fontColorSummary,
            secFontColorSummary: this.state.secFontColorSummary,
            author: this.state.author,
            secAuthor: this.state.secAuthor,
            buttonTitle: this.state.btnTitle,
            secButtonTitle: this.state.secbtnTitle,
            buttonLink: this.state.btnLink,
            secButtonLink: this.state.secbtnLink,
            buttonTitle2: this.state.btnTitle2,
            secButtonTitle2: this.state.secbtnTitle2,
            buttonLink2: this.state.btnLink2,
            secButtonLink2: this.state.secbtnLink2,
            buttonTitle3: this.state.btnTitle3,
            secButtonTitle3: this.state.secbtnTitle3,
            buttonLink3: this.state.btnLink3,
            secButtonLink3: this.state.secbtnLink3,
            buttonTitle4: this.state.btnTitle4,
            secButtonTitle4: this.state.secbtnTitle4,
            buttonLink4: this.state.btnLink4,
            secButtonLink4: this.state.secbtnLink4,
            buttonTitle5: this.state.btnTitle5,
            secButtonTitle5: this.state.secbtnTitle5,
            buttonLink5: this.state.btnLink5,
            secButtonLink5: this.state.secbtnLink5,
            surReaction: this.state.surReaction,
            secSurReaction: this.state.secSurReaction,
            reactionQuestion: this.state.reactionQuestion,
            secReactionQuestion: this.state.secReactionQuestion,
            surFreeText: this.state.surFreeText,
            secSurFreeText: this.state.secSurFreeText,
            freeTextQuestion: this.state.freeTextQuestion,
            secFreeTextQuestion: this.state.secFreeTextQuestion,
            surYesNo: this.state.surYesNo,
            secSurYesNo: this.state.secSurYesNo,
            yesNoQuestion: this.state.yesNoQuestion,
            secYesNoQuestion: this.state.secYesNoQuestion,
            surLinkToSurvey: this.state.surLinkToSurvey,
            secSurLinkToSurvey: this.state.secSurLinkToSurvey,
            linkToSurvey: this.state.linkToSurvey,
            secLinkToSurvey: this.state.secLinkToSurvey,
            teams: selectedTeams,
            rosters: selctedRosters,
            groups: selectedGroups,
            isScheduled: this.state.selectedSchedule,
            ScheduledDate: new Date(this.state.scheduledDate),
            isExpirySet: this.state.isExpirySet,
            expiryDate: new Date(this.state.expiryDate),
            isExpiredContentErased: false,
            uploadedList: this.state.uploadedList,
            uploadedListName: this.state.uploadedListName,
            emailOption: this.state.emailOption,
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
                video: null,
                alignment: "left",
                page: "TemplateSelection",
                language: "Primary",
                priLanguage: "English",
                secLanguage: "",
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
                secTitle: "",
                secImageLink: "",
                secVideoLink: "",
                secAlignment: "left",
                secSummary: "",
                secBoldSummary: "",
                secFontSummary: "",
                secFontSizeSummary: "",
                secFontColorSummary: "",
                secAuthor: "",
                secbtnTitle: "",
                secbtnLink: "",
                secbtnTitle2: "",
                secbtnLink2: "",
                secbtnTitle3: "",
                secbtnLink3: "",
                secbtnTitle4: "",
                secbtnLink4: "",
                secbtnTitle5: "",
                secbtnLink5: "",
                secSurReaction: false,
                secReactionQuestion: "",
                secSurFreeText: false,
                secFreeTextQuestion: "",
                secSurYesNo: false,
                secYesNoQuestion: "",
                secSurLinkToSurvey: false,
                secLinkToSurvey: "",
                secSenderTemplate: "",
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
        if (this.state.btnTitle5 && event.target.value) {
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
        if (event.target.files[0]) {
            this.setState({
                selectedFile: event.target.files[0],
                selectedFileName: event.target.files[0].name
            });
        }
        
        //console.log(event.target.files[0]);
    }

    private handleOnDrop = (data: any, file: any) => {
        this.setState({
            uploadedList: []
        });
        for (let i = 0; i < data.length - 1; i++) {
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
    }

    private handleOnError = (err: any , file: any, inputElem:any , reason:any) => {
        console.log(err)
    }

    private handleOnRemoveFile = (data: any) => {
        //console.log('---------------------------')
        //console.log(data)
        //console.log('---------------------------')
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

    private emailOptionOnChanged = () => {
        this.setState({
            emailOption: !this.state.emailOption
        })
    }

    //For Secondary Language Handle
    private onSecSenderTemplateChanged = (event: any, itemsData: any) => {
        this.setState({
            secSenderTemplate: itemsData.value
        });
    }

    private onSecTitleChanged = (event: any) => {
        this.setState({
            secTitle: event.target.value
        });
    }

    private onSecImageLinkChanged = (event: any) => {
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
        this.setState({
            secImageLink: event.target.value
        });
    }

    private onSecVideoLinkChanged = (event: any) => {
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
        this.setState({
            secVideoLink: event.target.value
        });
    }

    private onSecBoldSummaryChange = (event: any) => {
        if (this.state.secBoldSummary === "Bold") {
            this.setState({
                secBoldSummary: ""
            });
        } else {
            this.setState({
                secBoldSummary: "Bold"
            });
        }

    }

    private onSecFontSummaryChange = (event: any) => {
        if (this.state.secFontSummary === "Monospace") {
            this.setState({
                secFontSummary: ""
            });
        } else {
            this.setState({
                secFontSummary: "Monospace"
            });
        }
    }

    private onSecFontSizeChange = (event: any, itemsData: any) => {
        if (itemsData.value === "Default") {
            this.setState({
                secFontSizeSummary: "default"
            });
        } else if (itemsData.value === "Small") {
            this.setState({
                secFontSizeSummary: "small"
            });
        } else if (itemsData.value === "Medium") {
            this.setState({
                secFontSizeSummary: "medium"
            });
        } else if (itemsData.value === "Large") {
            this.setState({
                secFontSizeSummary: "large"
            });
        } else if (itemsData.value === "ExtraLarge") {
            this.setState({
                secFontSizeSummary: "extraLarge"
            });
        }
    }

    private onSecFontColorChange = (event: any, itemsData: any) => {
        if (itemsData.value === "Default") {
            this.setState({
                secFontColorSummary: "default"
            });
        } else if (itemsData.value === "Accent") {
            this.setState({
                secFontColorSummary: "accent"
            });
        } else if (itemsData.value === "Good") {
            this.setState({
                secFontColorSummary: "good"
            });
        } else if (itemsData.value === "Warning") {
            this.setState({
                secFontColorSummary: "warning"
            });
        } else if (itemsData.value === "Attention") {
            this.setState({
                secFontColorSummary: "attention"
            });
        } else if (itemsData.value === "Light") {
            this.setState({
                secFontColorSummary: "light"
            });
        } else if (itemsData.value === "Dark") {
            this.setState({
                secFontColorSummary: "dark"
            });
        }
    }

    private onSecAlignmentChange = (event: any, onAlignmentChange: any) => {
        if (onAlignmentChange === "left") {
            this.setState({
                secAlignment: "left"
            });
        } else if (onAlignmentChange === "center") {
            this.setState({
                secAlignment: "center"
            });
        } else if (onAlignmentChange === "right") {
            this.setState({
                secAlignment: "right"
            });
        }
    }

    private onSecSummaryChanged = (event: any) => {
        this.setState({
            secSummary: event.target.value
        });
    }

    private onSecAuthorChanged = (event: any) => {
        this.setState({
            secAuthor: event.target.value
        });
    }

    private onSecBtnTitleChanged = (event: any) => {
        if (event.target.value && this.state.secbtnLink) {
            this.setState({
                secbtnTitle: event.target.value
            });
        } else {
            this.setState({
                secbtnTitle: event.target.value
            });
        }
    }

    private onSecBtnTitleChanged2 = (event: any) => {
        if (event.target.value && this.state.secbtnLink2) {
            this.setState({
                secbtnTitle2: event.target.value
            });
        } else {
            this.setState({
                secbtnTitle2: event.target.value
            });
        }
    }

    private onSecBtnTitleChanged3 = (event: any) => {
        if (event.target.value && this.state.secbtnLink3) {
            this.setState({
                secbtnTitle3: event.target.value
            });
        } else {
            this.setState({
                secbtnTitle3: event.target.value
            });
        }
    }

    private onSecBtnTitleChanged4 = (event: any) => {
        if (event.target.value && this.state.secbtnLink4) {
            this.setState({
                secbtnTitle4: event.target.value
            });
        } else {
            this.setState({
                secbtnTitle4: event.target.value
            });
        }
    }

    private onSecBtnTitleChanged5 = (event: any) => {
        if (event.target.value && this.state.secbtnLink5) {
            this.setState({
                secbtnTitle5: event.target.value
            });
        } else {
            this.setState({
                secbtnTitle5: event.target.value
            });
        }
    }

    private onSecBtnLinkChanged = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
        }

        if (this.state.secbtnTitle && event.target.value) {
            this.setState({
                secbtnLink: event.target.value
            });
        } else {
            this.setState({
                secbtnLink: event.target.value
            });
        }
    }

    private onSecBtnLinkChanged2 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage2: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage2: ""
            });
        }

        if (this.state.secbtnTitle2 && event.target.value) {
            this.setState({
                secbtnLink2: event.target.value
            });
        } else {
            this.setState({
                secbtnLink2: event.target.value
            });
        }
    }

    private onSecBtnLinkChanged3 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage3: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage3: ""
            });
        }

        if (this.state.secbtnTitle3 && event.target.value) {
            this.setState({
                secbtnLink3: event.target.value
            });
        } else {
            this.setState({
                secbtnLink3: event.target.value
            });
        }
    }

    private onSecBtnLinkChanged4 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage4: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage4: ""
            });
        }

        if (this.state.secbtnTitle4 && event.target.value) {
            this.setState({
                secbtnLink4: event.target.value
            });
        } else {
            this.setState({
                secbtnLink4: event.target.value
            });
        }
    }

    private onSecBtnLinkChanged5 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage5: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage5: ""
            });
        }

        if (this.state.secbtnTitle5 && event.target.value) {
            this.setState({
                secbtnLink5: event.target.value
            });
        } else {
            this.setState({
                secbtnLink5: event.target.value
            });
        }
    }

    private onSecReactionChanged = (event: any) => {
        this.setState({
            secSurReaction: !this.state.secSurReaction
        });
    }

    private onSecReactionQuestionChanged = (event: any) => {
        this.setState({
            secReactionQuestion: event.target.value
        });
    }

    private onSecFreeTextChanged = (event: any) => {
        this.setState({
            secSurFreeText: !this.state.secSurFreeText
        });
    }

    private onSecFreeTextQuestionChanged = (event: any) => {
        this.setState({
            secFreeTextQuestion: event.target.value
        });
    }

    private onSecYesNoChanged = (event: any) => {
        this.setState({
            secSurYesNo: !this.state.secSurYesNo
        });
    }

    private onSecYesNoQuestionChanged = (event: any) => {
        this.setState({
            secYesNoQuestion: event.target.value
        });
    }

    private onSecToggleLinktoSurveyChanged = (event: any) => {
        this.setState({
            secSurLinkToSurvey: !this.state.secSurLinkToSurvey
        });
    }

    private onSecLinktoSurveyChanged = (event: any) => {
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

        this.setState({
            secLinkToSurvey: event.target.value
        });
    }
}

const newMessageWithTranslation = withTranslation()(NewMessage);
export default newMessageWithTranslation;
