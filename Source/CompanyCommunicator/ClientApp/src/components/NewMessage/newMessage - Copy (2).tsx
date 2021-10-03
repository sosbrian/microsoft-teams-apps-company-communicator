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
import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, searchGroups, getGroups, verifyGroupAccess } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn
} from '../AdaptiveCard/adaptiveCard';
import { getTempAdaptiveCard } from '../AdaptiveCard/adaptiveCard1';
import { getBaseUrl } from '../../configVariables';
import { ImageUtil } from '../../utility/imageutility';
import { TFunction } from "i18next";
//Using material-ui PublishIcon
import IconButton from '@material-ui/core/IconButton';
import PublishIcon from '@material-ui/icons/Publish';

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
    id?: string,
    title: string,
    imageLink?: string,
    summary?: string,
    author: string,
    buttonTitle?: string,
    buttonLink?: string,
    buttonTitle1?: string,
    buttonLink1?: string,
    teams: any[],
    rosters: any[],
    groups: any[],
    allUsers: boolean
}

export interface formState {
    template: string,
    title: string,
    summary?: string,
    btnLink?: string,
    btnLink1?: string,
    imageLink?: string,
    btnTitle?: string,
    btnTitle1?: string,
    author: string,
    card?: any,
    page: string,
    teamsOptionSelected: boolean,
    rostersOptionSelected: boolean,
    allUsersOptionSelected: boolean,
    groupsOptionSelected: boolean,
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
    errorImageUrlMessage: string,
    errorButtonUrlMessage: string,
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
                        }
                    ],
                        "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.0"
        };
        //this.setDefaultCard(this.card);

        this.state = {
            template: "",
            title: "",
            summary: "",
            author: "",
            btnLink: "",
            btnLink1: "",
            imageLink: "",
            btnTitle: "",
            btnTitle1: "",
            card: this.card,
            page: "CardCreation",
            teamsOptionSelected: true,
            rostersOptionSelected: false,
            allUsersOptionSelected: false,
            groupsOptionSelected: false,
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
            errorImageUrlMessage: "",
            errorButtonUrlMessage: "",
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
        const titleAsString = this.localize("TitleText");
        const summaryAsString = this.localize("Summary");
        const authorAsString = this.localize("Author1");
        const buttonTitleAsString = this.localize("ButtonTitle");

        setCardTitle(card, titleAsString);
        let imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
        setCardImageLink(card, imgUrl);
        setCardSummary(card, summaryAsString);
        setCardAuthor(card, authorAsString);
        setCardBtn(card, buttonTitleAsString, "https://adaptivecards.io");
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
            const response = await getDraftNotification(id);
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
            this.setState({
                teamsOptionSelected: draftMessageDetail.teams.length > 0,
                selectedTeamsNum: draftMessageDetail.teams.length,
                rostersOptionSelected: draftMessageDetail.rosters.length > 0,
                selectedRostersNum: draftMessageDetail.rosters.length,
                groupsOptionSelected: draftMessageDetail.groups.length > 0,
                selectedGroupsNum: draftMessageDetail.groups.length,
                selectedRadioBtn: selectedRadioButton,
                selectedTeams: draftMessageDetail.teams,
                selectedRosters: draftMessageDetail.rosters,
                selectedGroups: draftMessageDetail.groups
            });

            setCardTitle(this.card, draftMessageDetail.title);
            setCardImageLink(this.card, draftMessageDetail.imageLink);
            setCardSummary(this.card, draftMessageDetail.summary);
            setCardAuthor(this.card, draftMessageDetail.author);
            setCardBtn(this.card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink);

            this.setState({
                title: draftMessageDetail.title,
                summary: draftMessageDetail.summary,
                btnLink: draftMessageDetail.buttonLink,
                imageLink: draftMessageDetail.imageLink,
                btnTitle: draftMessageDetail.buttonTitle,
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
            const hiddenDiv = {
                display: 'block',
            };
            const handleSelect = (e: any) => {
                console.log(e);
            }
        const templates = [{ key: "1", header: "Announcements with image" }, { key: "2", header: "Announcement with video" }, { key: "3", header: "Announcement in plain text" }, { key: "4", header: "Event promotion with sessions to register/join" }];

            if (this.state.loader) {
                return (
                    <div className="Loader">
                        <Loader />
                    </div>
                );
            } else {
                if (this.state.page === "CardCreation") {
                    return (
                        <div className="taskModule">
                            <Flex column className="formContainer1" vAlign="stretch" gap="gap.small">
                                <Flex className="scrollableContent" vAlign="center" hAlign="start">
                                    <br /><br />
                                    <Flex.Item>
                                        <Flex column className="formContentContainer">
                                            
                                            <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                        </Flex>
                                    </Flex.Item>
                                    <Flex.Item>
                                        
                                        <Flex column>
                                            <Dropdown
                                                placeholder={this.localize("Template")}
                                                items={templates}
                                                value={this.state.template}
                                                onChange={this.onTemplateChange}
                                                noResultsMessage={this.localize("NoMatchMessage")}
                                            />
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

                
                    /*return (
                        
                        <div className="taskModule">
                            <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                                <Flex className="scrollableContent">
                                    <Flex.Item size="size.half">
                                        <Flex column className="formContentContainer">
                                            <Input className="inputField"
                                                value={this.state.title}
                                                label={this.localize("TitleText")}
                                                placeholder={this.localize("PlaceHolderTitle")}
                                                onChange={this.onTitleChanged}
                                                autoComplete="off"
                                                fluid
                                            />

                                            <Button
                                                id="hiddenDivBtn"
                                                onClick={this.hiddenField}
                                            />

                                            <div className="test2" style={hiddenDiv}>
                                                <Input fluid className="inputField2"
                                                    value={this.state.imageLink}
                                                    label={this.localize("ImageURL")}
                                                    placeholder={this.localize("ImageURL")}
                                                    onChange={this.onImageLinkChanged}
                                                    error={!(this.state.errorImageUrlMessage === "")}
                                                    autoComplete="off"
                                                />
                                                <IconButton color="primary" aria-label="upload picture" component="span">
                                                    <PublishIcon />
                                                </IconButton>
                                            </div>
                                            <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />

                                            <div className="textArea">
                                                <Text content={this.localize("Summary")} />
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
                                            <Input className="inputField"
                                                fluid
                                                value={this.state.btnTitle1}
                                                label={this.localize("ButtonTitle1")}
                                                placeholder={this.localize("ButtonTitle1")}
                                                onChange={this.onBtnTitle1Changed}
                                                autoComplete="off"
                                            />
                                            <Input className="inputField"
                                                fluid
                                                value={this.state.btnLink1}
                                                label={this.localize("ButtonURL1")}
                                                placeholder={this.localize("ButtonURL1")}
                                                onChange={this.onBtnLink1Changed}
                                                error={!(this.state.errorButtonUrlMessage === "")}
                                                autoComplete="off"
                                            />
                                            <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                        </Flex>
                                    </Flex.Item>
                                    <Flex.Item size="size.half">
                                        <div className="adaptiveCardContainer">
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
                    );*/
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
                                                    }
                                                ]}
                                            >

                                            </RadioGroup>
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
            allUsersOptionSelected: data.value === 'allUsers',
            selectedTeams: data.value === 'teams' ? this.state.selectedTeams : [],
            selectedTeamsNum: data.value === 'teams' ? this.state.selectedTeamsNum : 0,
            selectedRosters: data.value === 'rosters' ? this.state.selectedRosters : [],
            selectedRostersNum: data.value === 'rosters' ? this.state.selectedRostersNum : 0,
            selectedGroups: data.value === 'groups' ? this.state.selectedGroups : [],
            selectedGroupsNum: data.value === 'groups' ? this.state.selectedGroupsNum : 0,
        });
    }

    private isSaveBtnDisabled = () => {
        const teamsSelectionIsValid = (this.state.teamsOptionSelected && (this.state.selectedTeamsNum !== 0)) || (!this.state.teamsOptionSelected);
        const rostersSelectionIsValid = (this.state.rostersOptionSelected && (this.state.selectedRostersNum !== 0)) || (!this.state.rostersOptionSelected);
        const groupsSelectionIsValid = (this.state.groupsOptionSelected && (this.state.selectedGroupsNum !== 0)) || (!this.state.groupsOptionSelected);
        const nothingSelected = (!this.state.teamsOptionSelected) && (!this.state.rostersOptionSelected) && (!this.state.groupsOptionSelected) && (!this.state.allUsersOptionSelected);
        return (!teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected)
    }

    private isNextBtnDisabled = () => {
        const title = this.state.title;
        const btnTitle = this.state.btnTitle;
        const btnLink = this.state.btnLink;
        return !(title && ((btnTitle && btnLink) || (!btnTitle && !btnLink)) && (this.state.errorImageUrlMessage === "") && (this.state.errorButtonUrlMessage === ""));
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
            selectedGroupsNum: 0
        })
    }

    private onTemplateChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        //this.setState({
        //    selectedTeams: itemsData.value,
        //    selectedTeamsNum: itemsData.value.length,
        //    selectedRosters: [],
        //    selectedRostersNum: 0,
        //    selectedGroups: [],
        //    selectedGroupsNum: 0
        //})
        console.log(itemsData.value.header);
        if (itemsData.value.key === "1") {
            this.setState({
                template: itemsData.value.header,
                card:
                {
                    "type": "AdaptiveCard",
                    "body": [
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
                        }
                    ],
                    "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.0"
                }
            }, () => {
                this.updateCard();
            }
            );

        } else if (itemsData.value.key === "2") {
            this.setState({
                template: itemsData.value.header,
                card:
                {
                    "type": "AdaptiveCard",
                    "body": [
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
                        }
                    ],
                    "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.0"
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
                        }
                    ],
                    "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.0"
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
                        }
                    ],
                    "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.0"
                }
            }, () => {
                this.updateCard();
            });


        }
    }

    //private hiddenField = (event: any) => {
    //    const hiddenDiv = "display:none";
    //    console.log(hiddenDiv);
    //}

    private onRostersChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedRosters: itemsData.value,
            selectedRostersNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0
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
            selectedRostersNum: 0
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
            id: this.state.messageId,
            title: this.state.title,
            imageLink: this.state.imageLink,
            summary: this.state.summary,
            author: this.state.author,
            buttonTitle: this.state.btnTitle,
            buttonLink: this.state.btnLink,
            teams: selectedTeams,
            rosters: selctedRosters,
            groups: selectedGroups,
            allUsers: this.state.allUsersOptionSelected
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
        this.setState({
            page: "AudienceSelection"
        }, () => {
            this.updateCard();
        });
    }

    private onBack = (event: any) => {
        this.setState({
            page: "CardCreation"
        }, () => {
            this.updateCard();
        });
    }

    private onTitleChanged = (event: any) => {
        let showDefaultCard = (!event.target.value && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, event.target.value);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
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

        let showDefaultCard = (!this.state.title && !event.target.value && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle1 && !this.state.btnLink1);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, event.target.value);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
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

    private onSummaryChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !event.target.value && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle1 && !this.state.btnLink1);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, event.target.value);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
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
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !event.target.value && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle1 && !this.state.btnLink1);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, event.target.value);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
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
        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnLink && !this.state.btnTitle1 && !this.state.btnLink1);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
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
            delete this.card.actions;
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

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !event.target.value && !this.state.btnTitle1 && !this.state.btnLink1);
        setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
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
            delete this.card.actions;
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
    //btnLink1 and btnTitle1 OnChanged
    private onBtnTitle1Changed = (event: any) => {
        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnTitle && !this.state.btnLink && !this.state.btnLink1);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        if (event.target.value && this.state.btnLink1) {
            setCardBtn(this.card, event.target.value, this.state.btnLink1);
            this.setState({
                btnTitle1: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnTitle1: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLink1Changed = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
        }

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !event.target.value && !this.state.btnTitle1);
        setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        if (this.state.btnTitle1 && event.target.value) {
            setCardBtn(this.card, this.state.btnTitle1, event.target.value);
            this.setState({
                btnLink1: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
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
        const link = this.state.btnLink;
        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
        console.log(renderedCard);
    }
    private updateCard1 = () => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(this.state.card);
        const renderedCard = adaptiveCard.render();
        const container = document.getElementsByClassName('adaptiveCardContainer')[0].firstChild;
        if (container != null) {
            container.replaceWith(renderedCard);
            console.log("not null");
        } else {
            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
            console.log("null");
        }
        const link = this.state.btnLink;
        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
        console.log(renderedCard);
        console.log(this.state.card);
    }
}

const newMessageWithTranslation = withTranslation()(NewMessage);
export default newMessageWithTranslation;
