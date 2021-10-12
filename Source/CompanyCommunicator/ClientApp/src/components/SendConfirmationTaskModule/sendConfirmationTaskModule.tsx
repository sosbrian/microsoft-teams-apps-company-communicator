// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import { Loader, Button, Text, List, Image, Flex, Checkbox } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";

import './sendConfirmationTaskModule.scss';
import { getDraftNotification, getConsentSummaries, sendDraftNotification } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn, setCardBtn2, setCardBtn3, setCardBtn4, setCardBtn5, setVideoBtn, setSenderTemplate, setBoldSummary, setFontSummary, setAlignmentSummary, setFontSizeSummary, setFontColorSummary, setReaction, SetReactionQuestion, SetFreeText, SetFreeTextQuestion, setYesNo, setYesNoQuestion, setSubmitBtn, setToggleLinktoSurvey, setLinktoSurvey
} from '../AdaptiveCard/adaptiveCard';
import { ImageUtil } from '../../utility/imageutility';
import { TFunction } from "i18next";
import { Switch, FormControlLabel } from '@material-ui/core';
import { Check } from 'office-ui-fabric-react';

export interface IListItem {
    header: string,
    media: JSX.Element,
}

export interface IMessage {
    id: string;
    title: string;
    template?: string;
    senderTemplate?: string;
    acknowledgements?: number;
    reactions?: number;
    responses?: number;
    succeeded?: number;
    failed?: number;
    throttled?: number;
    sentDate?: string;
    imageLink?: string;
    videoLink?: string;
    summary?: string;
    alignment?: string;
    boldSummary?: string;
    fontSummary?: string;
    fontSizeSummary?: string;
    fontColorSummary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    buttonLink2?: string;
    buttonTitle2?: string;
    buttonLink3?: string;
    buttonTitle3?: string;
    buttonLink4?: string;
    buttonTitle4?: string;
    buttonLink5?: string;
    buttonTitle5?: string;
    surReaction?: boolean;
    reactionQuestion?: string;
    surFreeText?: boolean;
    freeTextQuestion?: string;
    surYesNo?: boolean;
    yesNoQuestion?: string;
    surLinkToSurvey?: boolean;
    linkToSurvey?: string;
}

export interface SendConfirmationTaskModuleProps extends RouteComponentProps, WithTranslation {
}

export interface IStatusState {
    message: IMessage;
    loader: boolean;
    teamNames: string[];
    rosterNames: string[];
    groupNames: string[];
    uploadedList: string[];
    uploadedListName: string;
    emailOption: boolean;
    allUsers: boolean;
    messageId: number;
    exclusionList: string;
}

class SendConfirmationTaskModule extends React.Component<SendConfirmationTaskModuleProps, IStatusState> {
    readonly localize: TFunction;
    private initMessage = {
        id: "",
        title: ""
    };

    private card: any;

    constructor(props: SendConfirmationTaskModuleProps) {
        super(props);
        this.localize = this.props.t;
        this.card = getInitAdaptiveCard(this.localize);

        this.state = {
            message: this.initMessage,
            loader: true,
            teamNames: [],
            rosterNames: [],
            groupNames: [],
            allUsers: false,
            uploadedList: [],
            uploadedListName: "",
            emailOption: false,
            messageId: 0,
            exclusionList: ""
        };
    }

    public componentDidMount() {
        microsoftTeams.initialize();

        let params = this.props.match.params;

        if ('id' in params) {
            let id = params['id'];
            this.getItem(id).then(() => {
                getConsentSummaries(id).then((response) => {
                    console.log(response.data);
                    this.setState({
                        teamNames: response.data.teamNames.sort(),
                        rosterNames: response.data.rosterNames.sort(),
                        groupNames: response.data.groupNames.sort(),
                        allUsers: response.data.allUsers,
                        uploadedList: response.data.uploadedList,
                        uploadedListName: response.data.uploadedListName,
                        emailOption: response.data.emailOption,
                        exclusionList: response.data.exclusionList,
                        messageId: id,
                    }, () => {
                        this.setState({
                            loader: false
                        }, () => {
                            setCardTitle(this.card, this.state.message.title);
                            setCardImageLink(this.card, this.state.message.imageLink);
                            setCardSummary(this.card, this.state.message.summary);
                            setCardAuthor(this.card, this.state.message.author);
                            //if (this.state.message.buttonTitle && this.state.message.buttonLink) {
                            //    setCardBtn(this.card, this.state.message.buttonTitle, this.state.message.buttonLink);
                            //}
                            setCardBtn(this.card, this.state.message.buttonTitle, this.state.message.buttonLink);
                            setCardBtn2(this.card, this.state.message.buttonTitle2, this.state.message.buttonLink2);
                            setCardBtn3(this.card, this.state.message.buttonTitle3, this.state.message.buttonLink3);
                            setCardBtn4(this.card, this.state.message.buttonTitle4, this.state.message.buttonLink4);
                            setCardBtn5(this.card, this.state.message.buttonTitle5, this.state.message.buttonLink5);
                            setVideoBtn(this.card, this.state.message.videoLink);
                            setSenderTemplate(this.card, this.state.message.senderTemplate);
                            setBoldSummary(this.card, this.state.message.boldSummary);
                            setFontSummary(this.card, this.state.message.fontSummary);
                            setAlignmentSummary(this.card, this.state.message.alignment);
                            setFontSizeSummary(this.card, this.state.message.fontSizeSummary);
                            setFontColorSummary(this.card, this.state.message.fontColorSummary);
                            setReaction(this.card, this.state.message.surReaction);
                            SetReactionQuestion(this.card, this.state.message.reactionQuestion);
                            SetFreeText(this.card, this.state.message.surFreeText);
                            SetFreeTextQuestion(this.card, this.state.message.freeTextQuestion);
                            setYesNo(this.card, this.state.message.surYesNo);
                            setYesNoQuestion(this.card, this.state.message.yesNoQuestion);
                            setToggleLinktoSurvey(this.card, this.state.message.surLinkToSurvey);
                            setLinktoSurvey(this.card, this.state.message.linkToSurvey);
                            if (this.state.message.surReaction || this.state.message.surFreeText || this.state.message.surYesNo) {
                                setSubmitBtn(this.card, true);
                            }

                            let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                            adaptiveCard.parse(this.card);
                            let renderedCard = adaptiveCard.render();
                            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                            //if (this.state.message.buttonLink) {
                            //    let link = this.state.message.buttonLink;
                            //    adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                            //}
                            if (this.state.message.buttonLink || this.state.message.buttonLink2 || this.state.message.buttonLink3 || this.state.message.buttonLink4 || this.state.message.buttonLink5) {
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
                        });
                    });
                });
            });
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getDraftNotification(id);
            this.setState({
                message: response.data
            });
        } catch (error) {
            return error;
        }
    }

    public render(): JSX.Element {
        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            return (
                <div className="taskModule">
                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                        <Flex className="scrollableContent" gap="gap.small">
                            <Flex.Item size="size.half">
                                <Flex column className="formContentContainer">
                                    <h3>{this.localize("ConfirmToSend")}</h3>
                                    <span>{this.localize("SendToRecipientsLabel")}</span>

                                    <div className="results">
                                        {this.renderAudienceSelection()}
                                    </div>
                                    <div className="results">
                                        <div key="exclusionList" hidden={!this.state.exclusionList}>
                                            <span className="label">Exclusion List</span>
                                        </div>
                                        {this.renderExclusionList()}
                                    </div>
                                    <div className="results">
                                        <Checkbox label="Send email to members" checked={this.state.emailOption} disabled />
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
                                    <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={this.localize("PreparingMessageLabel")} labelPosition="end" />
                                </Flex.Item>
                                <Button content={this.localize("Send")} id="sendBtn" onClick={this.onSendMessage} primary />
                                {/*onClick={this.onSendMessage}*/}
                            </Flex>
                        </Flex>
                    </Flex>
                </div>
            );
        }
    }

    private onSendMessage = () => {
        let spanner = document.getElementsByClassName("sendingLoader");
        spanner[0].classList.remove("hiddenLoader");
        sendDraftNotification(this.state.message).then(() => {
            microsoftTeams.tasks.submitTask();
        });
    }

    private getItemList = (items: string[]) => {
        let resultedTeams: IListItem[] = [];
        if (items) {
            resultedTeams = items.map((element) => {
                const resultedTeam: IListItem = {
                    header: element,
                    media: <Image src={ImageUtil.makeInitialImage(element)} avatar />
                }
                return resultedTeam;
            });
        }
        return resultedTeams;
    }

    private renderAudienceSelection = () => {
        if (this.state.teamNames && this.state.teamNames.length > 0) {
            return (
                <div key="teamNames"> <span className="label">{this.localize("TeamsLabel")}</span>
                    <List items={this.getItemList(this.state.teamNames)} />
                </div>
            );
        } else if (this.state.rosterNames && this.state.rosterNames.length > 0) {
            return (
                <div key="rosterNames"> <span className="label">{this.localize("TeamsMembersLabel")}</span>
                    <List items={this.getItemList(this.state.rosterNames)} />
                </div>
                );
        } else if (this.state.groupNames && this.state.groupNames.length > 0) {
            return (
                <div key="groupNames" > <span className="label">{this.localize("GroupsMembersLabel")}</span>
                    <List items={this.getItemList(this.state.groupNames)} />
                </div>);
        } else if (this.state.allUsers) {
            return (
                <div key="allUsers">
                    <span className="label">{this.localize("AllUsersLabel")}</span>
                    <div className="noteText">
                        <Text error content={this.localize("SendToAllUsersNote")} />
                    </div>
                </div>);
        } else if (this.state.uploadedList) {
            return (
                <div key="uploadedList">
                    <span className="label">Custom Uploaded List</span>
                    <div>{this.state.uploadedListName}</div>
                </div>);
        } else {
            return (<div></div>);
        }
    }

    private renderExclusionList = () => {
        var exArray = this.state.exclusionList.split(';');
        if (this.state.exclusionList) {
            return (
                exArray.map((item) => <div>{item}</div> )
            );
        } else {
            return (
                <div></div>
            );
        }
    }
}

const sendConfirmationTaskModuleWithTranslation = withTranslation()(SendConfirmationTaskModule);
export default sendConfirmationTaskModuleWithTranslation;