// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from "react";
import { RouteComponentProps } from "react-router-dom";
import { useTranslation } from 'react-i18next';
import { Text, Button, Flex, Dropdown } from "@fluentui/react-northstar";
import { PrimaryButton } from '@fluentui/react';
import * as microsoftTeams from "@microsoft/teams-js";
import { TFunction } from "i18next";
import i18n from "../../i18n";
import { getUsers, updateUserPreference } from '../../apis/messageListApi';

//const Preference: React.FunctionComponent<RouteComponentProps> = props => {

//    useEffect(() => {
//        microsoftTeams.initialize();

//        microsoftTeams.getContext(context => {
//            const aadId = context.userObjectId ? context.userObjectId : "";
//            console.log(aadId);
//        });
//    });
//async function test() {
//    let entitiesIter = clientWithSAS.listEntities();
//    let i = 1;
//    for await (const entity of entitiesIter) {
//        console.log(`Entity${i}: PartitionKey: ${entity.partitionKey} RowKey: ${entity.rowKey}`);
//        i++;
//    }
//}

interface MyProps {
    partitionKey?: string,
    rowKey?: string,
    aadId?: string,
    userId?: string,
    conversationId?: string,
    tenantId?: string,
    preference?: string,
    userType?: string,
}

interface formState {
    partitionKey?: string,
    rowKey?: string,
    aadId?: string,
    userId?: string,
    conversationId?: string,
    tenantId?: string,
    preference?: string,
    userType?: string,
    save?: string,
}

export interface IUserPreference {
    partitionKey?: string,
    rowKey?: string,
    aadId?: string,
    userId?: string,
    conversationId?: string,
    tenantId?: string,
    preference?: string,
    userType?: string,
}

class Preference extends React.Component<MyProps, formState> {
    constructor(props: any) {
        super(props);
        this.state = {
            partitionKey: "",
            rowKey: "",
            aadId: "",
            userId: "",
            conversationId: "",
            tenantId: "",
            preference: "",
            userType: "",
            save: "",
        }
    }
    componentDidMount() {
        microsoftTeams.initialize();

        microsoftTeams.getContext(context => {
            const aadId = context.userObjectId ? context.userObjectId : "";
            console.log("id: " + aadId);
            this.getItem(aadId).then(() => {
                //console.log(this.state.test)
            });
        });
        
    }
    private selectLanguage = (event: any, itemsData: any) => {
        this.setState({
            preference: itemsData.value,
            save: ""
        })
    }

    private getItem = async (id: string) => {
        try {
            const response = await getUsers(id);
            this.setState({
                partitionKey: response.data[0].partitionKey,
                rowKey: response.data[0].rowKey,
                aadId: response.data[0].aadId,
                userId: response.data[0].userId,
                conversationId: response.data[0].conversationId,
                tenantId: response.data[0].tenantId,
                preference: response.data[0].preference,
                userType: response.data[0].userType,
                // id: response.data[0].id,
                // aadId: response.data[0].aadId,
                // preference: response.data[0].preference
            });
            //console.log("res: " + JSON.stringify(response));//.preference))response.data[0]
            console.log("res: " + JSON.stringify(response.data[0]));
        } catch (error) {
            return error;
        }
    }

    private onSave = () => {
        const userPreference: IUserPreference = {
            partitionKey: this.state.partitionKey,
            rowKey: this.state.rowKey,
            aadId: this.state.aadId,
            userId: this.state.userId,
            conversationId: this.state.conversationId,
            tenantId: this.state.tenantId,
            preference: this.state.preference,
            userType: this.state.userType,
        };
        this.editPreference(userPreference).then(() => {
            //microsoftTeams.tasks.submitTask(this.refreshPage);
            this.setState({
                save: "Your preference has been saved!"
            });
        });
    }

    private editPreference = async (userPreference: IUserPreference) => {
        try {
            await updateUserPreference(userPreference);
        } catch (error) {
            return error;
        }
    }
    
    public render(): JSX.Element {
        const language = ["English", "Burmese", "Traditional Chinese", "Simplified Chinese", "Malay", "Filipino", "Indonesian", "Korea", "Vietnamese"];
        
        return (
            <div style={{ display: 'flex', justifyContent: 'center', padding: '25px' }}>
                <Flex column vAlign="stretch" gap="gap.small" hAlign="center">
                    <Flex gap="gap.smaller" vAlign="center">
                        <Text content="Prefered Language" />
                        {/*<select>*/}
                        {/*    <option value="eng">English</option>*/}
                        {/*    <option value="cht">Traditional Chinese</option>*/}
                        {/*    <option value="chs">Simplified Chinese</option>*/}
                        {/*</select>*/}
                        <Dropdown
                            value={this.state.preference}
                            items={language}
                            onChange={this.selectLanguage}
                        />
                    </Flex>
                    <Flex gap="gap.smaller" vAlign="end" hAlign="center">
                        <PrimaryButton type="submit" value="Submit" text="Submit" onClick={this.onSave}/>
                        {/*<Button onClick={test} content="Test"/>*/}
                    </Flex>
                    <Flex>
                        <Text content={this.state.save} />
                    </Flex>
                </Flex>
            </div>
        );
    }
};

export default Preference;
