// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from "react";
import { RouteComponentProps } from "react-router-dom";
//import { Button, Loader, Dropdown, Text, Flex, Input, TextArea, RadioGroup, FlexItem } from '@fluentui/react-northstar'
import { useTranslation } from 'react-i18next';
import { Text, Button, Flex } from "@fluentui/react-northstar";
import { PrimaryButton } from '@fluentui/react';
import * as microsoftTeams from "@microsoft/teams-js";
import i18n from "../../i18n";
import { BlobServiceClient, ContainerClient } from '@azure/storage-blob';
import { TableClient, AzureNamedKeyCredential, AzureSASCredential  } from "@azure/data-tables";

// connectionString = "DefaultEndpointsProtocol=https;AccountName=45u3yv4vigkqc;AccountKey=S8JCJe3RQvxzRtvRhuo2WxijZyHuwEU7zjep9pbmngagnVXt6dD6BojB7LqyjouA6zWrI4C08y4BtIoBdt01jg==;EndpointSuffix=core.windows.net";
//const client = TableClient.fromConnectionString(connectionString, "UserData");
const account = "45u3yv4vigkqc";
const sas = "?sv=2020-08-04&ss=t&srt=sco&sp=rwlacu&se=2022-01-12T16:45:27Z&st=2022-01-12T08:45:27Z&spr=https&sig=seF180D2sRhdzE5mJk50tNMHQLisszaVuHVE0JvT2Eo%3D";
//const accountKey = "S8JCJe3RQvxzRtvRhuo2WxijZyHuwEU7zjep9pbmngagnVXt6dD6BojB7LqyjouA6zWrI4C08y4BtIoBdt01jg==";
const tableName = "UserData";

//const credential = new AzureNamedKeyCredential(account, accountKey);
//const client = new TableClient(`https://${account}.table.core.windows.net`, tableName, credential);
const clientWithSAS = new TableClient(
    'https://${account}.table.core.windows.net',
    tableName,
    new AzureSASCredential(sas)
);


//const Preference: React.FunctionComponent<RouteComponentProps> = props => {

//    useEffect(() => {
//        microsoftTeams.initialize();

//        microsoftTeams.getContext(context => {
//            const aadId = context.userObjectId ? context.userObjectId : "";
//            console.log(aadId);
//        });
//    });
async function test() {
    let entitiesIter = clientWithSAS.listEntities();
    let i = 1;
    for await (const entity of entitiesIter) {
        console.log(`Entity${i}: PartitionKey: ${entity.partitionKey} RowKey: ${entity.rowKey}`);
        i++;
    }
}

class Preference extends React.Component {
    constructor(props: any) {
        super(props);
    }
    componentDidMount() {
        microsoftTeams.initialize();

        microsoftTeams.getContext(context => {
            const aadId = context.userObjectId ? context.userObjectId : "";
            console.log("id" + aadId);
        });
    }
    
    public render(): JSX.Element {
        return (
            <div style={{ display: 'flex', justifyContent: 'center', padding: '25px' }}>
                <Flex column vAlign="stretch" gap="gap.small" hAlign="center">
                    <Flex gap="gap.smaller" vAlign="end">
                        <Text content="Prefered Language" />
                        <select>
                            <option value="eng">English</option>
                            <option value="cht">Traditional Chinese</option>
                            <option value="chs">Simplified Chinese</option>
                        </select>
                    </Flex>
                    <Flex gap="gap.smaller" vAlign="end" hAlign="center">
                        <PrimaryButton type="submit" value="Submit" text="Submit" />
                        <Button onClick={test} content="Test"/>
                    </Flex>
                </Flex>
            </div>
        );
    }
};

export default Preference;
