import { Button, Text } from "@fluentui/react-components";
import {
    TeamsUserCredential,
    createMicrosoftGraphClientWithCredential
} from "@microsoft/teamsfx";

import FilteteredResult from './custom/FilteredResult';
import React from "react";
import Suppliers from "./custom/Suppliers"
import { app } from "@microsoft/teams-js";
import config from "../lib/config";
import { getAttachmentFromLocalStorage } from "./helper/helper";

export default class LaunchPage extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            userInfo: {},
            actionId: "01JVL355MQ3LUVYVKHMZEJIV2AOP372PEM",
            actionItem: undefined,
            showLoginPage: undefined,
            sheetData: undefined,
            suppliers: [],
            filteredSupplierList: undefined
        };
        this.clearFilter = this.clearFilter.bind(this);
    }

    async componentDidMount() {
        await this.initTeamsFx();
        await this.initData();
        await this.fetchData();
    }

    async initTeamsFx() {
        const authConfig = {
            clientId: config.clientId,
            initiateLoginEndpoint: config.initiateLoginEndpoint,
        };

        const credential = new TeamsUserCredential(authConfig);
        const userInfo = await credential.getUserInfo();

        this.setState({
            userInfo: userInfo,
        });

        this.scope = ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Files.Read"]; //Files.Read.All
        this.credential = credential;
    }

    async initData() {
        await this.checkIsConsentNeeded();
    }

    async fetchData() {
        try {
            const context = await app.getContext();
            const objectId = context.user && context.user.id;
            let actionId = context.actionInfo && context.actionInfo.actionObjects[0].itemId;
            if (!actionId) {
                actionId = this.state.actionId;
            }
            this.setState({ actionId: actionId });
            // Get Microsoft graph client
            const graphClient = createMicrosoftGraphClientWithCredential(
                this.credential,
                this.scope
            );
            await this.readActionItem(graphClient, objectId, actionId);
            const sheetData = await this.readActionItemData(graphClient, objectId, actionId);

            const response = await fetch("https://services.odata.org/V4/Northwind/Northwind.svc/Suppliers");
            const data = await response.json();
            this.setState({ suppliers: data.value });

            const filteredSupplierList = data.value.map((item, index) => {
                const isExist = sheetData.find(element => element[0].toLowerCase() === item.CompanyName.toLowerCase());
                if (isExist) {
                    return item;
                }
                return undefined;
            });

            const attachments = getAttachmentFromLocalStorage();
            const filteredSupplierListWithAttachment = filteredSupplierList.filter(x => x !== undefined).map(y => {
                const attachmentName = attachments && attachments.storeData.find(item => item.selectedSuplierCompanyName === y.CompanyName);
                return {
                    CompanyName: y.CompanyName,
                    ContactName: y.ContactName,
                    Phone: y.Phone,
                    Country: y.Country,
                    Attachments: attachmentName ? attachmentName.actionItemName : ""
                }
            });

            this.setState({ filteredSupplierList: filteredSupplierListWithAttachment });
        } catch (error) {
            console.log(error);
        }
    }

    async readActionItem(graphClient, objectId, actionId) {
        try {
            const actionItem = await graphClient.api(`/users/${objectId}/drive/items/${actionId}`).get();
            this.setState({
                actionItem: actionItem
            });
            console.log("driveData", actionItem);
        } catch (error) {
            console.log(error);
        }
    }
    async readActionItemData(graphClient, objectId, actionId) {
        try {
            const worksheets = (await graphClient.api(`/users/${objectId}/drive/items/${actionId}/workbook/worksheets`).get()).value;
            const sheetData = (await graphClient.api(`/users/${objectId}/drive/items/${actionId}/workbook/worksheets('${worksheets[0].name}')/usedRange`).get()).values;
            this.setState({ sheetData: sheetData });
            return sheetData;
        } catch (error) {
            console.log(error);
        }
    }
    async checkIsConsentNeeded() {
        try {
            await this.credential.getToken(this.scope);
        } catch (error) {
            this.setState({
                showLoginPage: true,
            });
            return true;
        }
        this.setState({
            showLoginPage: false,
        });
        return false;
    }
    async loginBtnClick() {
        try {
            // Popup login page to get user's access token
            await this.credential.login(this.scope);
        } catch (err) {
            console.log(err);
            if (err instanceof Error && err.message?.includes("CancelledByUser")) {
                const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
                err.message +=
                    '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
                    "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
                    `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}`;
            }

            alert("Login failed: " + err);
            return;
        }
        await this.initData();
    }
    async clearFilter() {
        this.setState({
            actionId: undefined,
            actionItem: undefined,
        })
    }

    render() {
        return (
            <div>
                {this.state.showLoginPage === false && (
                    <>
                        <FilteteredResult
                            actionId={this.state.actionId}
                            actionItem={this.state.actionItem}
                            sheetData={this.state.filteredSupplierList}
                            clearFilter={this.clearFilter}

                        />
                        <Suppliers suppliers={this.state.filteredSupplierList} />
                    </>
                )}
                {this.state.showLoginPage === true && (
                    <div className="auth">
                        <Text>Welcome Northwind App!</Text>
                        <Button appearance="primary" onClick={() => this.loginBtnClick()}>
                            Start
                        </Button>
                    </div>
                )}
            </div>
        );
    }
}