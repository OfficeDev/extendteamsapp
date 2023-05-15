import './Dialog.css';

import * as fabIcons from '@uifabric/icons';

import { Button, Link, Text } from "@fluentui/react-components";
import { TeamsUserCredential, createMicrosoftGraphClientWithCredential } from "@microsoft/teamsfx";
import { app, dialog } from "@microsoft/teams-js";

import Attachment from './custom/Attachment';
import { Document16Regular } from "@fluentui/react-icons";
import { Icon } from '@fluentui/react/lib/Icon';
import React from "react";
import config from "../lib/config";
import { setValuesToLocalStorage } from './helper/helper';

class DialogPage extends React.Component {
    constructor(props) {
        super(props);
        fabIcons.initializeIcons();
        this.state = {
            userInfo: {},
            actionId: "01JVL355JYVRKAWYPNWBCLB2GFIMNNFFTK",
            actionItem: undefined,
            showLoginPage: undefined,
            sheetData: undefined,
            suppliers: [],
            filteredSupplierList: undefined,
            selectedSupplier: undefined,
        };
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

            const response = await fetch("https://services.odata.org/V4/Northwind/Northwind.svc/Suppliers");
            const data = await response.json();
            this.setState({ suppliers: data.value });
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
    async onSubmit(actionItem, selectedSuplier) {

        const json = {
            selectedSuplierCompanyName: selectedSuplier.CompanyName,
            actionItem: actionItem.id,
            actionItemName: actionItem.name
        }

        // Use const appIDs=['YOUR_APP_IDS_HERE']; instead of the following one
        // if you want to restrict which applications your dialog can submit to
        const appIDs = undefined;
        console.log(json);
        dialog.url.submit(json, appIDs);
        this.setState({
            selectedSupplier: undefined
        });
        setValuesToLocalStorage(json)
    }

    handleRowClick = (supplier) => {
        this.setState({ selectedSupplier: supplier });
        console.log(supplier);
    };
    render() {
        return (
            <div className="dialog" >
                {this.state.showLoginPage === false && this.state.suppliers && (
                    <form className="dialog_form">
                        <div className='dialog_header'>
                            {!this.state.selectedSupplier &&
                                <Text className="dialog_text" style={{ display: 'flex' }}>Select a Suplier :</Text>
                            }

                            {this.state.selectedSupplier &&
                                <>
                                    <Text className="dialog_text" style={{ display: 'flex' }}>
                                        {'Selected Supplier :'}
                                        <Text weight='bold' style={{ paddingLeft: "2px" }}>
                                            {this.state.selectedSupplier.CompanyName}
                                        </Text>
                                    </Text>

                                    <Button appearance="transparent" icon={<Icon iconName={"Cancel"} />} onClick={() => this.setState({ selectedSupplier: undefined })}></Button>
                                </>
                            }
                        </div>
                        <div style={{ paddingBottom: "20px" }}>
                            <div className="dialog_list">
                                {this.state.suppliers.length > 0 && this.state.suppliers.map(item => {
                                    return (
                                        <div className='dialog_listitem' key={item.SupplierID}>
                                            <Document16Regular />
                                            <Link style={{ fontSize: '12px', paddingLeft: "1px" }} onClick={() => this.handleRowClick(item)}>
                                                {item.CompanyName}
                                            </Link>
                                        </div>
                                    );
                                })}
                            </div>
                        </div>
                        {this.state.actionItem &&
                            <div className='dialog_attachment'>
                                <Text className='dialog_text'>Attach this document to the supplier:</Text>
                                <div>
                                    <Attachment actionItem={this.state.actionItem} appearance={"subtle"} />
                                </div>
                            </div>
                        }
                        <div className="dialog_button">
                            <Button type="button" appearance="primary" onClick={() => this.onSubmit(this.state.actionItem, this.state.selectedSupplier)}>Add</Button>
                        </div>
                    </form>
                )
                }
                {
                    this.state.showLoginPage === true && (
                        <div className="auth">
                            <Text>Authenticate</Text>
                            <Button appearance="primary" onClick={() => this.loginBtnClick()}>
                                Start
                            </Button>
                        </div>
                    )
                }
            </div >
        );
    }
}
export default DialogPage;