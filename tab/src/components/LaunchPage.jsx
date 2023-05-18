import { ActionObjectType, app } from "@microsoft/teams-js";
import {
    TeamsUserCredential,
    createMicrosoftGraphClientWithCredential
} from "@microsoft/teamsfx";
import { fetchSuppliers, filteredSupplierListWithAttachment, loginBtnClick } from "./helper/helper";

import FilteteredResult from './custom/FilteredResult';
import Login from "./custom/Login";
import MicrosoftGraph from './helper/graph';
import React from "react";
import Suppliers from "./custom/Suppliers"
import config from "../lib/config";

export default class LaunchPage extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            itemId: "01JVL355MQ3LUVYVKHMZEJIV2AOP372PEM",
            actionItem: undefined,
            showLoginPage: undefined,
            suppliers: [],
            filteredSupplierList: undefined
        };
        this.clearFilter = this.clearFilter.bind(this);
    }

    async componentDidMount() {
        await this.initTeamsFx();
        await this.checkIsConsentNeeded();
        await this.fetchData();
    }

    async initTeamsFx() {
        const authConfig = {
            clientId: config.clientId,
            initiateLoginEndpoint: config.initiateLoginEndpoint,
        };
        this.credential = new TeamsUserCredential(authConfig);
        this.scope = config.scopes;
    }

    async fetchData() {
        try {
            const context = await app.getContext();
            const userId = context.user && context.user.id;
            let itemId = context.actionInfo && context.actionInfo.actionObjects[0].itemId;

            if (!itemId) {
                itemId = this.state.itemId
            }
            this.setState({ itemId: itemId });

            // Get Microsoft graph client
            const graphClient = createMicrosoftGraphClientWithCredential(
                this.credential,
                this.scope
            );

            if (context.actionInfo.actionObjects[0].type === ActionObjectType.M365Content) {
                const msGraph = new MicrosoftGraph(graphClient, userId, itemId);

                const actionItem = await msGraph.readActionItem();
                this.setState({
                    actionItem: actionItem
                });

                const response = await fetchSuppliers();
                this.setState({ suppliers: response.value });

                const sheetData = await msGraph.readActionItemData();

                const filteredSupplierList = response.value.map((item, index) => {
                    const isExist = sheetData.find(element => element[0].toLowerCase() === item.CompanyName.toLowerCase());
                    if (isExist) {
                        return item;
                    }
                    return undefined;
                });

                const filteredSupplierListWithAttachmentData = filteredSupplierListWithAttachment(filteredSupplierList)

                this.setState({ filteredSupplierList: filteredSupplierListWithAttachmentData });
            }
        } catch (error) {
            console.log("fetchData", error);
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

    async loginBtn() {
        await loginBtnClick(this.credential, this.scope);
    }

    async clearFilter() {
        this.setState({
            itemId: undefined,
            actionItem: undefined,
        });
    }

    render() {
        return (
            <div>
                {this.state.showLoginPage === false && (
                    <>
                        <FilteteredResult
                            itemId={this.state.itemId}
                            actionItem={this.state.actionItem}
                            sheetData={this.state.filteredSupplierList}
                            clearFilter={this.clearFilter}

                        />
                        <Suppliers suppliers={this.state.filteredSupplierList} />
                    </>
                )}
                {this.state.showLoginPage === true && (
                    <Login loginBtnClick={this.loginBtn} />
                )}
            </div>
        );
    }
}