import './Dialog.css';

import * as fabIcons from '@uifabric/icons';

import { ActionObjectType, app, dialog } from '@microsoft/teams-js';
import { Button, Link, Text } from '@fluentui/react-components';
import { TeamsUserCredential, createMicrosoftGraphClientWithCredential } from '@microsoft/teamsfx';
import { fetchSuppliers, loginBtnClick, setValuesToLocalStorage } from './helper/helper';

import Attachment from './custom/Attachment';
import { Document16Regular } from '@fluentui/react-icons';
import { Icon } from '@fluentui/react/lib/Icon';
import Login from './custom/Login';
import MicrosoftGraph from './helper/graph';
import React from 'react';
import config from '../lib/config';

class DialogPage extends React.Component {
    constructor(props) {
        super(props);
        fabIcons.initializeIcons();
        this.state = {
            itemId: '01JVL355JYVRKAWYPNWBCLB2GFIMNNFFTK',
            actionItem: undefined,
            showLoginPage: undefined,
            suppliers: [],
            filteredSupplierList: undefined,
            selectedSupplier: undefined,
        };
        this.loginBtn = this.loginBtn.bind(this);
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
                itemId = this.state.itemId;
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
                    actionItem: actionItem,
                });

                const response = await fetchSuppliers();
                this.setState({ suppliers: response.value });
            }
        } catch (error) {
            console.log('fetchData', error);
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
    async onSubmit(actionItem, selectedSuplier) {
        const actionAddData = {
            selectedSuplierCompanyName: selectedSuplier.CompanyName,
            actionItem: actionItem.id,
            actionItemName: actionItem.name,
        };

        this.setState({
            selectedSupplier: undefined,
        });
        setValuesToLocalStorage(actionAddData);
        dialog.url.submit(actionAddData);
    }

    handleRowClick = (supplier) => {
        this.setState({ selectedSupplier: supplier });
    };
    render() {
        return (
            <div className="dialog">
                {this.state.showLoginPage === false && this.state.suppliers && (
                    <form className="dialog_form">
                        <div className="dialog_header">
                            {!this.state.selectedSupplier && (
                                <Text className="dialog_text" style={{ display: 'flex' }}>
                                    Select a Suplier :
                                </Text>
                            )}
                            {this.state.selectedSupplier && (
                                <>
                                    <Text className="dialog_text" style={{ display: 'flex' }}>
                                        {'Selected Supplier :'}
                                        <Text weight="bold" style={{ paddingLeft: '2px' }}>
                                            {this.state.selectedSupplier.CompanyName}
                                        </Text>
                                    </Text>

                                    <Button
                                        appearance="transparent"
                                        icon={<Icon iconName={'Cancel'} />}
                                        onClick={() =>
                                            this.setState({ selectedSupplier: undefined })
                                        }
                                    ></Button>
                                </>
                            )}
                        </div>
                        <div style={{ paddingBottom: '20px' }}>
                            <div className="dialog_list">
                                {this.state.suppliers.length > 0 &&
                                    this.state.suppliers.map((item) => {
                                        return (
                                            <div className="dialog_listitem" key={item.SupplierID}>
                                                <Document16Regular />
                                                <Link
                                                    style={{ fontSize: '12px', paddingLeft: '1px' }}
                                                    onClick={() => this.handleRowClick(item)}
                                                >
                                                    {item.CompanyName}
                                                </Link>
                                            </div>
                                        );
                                    })}
                            </div>
                        </div>
                        {this.state.actionItem && (
                            <div className="dialog_attachment">
                                <Text className="dialog_text">
                                    Attach this document to the supplier:
                                </Text>
                                <div>
                                    <Attachment
                                        actionItem={this.state.actionItem}
                                        appearance={'subtle'}
                                    />
                                </div>
                            </div>
                        )}
                        <div className="dialog_button">
                            <Button
                                type="button"
                                appearance="primary"
                                onClick={() =>
                                    this.onSubmit(
                                        this.state.actionItem,
                                        this.state.selectedSupplier
                                    )
                                }
                            >
                                Add
                            </Button>
                        </div>
                    </form>
                )}
                {this.state.showLoginPage === true && (
                    <Login loginBtnClick={this.loginBtn} />
                )}
            </div>
        );
    }
}
export default DialogPage;