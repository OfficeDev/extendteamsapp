import { DefaultButton, DetailsList, SelectionMode } from '@fluentui/react';
import { Link, Text } from '@fluentui/react-components'

import React from 'react';

class Suppliers extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            selectedSupplier: null
        };
    }
    handleRowClick = (supplier) => {
        this.setState({ selectedSupplier: supplier });
    };

    render() {
        const supplierColumns = [
            {
                key: 'companyName',
                name: 'Name',
                fieldName: 'CompanyName',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true,
                onRender: (item) => {
                    return (
                        <Link key={item.id} style={{ fontSize: '12px' }} onClick={() => this.handleRowClick(item)}>
                            {item.CompanyName}
                        </Link>
                    );
                }
            },
            {
                key: 'contactName',
                name: 'Contact',
                fieldName: 'ContactName',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            },
            {
                key: 'phone',
                name: 'Phone',
                fieldName: 'Phone',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            },
            {
                key: 'country',
                name: 'Country',
                fieldName: 'Country',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            },
            {
                key: 'attachments',
                name: 'Attachments',
                fieldName: 'Attachments',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            }
        ];
        const selectSuppliercolumn = [
            {
                key: 'companyName',
                name: 'Name',
                fieldName: 'CompanyName',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            },
            {
                key: 'contactName',
                name: 'Contact',
                fieldName: 'ContactName',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            },
            {
                key: 'phone',
                name: 'Phone',
                fieldName: 'Phone',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            },
            {
                key: 'country',
                name: 'Country',
                fieldName: 'Country',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            },
            {
                key: 'attachments',
                name: 'Attachments',
                fieldName: 'Attachments',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true
            }
        ];
        return (<div>
            {!this.state.selectedSupplier && this.props.suppliers && this.props.suppliers.length > 0 && (
                <div>
                    <div className='headingSupplier'>
                        <Text size={500} as="h2" style={{ margin: "15px" }}>Suppliers</Text>
                    </div>
                    <DetailsList
                        items={this.props.suppliers}
                        columns={supplierColumns}
                        selectionMode={SelectionMode.single}
                        onItemInvoked={this.handleRowClick}
                    />
                </div>)}
            {this.state.selectedSupplier && (
                <div>
                    <DetailsList
                        items={[this.state.selectedSupplier]}
                        columns={selectSuppliercolumn}
                        selectionMode={SelectionMode.none}
                    />
                    <DefaultButton key={""} onClick={() => this.handleRowClick(null)}>
                        Back to Suppliers
                    </DefaultButton>
                </div>
            )}
        </div>)
    }
}
export default Suppliers;