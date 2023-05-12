import './FilteredResult.css';

import { Button, Label, Text } from "@fluentui/react-components";

import Attachment from './Attachment';
import React from "react";

class FilteteredResult extends React.Component {
    constructor(props) {
        super(props);
        this.state = {

        }
    }

    render() {
        console.log("Filtered", this.props);
        return (
            <>
                {this.props.actionId && this.props.actionItem && (
                    <div className='filteredResult'>
                        <div className='filteredHeader'>
                            <Text as="h4" style={{ margin: "10px 0" }} size={500} weight='bold'>FilteteredResult:</Text>
                            <Button appearance='primary'>
                                Clear Filter
                            </Button>
                        </div>
                        <div className='filteredBody'>
                            <div className='filteredAttachment'>
                                {this.props.sheetData && this.props.sheetData.length > 0 &&
                                    <Text className='filteredText'>{`Found ${this.props.sheetData.length} suppliers mentioned in this file`}</Text>
                                }
                                <Attachment actionItem={this.props.actionItem} />
                            </div>
                            <div className="filteredList">
                                {this.props.sheetData && this.props.sheetData.length > 0 &&
                                    <>
                                        <Text as='h4' weight='bold'> List of Supplier Names :</Text>
                                        <div className='filteredListItems'>
                                            {this.props.sheetData.map((item, index) => {
                                                return (
                                                    <div className='filteredListItem' key={item.SupplierId}>
                                                        <Label style={{ fontSize: '12px', paddingLeft: "1px" }}>
                                                            {item.CompanyName}
                                                        </Label>
                                                    </div>
                                                );
                                            })}
                                        </div>
                                    </>
                                }
                            </div>
                        </div>
                    </div>)
                }
            </>
        );

    }
}

export default FilteteredResult;