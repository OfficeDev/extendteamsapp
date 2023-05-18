import './FilteredResult.css';

import { Button, Text } from "@fluentui/react-components";

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
                {this.props.itemId && this.props.actionItem && (
                    <div className='filteredResult'>
                        <div className='filteredHeader'>
                            <Text as="h4" style={{ margin: "10px 0" }} size={500} weight='bold'>Filtered Result:</Text>
                            <Button className="filteredHeaderBtn" appearance='primary' onClick={this.props.clearFilter}>
                                Clear Filter
                            </Button>
                        </div>
                        <div className='filteredBody'>
                            <div className='filteredAttachment'>
                                {this.props.sheetData && this.props.sheetData.length > 0 &&
                                    <Text className='filteredText'>{`Found ${this.props.sheetData.length} suppliers mentioned in this file`}</Text>
                                }
                                <Attachment actionItem={this.props.actionItem} appearance="subtle" />
                            </div>
                        </div>
                    </div>)
                }
            </>
        );

    }
}

export default FilteteredResult;