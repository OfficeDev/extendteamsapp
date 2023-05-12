import * as fabIcons from '@uifabric/icons';

import {
    Caption1,
    Card,
    CardHeader,
    CardPreview,
    Text,
    tokens
} from "@fluentui/react-components";

import { Icon } from '@fluentui/react/lib/Icon';
import React from "react";

class Attachment extends React.Component {
    constructor(props) {
        super(props);
        fabIcons.initializeIcons();
    }
    render() {
        console.log(this.props);
        const useStyles = {
            card: {
                width: "360px",
                maxWidth: "100%",
                height: "fit-content",
                marginBottom: "10px"
            },
            caption: {
                color: tokens.colorNeutralForeground3,
            }
        };
        return (
            <>{this.props.actionItem &&
                <Card style={useStyles.card} orientation="horizontal" appearance={this.props.appearance ? this.props.appearance : undefined}>
                    <CardPreview style={{ marginLeft: 'unset' }}>
                        <Icon iconName={"ExcelDocument"} style={{ fontSize: "28px", paddingTop: "10px" }} />
                    </CardPreview>
                    <CardHeader
                        header={<Text weight="semibold">{this.props.actionItem.name}</Text>}
                        description={
                            <Caption1 style={useStyles.caption}>Path: {this.props.actionItem.parentReference.path}</Caption1>
                        }
                    />
                </Card>
            }</>
        );
    }
}

export default Attachment;