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
import { getIconName } from '../helper/helper';

class Attachment extends React.Component {
    constructor(props) {
        super(props);
        fabIcons.initializeIcons();
    }
    render() {
        console.log(this.props);
        const useStyles = {
            card: {
                maxWidth: "100%",
            },
            caption: {
                color: tokens.colorNeutralForeground3,
            }
        };
        return (
            <>{this.props.actionItem &&
                <Card style={useStyles.card} orientation="horizontal" appearance={this.props.appearance ? this.props.appearance : undefined}>
                    <CardPreview>
                        <Icon iconName={getIconName(this.props.actionItem.file.mimeType)} style={{ fontSize: "28px", paddingTop: `${this.props.iconPaddingTop}` }} />
                    </CardPreview>
                    <CardHeader
                        header={<Text weight="semibold">{this.props.actionItem.name}</Text>}
                        description={
                            <Caption1 style={useStyles.caption}></Caption1>
                        }
                    />
                </Card>
            }</>
        );
    }
}

export default Attachment;