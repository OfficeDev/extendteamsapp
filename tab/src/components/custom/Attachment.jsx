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
import { getImageIcon } from '../helper/helper';

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
                marginTop: "5px"
            },
            caption: {
                color: tokens.colorNeutralForeground3,
            }
        };
        return (
            <>{this.props.actionItem &&
                <Card style={useStyles.card} orientation="horizontal" appearance={this.props.appearance ? this.props.appearance : undefined}>
                    <CardPreview>
                        <Icon style={{ fontSize: "28px", paddingTop: `${this.props.iconPaddingTop}` }}>
                            <img height={32} width={32} src={getImageIcon(this.props.actionItem.file.mimeType)} alt={this.props.actionItem.name} />
                        </Icon>
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