import { Button, Text } from "@fluentui/react-components";

import React from "react";

class Login extends React.Component {
    render() {
        return (
            <div className="auth">
                <Text as="h1" size={700} >Welcome Northwind App!</Text>
                <Button appearance="primary" onClick={() => this.props.loginBtnClick()}>
                    Login
                </Button>
            </div>
        );
    }
}

export default Login;