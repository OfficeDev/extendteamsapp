const config = {
    initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
    clientId: process.env.REACT_APP_CLIENT_ID,
    apiEndpoint: process.env.REACT_APP_FUNC_ENDPOINT,
    scopes: [
        "https://graph.microsoft.com/User.Read",
        "https://graph.microsoft.com/Files.Read"
    ]
}

export default config;