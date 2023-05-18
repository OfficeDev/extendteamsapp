// https://fluentsite.z22.web.core.windows.net/quick-start
import { FluentProvider, teamsLightTheme, tokens } from "@fluentui/react-components";
import { Navigate, Route, HashRouter as Router, Routes } from "react-router-dom";

import DialogPage from "./Dialog";
import LaunchPage from "./LaunchPage";
import Privacy from "./Privacy";
import React from "react";
import Tab from "./Tab";
import TabConfig from "./TabConfig";
import TermsOfUse from "./TermsOfUse";
import { useTeams } from "@microsoft/teamsfx-react";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export default function App() {
  const { theme } = useTeams({})[0];
  return (
    <FluentProvider
      theme={
        theme || {
          ...teamsLightTheme,
          colorNeutralBackground3: "#eeeeee",
        }
      }
      style={{ background: tokens.colorNeutralBackground3 }}
    >
      <Router>
        <Routes>
          <Route path="/privacy" element={<Privacy />} />
          <Route path="/termsofuse" element={<TermsOfUse />} />
          <Route path="/dialogPage" element={<DialogPage />} />
          <Route path="/launchPage" element={<LaunchPage />} />
          <Route path="/tab" element={<Tab />} />
          <Route path="/tabconfig" element={<TabConfig />} />
          <Route path="*" element={<Navigate to={"/tab"} />}></Route>
        </Routes>
      </Router>
    </FluentProvider>
  );
}
