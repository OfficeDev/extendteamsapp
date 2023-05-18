import React from "react";
import { app, pages } from "@microsoft/teams-js";
import { Select, Title3, Divider} from "@fluentui/react-components";

import "../index.css";

/**
 * The 'Config' component is used to display your group tabs
 * user configuration options.  Here you will allow the user to
 * make their choices and once they are done you will need to validate
 * their choices and communicate that to Teams to enable the save button.
 */
class TabConfig extends React.Component {
  render() {
    // Initialize the Microsoft Teams SDK
    app.initialize().then(() => {
      /**
       * When the user clicks "Save", save the url for your configured tab.
       * This allows for the addition of query string parameters based on
       * the settings selected by the user.
       */
      const countrySelect = document.getElementById('countrySelect');

     // let selectedCountryId = 0;
      let selectedCountryName = '';
      pages.config.registerOnSaveHandler((saveEvent) => {
        const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
        pages.config
          .setConfig({
            suggestedDisplayName: selectedCountryName,
            entityId: "Test",
            contentUrl: baseUrl + `/index.html#/tab?country=${selectedCountryName}`,
            websiteUrl: baseUrl + `/index.html#/tab?country=${selectedCountryName}`
          })
          .then(() => {
            saveEvent.notifySuccess();
          });
      });
     

      /**
       * After verifying that the settings for your tab are correctly
       * filled in by the user you need to set the state of the dialog
       * to be valid.  This will enable the save button in the configuration
       * dialog.
       */
      const countries = [{id:1,name:"North America"},{id:2,name:"Italy"},{id:3,name:"Australia"},{id:4,name:"France"},{id:4,name:"Japan"}];     
   
      countries.forEach((c) => {    
        const options = document.createElement('option');     
          options.value = c.id;
          options.innerText = c.name;
          countrySelect.appendChild(options);
      });

      // When a category is selected, it's OK to save
      countrySelect.addEventListener('change', (ev) => {
          selectedCountryName = ev.target.options[ev.target.selectedIndex].innerText;
         // selectedCountryId = ev.target.value;
          pages.config.setValidityState(true);
      });
      //
    });


    return (
        <div className="div-config">
          <></>
          <Title3>Please select a country to display local suppliers</Title3>
          <Divider appearance="brand"></Divider>
          <Select id="countrySelect">
              <option disabled="disabled" selected="selected">Select a country</option>
          </Select>
        </div>
    );
  }
}

export default TabConfig;
