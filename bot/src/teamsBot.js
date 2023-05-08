const { TeamsActivityHandler } = require("botbuilder");
const { SupplierME } = require("./messageExtensions/supplierME");
const { NorthwindLinkME } = require("./messageExtensions/northwindLinkME");
class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {

    const queryName = query.parameters[0].name;
    const searchQuery = query.parameters[0].value;

    switch (queryName) {
      case "supplierME":  // Search for suppliers
        return await SupplierME.handleTeamsMessagingExtensionQuery(context, searchQuery);
      default:
        return null;
    }
  }

  async handleTeamsMessagingExtensionSelectItem(context, item) {

    switch (item.queryType) {
      case "supplierME":  // Search for suppliers
        return SupplierME.handleTeamsMessagingExtensionSelectItem(context, item);
      default:
        return null;
    }
  }
  //use https://test.northwindtraders.com?supplierID=3
  async handleTeamsAppBasedLinkQuery(context, query) {
    
    return NorthwindLinkME.linkQuery(context, query);

  }
}

module.exports.TeamsBot = TeamsBot;
