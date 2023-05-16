
export default class MicrosoftGraph {
    constructor(graphClient, objectId, actionId) {
        this.graphClient = graphClient;
        this.objectId = objectId;
        this.actionId = actionId;
    }

    /**
     * @returns the action information based actionId
     */
    async readActionItem() {
        try {
            return await this.graphClient.api(`/users/${this.objectId}/drive/items/${this.actionId}`).get();
        } catch (error) {
            console.log("readActionItem", error);
        }
    }

    /**
     * @returns the excel's sheet data
     */
    async readActionItemData() {
        try {
            //Gets the excel worksheets
            const worksheets = (await this.graphClient.api(`/users/${this.objectId}/drive/items/${this.actionId}/workbook/worksheets`).get()).value;

            //Gets the sheet range based on worksheet name
            const sheetData = (await this.graphClient.api(`/users/${this.objectId}/drive/items/${this.actionId}/workbook/worksheets('${worksheets[0].name}')/usedRange`).get()).values;

            return sheetData;
        } catch (error) {
            console.log("readActionItemData", error);
        }
    }
}