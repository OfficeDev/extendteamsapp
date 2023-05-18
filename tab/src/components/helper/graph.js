
export default class MicrosoftGraph {
    constructor(graphClient, objectId, itemId) {
        this.graphClient = graphClient;
        this.objectId = objectId;
        this.itemId = itemId;
    }

    /**
     * @returns the action information based itemId
     */
    async readActionItem() {
        try {
            return await this.graphClient.api(`/users/${this.objectId}/drive/items/${this.itemId}`).get();
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
            const worksheets = (await this.graphClient.api(`/users/${this.objectId}/drive/items/${this.itemId}/workbook/worksheets`).get()).value;

            //Gets the sheet range based on worksheet name
            const sheetData = (await this.graphClient.api(`/users/${this.objectId}/drive/items/${this.itemId}/workbook/worksheets('${worksheets[0].name}')/usedRange`).get()).values;

            return sheetData;
        } catch (error) {
            console.log("readActionItemData", error);
        }
    }
}