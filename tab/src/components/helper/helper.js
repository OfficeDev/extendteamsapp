const storageKey = "store";

/**
 * @param mimeType
 * Returns the icon name based on file's mimeType
*/
export function getIconName(mimeType) {
    switch (mimeType) {
        case "application/pdf": return "PDF";
        case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": return "ExcelDocument";
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document": return "WordDocument";
        case "application/vnd.openxmlformats-officedocument.presentationml.presentation": return "PowerPointDocument";
        default: return "FileCode";
    }
}

/**
 * @param value
 * Sets the data in the localStorage
*/
export function setValuesToLocalStorage(value) {
    const storeData = [];
    if (value) {
        const retrievedData = localStorage.getItem(storageKey);
        if (retrievedData) {
            const parsedData = JSON.parse(retrievedData);
            if (!isDataExist(value, parsedData)) {
                parsedData.storeData.push(value);
                const jsonString = JSON.stringify(parsedData);
                localStorage.setItem(storageKey, jsonString)
            }
        }
        else {
            storeData.push(value);
            const storeAll = {
                storeData: storeData
            }
            const jsonString = JSON.stringify(storeAll);
            localStorage.setItem(storageKey, jsonString)
        }
    }
}
/**
 * @param value
 * @param retrievedData
 * checks and returns boolean if value exist in retrieved data from localStorage
*/
export function isDataExist(value, retrievedData) {
    const val = retrievedData.storeData.find(data => data.selectedSuplierCompanyName === value.selectedSuplierCompanyName);
    if (val) {
        return true;
    }
    return false;
}
/**
 * Returns attachment's info from localStorage
*/
export function getAttachmentFromLocalStorage() {
    const retrievedData = localStorage.getItem(storageKey);
    if (retrievedData) {
        return JSON.parse(retrievedData);
    }
    return undefined;
}

/**
 * Returns response from the api
*/
export async function fetchSuppliers() {
    const url = "https://services.odata.org/V4/Northwind/Northwind.svc/Suppliers"
    const response = await fetch(url);
    return await response.json()
}

/**
 * @returns filtered SupplierList With Attachment
 */
export function filteredSupplierListWithAttachment(filteredSupplierList) {
    const attachments = getAttachmentFromLocalStorage();

    return filteredSupplierList.filter(x => x !== undefined).map(y => {
        const attachmentName = attachments && attachments.storeData.find(item => item.selectedSuplierCompanyName === y.CompanyName);
        return {
            CompanyName: y.CompanyName,
            ContactName: y.ContactName,
            Phone: y.Phone,
            Country: y.Country,
            Attachments: attachmentName ? attachmentName.actionItemName : ""
        }
    });
}

/**
 * Login
 */
export async function loginBtnClick(credential, scope) {
    try {
        // Popup login page to get user's access token
        await credential.login(scope);
    } catch (err) {
        console.log(err);
        if (err instanceof Error && err.message?.includes("CancelledByUser")) {
            const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
            err.message +=
                '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
                "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
                `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}`;
        }

        alert("Login failed: " + err);
        return;
    }
}