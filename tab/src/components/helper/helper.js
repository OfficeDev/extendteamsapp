export function getIconName(iconType) {
    switch (iconType) {
        case "application/pdf": return "PDF";
        case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": return "ExcelDocument";
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document": return "WordDocument";
        case "application/vnd.openxmlformats-officedocument.presentationml.presentation": return "PowerPointDocument";
        default: return "FileCode";

    }
}

const storageKey = "store";

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

export function isDataExist(value, retrievedData) {
    const val = retrievedData.storeData.find(data => data.selectedSuplierCompanyName === value.selectedSuplierCompanyName);
    if (val) {
        return true;
    }
    return false;
}

export function getAttachmentFromLocalStorage() {
    const retrievedData = localStorage.getItem(storageKey);
    if (retrievedData) {
        return JSON.parse(retrievedData);
    }
    return undefined;
}