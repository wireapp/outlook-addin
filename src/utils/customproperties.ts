/* global Office, console */

export function setCustomPropertyAsync(item, key, value) {
  return new Promise((resolve, reject) => {
    item.loadCustomPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const customProperties = result.value;

        customProperties.set(key, value);

        customProperties.saveAsync((saveResult) => {
          if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(value);
          } else {
            console.error("Error saving custom properties:", saveResult.error);
            reject(saveResult.error);
          }
        });
      } else {
        console.error("Error loading custom properties:", result.error);
        reject(result.error);
      }
    });
  });
}

export function getCustomPropertyAsync(item, key) {
  return new Promise((resolve, reject) => {
    item.loadCustomPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const customProperties = result.value;

        const value = customProperties.get(key);

        resolve(value);
      } else {
        console.error("Error loading custom properties:", result.error);
        reject(result.error);
      }
    });
  });
}
