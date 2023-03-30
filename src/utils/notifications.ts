export function showNotification(key, message, type, icon) {
    const notificationMessage = {
      type,
      icon,
      message,
      persistent: false,
    };
  
    Office.context.mailbox.item.notificationMessages.addAsync(key, notificationMessage, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`Notification with key "${key}" added successfully.`);
      } else {
        console.error('Failed to add notification message:', result.error);
      }
    });
}

function removeNotification(key) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`Notification with key "${key}" removed successfully.`);
      } else {
        console.error('Failed to remove notification:', result.error);
      }
    });
}