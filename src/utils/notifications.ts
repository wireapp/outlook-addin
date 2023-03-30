export function showNotification(notificationId, message, type, icon) {
    const notificationMessage = {
      type,
      icon,
      message,
      persistent: false,
    };
  
    Office.context.mailbox.item.notificationMessages.addAsync(notificationId, message, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log('Notification message added successfully.');
      } else {
        console.error('Failed to add notification message:', result.error);
      }
    });
}