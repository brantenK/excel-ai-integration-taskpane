// Office Add-in Commands
// This file contains the command functions for the Excel AI Integration add-in

/**
 * Shows a notification when the add-in command is executed.
 * @param event The event object from the add-in command
 */
function action(event: Office.AddinCommands.Event): void {
    const message: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Excel AI Integration add-in command executed successfully!",
        icon: "Icon.80x80",
        persistent: true
    };

    // Show the message
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
    
    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

// Register the function with Office
Office.actions.associate("action", action);