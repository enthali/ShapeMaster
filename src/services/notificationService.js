class NotificationService {
    showNotification(message) {
        // Create a notification element
        const notification = document.createElement('div');
        notification.className = 'notification';
        notification.innerText = message;

        // Append the notification to the body
        document.body.appendChild(notification);

        // Automatically hide the notification after 3 seconds
        setTimeout(() => {
            this.hideNotification(notification);
        }, 3000);
    }

    hideNotification(notification) {
        // Remove the notification from the DOM
        if (notification) {
            document.body.removeChild(notification);
        }
    }
}

export default NotificationService;