// This file contains the JavaScript code for the task pane. It initializes the task pane, handles user interactions, and communicates with the Office JavaScript API.

document.addEventListener("DOMContentLoaded", function () {
    // Initialize the task pane
    Office.onReady(function (info) {
        if (info.host === Office.HostType.PowerPoint) {
            // Task pane is ready
            console.log("Task pane is ready.");
            // Add event listeners or initialize components here
        }
    });

    // Example function to handle a button click
    document.getElementById("exampleButton").addEventListener("click", function () {
        // Handle button click
        console.log("Button clicked!");
        // Call Office API or service methods here
    });
});