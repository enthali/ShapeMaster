# Shape Master JS

A PowerPoint add-in written in JavaScript that enhances shape manipulation capabilities, providing users with efficient tools for precise positioning and resizing of shapes within PowerPoint presentations.

## Current Version

**Build #1** (Mai 9, 2025)

## Features

- **Resize Shapes**: Match size, width, or height of multiple shapes to the first selected shape.
- **Swap Positions**: Exchange the positions of two selected shapes with a single click.
- **Bold Text Coloring**: Apply color to bold text within shapes, with color selection options.
- **Notification System**: Display user notifications for actions performed within the add-in.

## Status

ShapeMasterJS is now open source and released under the MIT License. Please report any bugs or feature requests via GitHub issues.

## Installation & Contribution

### For Trusting Users

To install ShapeMasterJS, follow the instructions on the [GitHub Pages project site](https://github.com/yourusername/ShapeMasterJS). After installation, open PowerPoint and check for the "Shape Master" tab in the ribbon.

> **Security Notice:**
>
> If you see a security warning, it is because the add-in is signed with a self-signed certificate. Installing and trusting this certificate is not recommended unless you fully understand the risks and trust the source.
>
> Proceed only if you are comfortable with these risks.

### For Developers

If you want to contribute or build ShapeMasterJS yourself:

- See [InstallForDevelopers.md](InstallForDevelopers.md) for detailed developer setup and build instructions.
- See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines and the pull request workflow.

## Usage

### Resizing Tools
1. Select two or more shapes (the first shape will be the reference).
2. Click the desired resize button (Match Size, Match Width, or Match Height) or use the keyboard shortcuts.

### Swapping Positions Tool
1. Select exactly two shapes.
2. Click the "Swap Positions" button.
3. The shapes will exchange their positions on the slide.

### Bold Text Coloring Tool
1. Select one or more shapes containing text.
2. Use the split button:
   - Main button: Apply the currently selected color to all bold text.
   - Dropdown: Select a new color from the theme colors.

### Notes
1. Use the TODO Note, Review Note, or Comment Note buttons in the Shape Master ribbon group.
2. Each button inserts a snipped-corner rectangle (sticky note style) at the top left of the slide.
3. The note icon on the button matches the note color and displays a letter (T, R, or C) for TODO, Review, or Comment.

### Keyboard Shortcuts
ShapeMasterJS supports keyboard navigation through the Office ribbon interface:

1. Press **Alt** to show keyboard shortcuts.
2. Press **M** to navigate to the ShapeMasterJS tab.
3. Press the keytip for the specific command you want to use.

Common shortcuts:
- Match Size: **Alt, M, MS**
- Match Width: **Alt, M, MW** 
- Match Height: **Alt, M, MH**
- Swap Positions: **Alt, M, SP**
- Apply Color to Bold Text: **Alt, M, BT**
- Insert TODO Note: **Alt, M, TN**
- Insert Review Note: **Alt, M, RN**
- Insert Comment Note: **Alt, M, CN**

## Technical Architecture

### Project Structure
The ShapeMasterJS project is organized as a JavaScript-based add-in targeting PowerPoint. Key components include:

1. **Task Pane**: HTML, CSS, and JavaScript files that create the user interface for the add-in.
2. **Service Classes**: Specialized classes for different functionality areas, located in the Services directory.
3. **Command Handlers**: Functions that define the commands available in the add-in.
4. **Ribbon UI**: Custom UI defined in JavaScript that extends PowerPoint's ribbon.

### Key Files
- **Services Directory**: Contains service classes for different functionality groups.
  - **shapePositioningService.js**: Handles shape position operations.
  - **shapeResizingService.js**: Handles shape resizing operations.
  - **textFormattingService.js**: Handles text formatting operations.
  - **notificationService.js**: Centralizes user notifications.
- **commands.js**: Defines the command handlers for the add-in.
- **ribbon.js**: Initializes the custom ribbon UI for the add-in.
- **taskpane.html**: The HTML structure for the task pane.
- **taskpane.js**: The JavaScript code for the task pane.
- **taskpane.css**: The CSS styles for the task pane.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Privacy

ShapeMasterJS does not collect, transmit, or store any personal data or user information. All operations are performed locally within your PowerPoint environment.
