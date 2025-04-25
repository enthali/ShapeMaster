# Shape Master

A PowerPoint add-in that enhances shape manipulation capabilities, providing users with efficient tools for precise positioning and resizing of shapes within PowerPoint presentations.

## Features

- **Resize Shapes**: Match size, width, or height of multiple shapes to the first selected shape
- **Swap Positions**: Exchange the positions of two selected shapes with a single click
- **Bold Text Coloring** (Planned): Apply color to bold text within shapes, with color selection options
- **Notes Buttons**: Quickly insert color-coded TODO, Review, or Comment notes as shapes on your slide. Each note type uses a distinct color and icon, defined in the ribbon XML.

## Status

ShapeMaster is now open source and released under the MIT License. Please report any bugs or feature requests via GitHub issues.

## Installation & Contribution

### For Trusting Users

To install ShapeMaster, follow the instructions on the [GitHub Pages project site](https://enthali.github.io/ShapeMaster). After installation, open PowerPoint and check for the "Shape Master" tab in the ribbon.

>**Security Notice:**
>
>If you see a security warning, it is because the add-in is signed with a self-signed certificate. Installing and trusting this certificate is not recommended unless you fully understand the risks and trust the source. 
>
>Proceed only if you are comfortable with these risks.

### For Developers

If you want to contribute or build ShapeMaster yourself:

- See [InstallForDevelopers.md](InstallForDevelopers.md) for detailed developer setup and build instructions.
- See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines and the pull request workflow.


## Usage

### Resizing Tools
1. Select two or more shapes (the first shape will be the reference)
2. Click the desired resize button (Match Size, Match Width, or Match Height)

### Swapping Positions Tool
1. Select exactly two shapes
2. Click the "Swap Positions" button
3. The shapes will exchange their positions on the slide

### Bold Text Coloring Tool
1. Select one or more shapes containing text
2. Use the split button:
   - Main button: Apply the currently selected color to all bold text
   - Dropdown: Select a new color from the theme colors

### Notes
1. Use the TODO Note, Review Note, or Comment Note buttons in the Shape Master ribbon group.
2. Each button inserts a snipped-corner rectangle (sticky note style) at the top left of the slide
3. The note icon on the button matches the note color and displays a letter (T, R, or C) for TODO, Review, or Comment.

## Technical Architecture

### Project Structure
The ShapeMaster project is organized as a VSTO (Visual Studio Tools for Office) add-in written in C# targeting PowerPoint. Key components include:

1. **Ribbon Interface**: Custom UI defined in XML and implemented in C# that extends PowerPoint's ribbon
2. **Service Classes**: Specialized classes for different functionality areas, located in the Services directory
3. **Core Logic**: Coordination in ThisAddIn.cs which delegates to appropriate service classes
4. **Notification System**: Non-modal tooltip notifications for user feedback

### Key Files
- **Services Directory**: Contains service classes for different functionality groups
  - **ShapePositioningService.cs**: Handles shape position operations
  - **TextFormattingService.cs**: Handles text formatting operations
  - **ShapeResizingService.cs** : Handles shape resizing operations
  - **NotificationService.cs**: Centralizes user notifications
  - **RibbonUIService.cs**: Manages ribbon UI interactions, including dynamic note button icons colored and labeled according to the XML tag
  - **ServiceManager.cs**: Coordinates service dependencies
- **ShapeMasterRibbon.cs**: Defines the ribbon UI callbacks and implements the IRibbonExtensibility interface
- **ShapeMasterRibbon.xml**: XML definition of the custom ribbon UI elements
- **ThisAddIn.cs**: Initialization and coordination between services


## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Privacy

ShapeMaster does not collect, transmit, or store any personal data or user information. All operations are performed locally within your PowerPoint environment.