# Developer Installation Guide for ShapeMaster

This guide provides step-by-step instructions for developers who want to build, test, or contribute to ShapeMaster.

## Prerequisites

- **Windows OS**
- **Visual Studio 2022 Community Edition** (or later)
- **.NET Framework 4.8 Developer Pack**
- **VSTO (Visual Studio Tools for Office) Extensions**
- **Git**

## Setup Steps

1. **Clone the Repository**
   ```pwsh
   git clone https://github.com/enthali/ShapeMaster.git
   cd ShapeMaster
   ```

2. **Open the Solution**
   - Open `ShapeMaster.sln` in Visual Studio.

3. **Configure Code Signing (Optional for Local Builds)**
   - If you want to test ClickOnce deployment, you will need a code-signing certificate.
   - For local development, you can use the provided test certificate or create your own self-signed certificate.
   - See the README or project documentation for instructions on generating and using a self-signed certificate.

4. **Build the Project**
   - Select the desired configuration (Debug or Release).
   - Build the solution (Ctrl+Shift+B).

5. **Run and Debug**
   - Set PowerPoint as the start action if not already set.
   - Press F5 to launch PowerPoint with the add-in loaded.

6. **Testing**
   - Test your changes in PowerPoint.
   - Add or update tests as needed.

7. **Create a Feature Branch**
   - Before making changes, create a new branch:
     ```pwsh
     git checkout -b my-feature-branch
     ```

8. **Commit and Push**
   - Commit your changes and push to your fork.

9. **Open a Pull Request**
   - Submit a pull request to the main repository for review.

## Troubleshooting

- If you encounter issues with dependencies, ensure all prerequisites are installed.
- For ClickOnce or signing issues, refer to the README or open an issue on GitHub.

---

Thank you for contributing to ShapeMaster!
