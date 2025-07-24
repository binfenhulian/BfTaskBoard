# Contributing to BfTaskBoard

Thank you for your interest in contributing to BfTaskBoard! We welcome contributions from the community.

## How to Contribute

### Reporting Bugs
- Use the GitHub Issues page to report bugs
- Describe the bug in detail
- Include steps to reproduce
- Include expected behavior and actual behavior
- Include system information (Windows version, .NET version)

### Suggesting Features
- Use GitHub Issues with the "enhancement" label
- Clearly describe the feature and its use case
- Explain why this feature would be useful

### Code Contributions

1. **Fork the Repository**
   - Click the "Fork" button on GitHub
   - Clone your fork locally

2. **Create a Branch**
   ```bash
   git checkout -b feature/YourFeatureName
   ```

3. **Make Changes**
   - Write clean, readable code
   - Follow existing code style
   - Add comments where necessary
   - Update documentation if needed

4. **Test Your Changes**
   - Ensure all existing features still work
   - Test your new feature thoroughly
   - Add unit tests if applicable

5. **Commit Guidelines**
   - Use clear, descriptive commit messages
   - Reference issue numbers if applicable
   - Example: `Fix #123: Add validation for empty cells`

6. **Submit Pull Request**
   - Push your branch to your fork
   - Create a pull request from your fork to the main repository
   - Describe your changes in the PR description
   - Link any related issues

## Code Style Guidelines

### C# Coding Standards
- Use PascalCase for class names and method names
- Use camelCase for local variables and parameters
- Use meaningful variable and method names
- Keep methods small and focused
- Use async/await for asynchronous operations

### Comments
- Add XML documentation comments for public methods
- Use inline comments sparingly and only when necessary
- Keep comments up-to-date with code changes

### File Organization
- One class per file
- Organize files into appropriate folders
- Keep related functionality together

## Development Setup

1. Install .NET 6.0 SDK or higher
2. Install Visual Studio 2022 or VS Code
3. Clone the repository
4. Open the solution file in your IDE
5. Build and run the project

## Testing

- Test all new features thoroughly
- Verify no existing functionality is broken
- Test with different data sizes
- Test UI responsiveness
- Test with both themes (light/dark)
- Test with both languages (Chinese/English)

## Documentation

- Update README.md if adding new features
- Add inline documentation for complex logic
- Update user documentation if UI changes

## Questions?

Feel free to open an issue for any questions about contributing.