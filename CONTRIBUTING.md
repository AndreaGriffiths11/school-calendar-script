# Contributing to School Email to Calendar Integration

Thank you for your interest in contributing to this project! This Google Apps Script helps automate the process of extracting events from school emails and adding them to Google Calendar.

## How to Contribute

### Reporting Issues

Before creating an issue, please:

1. **Search existing issues** to avoid duplicates
2. **Check the troubleshooting section** in the README.md
3. **Provide detailed information** including:
   - Steps to reproduce the problem
   - Expected vs actual behavior
   - Relevant log output from Google Apps Script
   - Your configuration (with sensitive information removed)

### Suggesting Enhancements

We welcome suggestions for new features or improvements! Please:

1. **Check if the feature already exists** or is planned
2. **Create an issue** describing:
   - The problem you're trying to solve
   - Your proposed solution
   - Why this would be valuable to other users
   - Any alternative approaches you've considered

### Code Contributions

#### Getting Started

1. **Fork the repository** on GitHub
2. **Create a new branch** from `main` for your changes
3. **Set up your development environment**:
   - Go to [Google Apps Script](https://script.google.com)
   - Create a new project
   - Copy the code from `school-calendar-script.js`
   - Configure the variables for testing

#### Development Guidelines

**Code Style:**
- Use clear, descriptive variable and function names
- Add JSDoc comments for all functions
- Follow existing code formatting and style
- Use `const` for constants and `let` for variables
- Include error handling in new functions

**Testing Your Changes:**
- Test with the `testCalendarAccess` function first
- Run `processSchoolEmails` manually to verify functionality
- Test edge cases (emails without dates, malformed dates, etc.)
- Verify that your changes don't break existing functionality

**Documentation:**
- Update JSDoc comments for any modified functions
- Update README.md if your changes affect setup or usage
- Include examples in your code comments where helpful

#### Types of Contributions Welcome

**Bug Fixes:**
- Date parsing improvements
- Calendar integration issues
- Email processing edge cases
- Performance optimizations

**Feature Enhancements:**
- Additional date format support
- Better event description extraction
- New configuration options
- Improved error handling and logging

**Documentation:**
- Code comments and JSDoc improvements
- README.md enhancements
- Setup instruction clarifications
- Troubleshooting guide additions

#### Pull Request Process

1. **Create a focused PR** that addresses a single issue or feature
2. **Write a clear title** that summarizes your changes
3. **Provide a detailed description** including:
   - What changes you made and why
   - How to test the changes
   - Any potential breaking changes
   - Screenshots/logs if applicable

4. **Ensure your code:**
   - Follows the existing code style
   - Includes appropriate error handling
   - Has been tested manually in Google Apps Script
   - Doesn't break existing functionality

5. **Update documentation** as needed
6. **Be responsive** to feedback and requested changes

#### Code Review

All submissions require review. We'll look for:

- **Functionality:** Does the code work as intended?
- **Code Quality:** Is it readable, maintainable, and well-structured?
- **Documentation:** Are changes properly documented?
- **Testing:** Has it been adequately tested?
- **Compatibility:** Does it work with different email formats and calendar setups?

### Development Setup

Since this is a Google Apps Script project, the development process is unique:

1. **Create a test Google Apps Script project**
2. **Set up a test calendar** for development
3. **Use test email data** or create test emails
4. **Configure logging** to debug your changes:
   ```javascript
   Logger.log("Debug message");
   ```
5. **Access logs** via the Apps Script editor execution log

### Communication

- **Be respectful** and constructive in all interactions
- **Ask questions** if anything is unclear
- **Provide context** when discussing issues or changes
- **Follow up** on your contributions

### License

By contributing to this project, you agree that your contributions will be licensed under the MIT License.

## Getting Help

- Check the [README.md](README.md) for setup and usage instructions
- Review existing [issues](../../issues) for common problems
- Create a new issue if you need help with your contribution

Thank you for helping make this project better! 🎉