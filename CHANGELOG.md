# CHANGELOG

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to Semantic Versioning (though versions are not explicitly
tracked in this file).

## [Unreleased]

### Added
- Added CHANGELOG.md for tracking security and feature changes

### Changed
- Updated README.md with comprehensive Architecture section documenting the two-check trust model (Google ID token verification in Next.js proxy + shared-secret check in Api.gs::doPost)
- Added secret rotation warning to README.md with instructions for `setupApiSharedSecret()`
- Added Cloudflare Workers AI rationale to README.md explaining data-handling posture for timesheet data
- Reorganized README.md sections for better flow: Architecture, Why Cloudflare AI, Development, Web App deployment (with rotation subsection), Sibling frontend project

### Security
- **Api.gs**: Fixed information leak where "Missing request body" error was thrown before shared secret check; now returns generic "Not authorized" error for all pre-auth failures
- **Api.gs**: Changed "Unknown action: {actionName}" error to generic "Invalid action" to prevent action enumeration
- **Api.gs**: Added JSON parse error handling before auth check to prevent parsing errors from leaking info
- **Notifications.gs**: Added `sanitizeSheetText()` function and applied it to message, type, and link fields when writing to the Notifications sheet to prevent formula injection

### Fixed
- **Api.gs**: Ensured all code paths in `doPost` either authenticate first or return generic errors, preventing information disclosure

## Previous Changes

Prior changes are not tracked in this changelog. This file was created to document changes starting from the security hardening pass.