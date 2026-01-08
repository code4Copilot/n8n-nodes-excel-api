# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> 📖 **[中文版本](CHANGELOG_zh-tw.md)** | **[English Version](CHANGELOG.md)**

## [1.0.2] - 2026-01-08

### Added
- Lookup Column Selection feature: Dynamic dropdown for selecting Excel column headers
- Comprehensive test coverage for Lookup Column selection functionality
- Support for Chinese and special character column names in lookup operations
- Automatic URL encoding for special characters in API requests
- Enhanced user experience with visual column selection interface

## [1.0.1] - 2026-01-06

### Added
- **Process Mode** option for lookup operations in Update and Delete
  - `All Matching Records` (default): Process all rows matching the lookup criteria
  - `First Match Only`: Process only the first matching row
- Batch update support: Update multiple records at once using lookup
- Batch delete support: Delete multiple records at once using lookup
- Performance optimization: Use "First Match Only" mode for unique identifier lookups

### Changed
- Default behavior: Lookup operations now process all matching records by default
- Improved error messages for zero-match scenarios in lookup operations
- Updated documentation with detailed examples for both process modes

### Tests
- Added comprehensive test cases for process mode functionality
- Test coverage for batch update scenarios
- Test coverage for batch delete scenarios
- Validation tests for process mode parameter handling

## [1.0.0] - 2025-12-23

### Added
- Initial release of n8n-nodes-excel-api
- Five core operations: Append, Read, Update, Delete, Batch
- **Object Mode** for Append operation
  - Automatic Excel header reading (first row)
  - Smart column name mapping
  - Ignore unknown columns with warnings
  - Flexible column order
- **Array Mode** for Append operation
  - Exact column order specification
- **Read Operation**
  - Optional cell range specification
  - Automatic header detection and object conversion
- **Update Operation**
  - Update by row number
  - Update by lookup (column value search)
- **Delete Operation**
  - Delete by row number
  - Delete by lookup (column value search)
- **Batch Operation**
  - Execute multiple operations in one API call
  - Mix different operation types
- Integration with [Excel API Server](https://github.com/code4Copilot/excel-api-server)
- Concurrent safety protection via API server file locking
- Support for multi-user simultaneous access
- Comprehensive error handling
- Detailed documentation with usage examples
- TypeScript implementation with full type safety
- ESLint configuration for code quality
- Jest test framework setup
- Gulp build system for icon processing

### Documentation
- Complete README with installation instructions
- Usage examples for common scenarios
- API reference for all operations
- Troubleshooting guide
- Security best practices
- Performance optimization tips

### Developer Experience
- TypeScript support with type definitions
- Automated build process
- Linting and formatting tools
- Test suite with Jest
- GitHub repository setup
- npm package publication ready

## [0.9.0] - 2025-12-20 (Pre-release)

### Added
- Initial development version
- Basic CRUD operations
- Excel API Server integration
- Development and testing infrastructure

---

## Release Notes

### Version 1.0.1 Highlights
This release introduces powerful batch processing capabilities for Update and Delete operations. The new Process Mode option allows you to control whether to process all matching records or just the first match, making it more efficient to handle both unique identifier lookups and batch operations.

**Key Use Cases:**
- **Unique ID Lookups**: Use "First Match Only" mode with Employee ID, Email, etc.
- **Batch Updates**: Use "All Matching Records" mode to update entire departments
- **Batch Deletions**: Use "All Matching Records" mode for data cleanup operations

### Version 1.0.0 Highlights
The first stable release provides a complete solution for managing Excel files in n8n workflows with concurrent safety. The Object Mode feature makes it much easier to work with Excel data by using column names instead of remembering column positions.

**Key Features:**
- **Concurrent Safety**: Built-in file locking prevents data corruption
- **User-Friendly**: Object Mode with automatic header mapping
- **Flexible**: Support for both simple and complex use cases
- **Production Ready**: Comprehensive testing and documentation

---

## Migration Guide

### Upgrading from 1.0.0 to 1.0.1

No breaking changes. All existing workflows will continue to work without modification.

**New Optional Parameter:**
- `processMode` parameter added to Update and Delete operations when using lookup
- Default value is `"all"` to maintain backward compatibility
- To optimize performance for unique identifier lookups, explicitly set to `"first"`

**Example:**
```json
{
  "operation": "update",
  "identifyBy": "lookup",
  "lookupColumn": "Employee ID",
  "lookupValue": "E100",
  "processMode": "first",  // NEW: Add this for better performance
  "valuesToSet": {
    "Salary": "85000"
  }
}
```

---

## Links

- [GitHub Repository](https://github.com/code4Copilot/n8n-nodes-excel-api)
- [npm Package](https://www.npmjs.com/package/n8n-nodes-excel-api)
- [Excel API Server](https://github.com/code4Copilot/excel-api-server)
- [n8n Documentation](https://docs.n8n.io)

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

## Support

- Report bugs: [GitHub Issues](https://github.com/code4Copilot/n8n-nodes-excel-api/issues)
- Ask questions: [n8n Community Forum](https://community.n8n.io)
- Email: hueyan.chen@gmail.com
