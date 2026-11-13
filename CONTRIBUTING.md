# Contributing to ExcelClone

Thank you for your interest in contributing to **ExcelClone**!

ExcelClone is a community-driven, desktop-native spreadsheet application for Windows. We welcome code contributions, bug reports, feature suggestions, and documentation improvements.

---

## 1. Code of Conduct

Please treat all maintainers and contributors with respect, courtesy, and professionalism. Collaborative technical discussions must remain constructive and welcoming to developers of all backgrounds.

---

## 2. Setting Up Your Development Environment

### Prerequisites
- **Operating System:** Windows 10 (version 1809+) or Windows 11
- **SDK:** [.NET 10.0 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- **IDE:** Visual Studio 2022 / 2025 (with `.NET desktop development` workload) or VS Code with C# Dev Kit.
- **Syncfusion License:** Syncfusion Essential Studio components require a license key to suppress trial evaluation watermarks. Eligible individuals and small businesses can obtain a free [Syncfusion Community License](https://www.syncfusion.com/products/communitylicense).

### Cloning and Building
```powershell
# 1. Clone the repository
git clone https://github.com/Sandeep37-s/excel-clone.git
cd excel-clone

# 2. Restore NuGet dependencies
dotnet restore ExcelClone.slnx

# 3. Build the solution
dotnet build ExcelClone.slnx -c Debug

# 4. Run automated tests
dotnet test tests/ExcelClone.Tests/ExcelClone.Tests.csproj
```

---

## 3. Contribution Workflow

1. **Check Existing Issues:** Before starting work, search existing issues and pull requests to ensure someone else isn't already working on it.
2. **Create a Feature Branch:**
   ```powershell
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/your-bug-fix
   ```
3. **Coding Standards:**
   - Follow standard Microsoft C# coding conventions.
   - Keep ViewModels and UI logic decoupled from low-level spreadsheet primitives where possible.
   - Do **NOT** hardcode API keys, secrets, or personal license strings.
4. **Writing Tests:**
   - Any new formula normalization, data processing method, or workbook feature should include automated tests in `tests/ExcelClone.Tests/`.
   - Ensure all automated tests pass before submitting.
5. **Submitting a Pull Request:**
   - Provide a clear PR title and descriptive summary of changes.
   - Link any related issue numbers.

---

## 4. Reporting Bugs

When reporting an issue, please provide:
- A clear, concise title.
- Operating system version and architecture (e.g., Windows 11 Pro 64-bit).
- Detailed steps to reproduce the behavior.
- Expected versus actual behavior.
- If applicable, a sanitized sample `.xlsx` file demonstrating the issue.

