# logon_logger_rs

[![Build Status](https://img.shields.io/github/actions/workflow/status/MrDwarf7/logon_logger_rs/build.yml?branch=main)](https://github.com/MrDwarf7/logon_logger_rs/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Rust Version](https://img.shields.io/badge/rust-nightly-orange.svg)](rust-toolchain.toml)

> An idiomatic Rust port of the classic PowerShell logon logging script – async, composable, and trait-driven.

> [!IMPORTANT]
> **WORK IN PROGRESS** – This project is currently in early development and was thrown together extremely quickly. As such, you may encounter some... unconventional... design choices. The codebase is functional but will undergo significant refactoring as it matures. Use in production environments at your own discretion, and expect breaking changes in future updates.

## Overview

**logon_logger_rs** is a modern, high-performance replacement for legacy PowerShell-based logon logging scripts commonly used by IT administrators in educational environments. Originally designed for high school IT departments, this tool tracks user workstation logons by gathering comprehensive system and user information, then logging it to organized Excel files on a network share.

### Why Rust?

The original PowerShell script, while functional, suffered from:
- Slow execution times during user logon
- Blocking I/O operations that delayed the desktop
- Limited error handling and recovery
- Difficult to maintain and extend

This Rust implementation provides:
- **⚡ Blazing fast execution** via async/await with Tokio
- **🛡️ Memory safety** with zero-cost abstractions
- **📊 Reliable Excel handling** with proper formatting and error recovery
- **🔧 Easy maintenance** through idiomatic, modular code

## Features

### Core Capabilities

| Feature | Description |
|---------|-------------|
| **Async Data Collection** | Concurrently gathers computer info, user details, hardware specs, and OS data |
| **Active Directory Integration** | Queries AD for Organizational Unit (OU) information for both users and computers |
| **School Period Tracking** | Automatically determines the current school period based on time of day |
| **Dual Log Format** | Creates both workstation-centric and user-centric log files |
| **Excel Output** | Generates properly formatted `.xlsx` files with tables, filters, and auto-sizing |
| **Safe & Secure** | Forbids unsafe code; no memory vulnerabilities |

### Data Collected

The following information is gathered during each logon event:

#### User Information
- **Username** – Active Directory username
- **User OU** – Organizational Unit the user belongs to
- **Date & Time** – Precise timestamp of the logon event
- **School Period** – Current period (Before School, Form, Period 1-4, Morning Tea, Second Lunch, After Hours)

#### Workstation Information
- **Computer Name** – Network identifier of the workstation
- **WS OU** – Workstation's Organizational Unit
- **Full OU** – Complete OU path
- **Description** – System description from WMI

#### Hardware Information
- **Make** – Computer manufacturer (e.g., Dell, HP, Lenovo)
- **Model** – Specific model identifier
- **UUID** – Universally Unique Identifier
- **Serial Number** – BIOS serial number

#### Operating System
- **OS Name** – Windows version name
- **OS Version** – Display version (e.g., 22H2, 23H2)

## Architecture

The project follows a modular, trait-driven architecture:

```
src/
├── main.rs           # Entry point – orchestrates data collection and logging
├── collect.rs        # Data gathering: base info, hardware (WMI), OS (registry)
├── workstation.rs    # WorkStationEntry struct with all collected data
├── user_entry.rs     # Wrapper for user-centric logging
├── period.rs         # School period definitions and time-based lookup
├── append.rs         # Excel file creation, appending, and formatting
├── executor.rs       # PowerShell command executor for AD queries
├── error.rs          # Custom error types with thiserror
└── prelude.rs        # Common imports and utilities
```

### Key Traits

| Trait | Purpose |
|-------|---------|
| `ExcelLoggable` | Defines how entries are written to and parsed from Excel |
| `HasDateTime` | Provides datetime access for sorting and filtering |
| `FieldLengths` | Enables dynamic column width calculation |

### Data Flow

1. **Logon Trigger** → Binary executes via Group Policy or scheduled task
2. **Parallel Collection** → Three async tasks gather base, hardware, and OS info concurrently
3. **Entry Construction** → Data is composed into `WorkStationEntry` and `UserEntry`
4. **Excel Logging** → Entries are appended to daily log files on the network share
5. **Completion** → Process exits cleanly (typically sub-second execution)

## Installation

### Prerequisites

- **Windows 10/11** or **Windows Server 2016+**
- **Active Directory** environment
- **Network share** with appropriate permissions
- **Rust nightly toolchain** (for building from source)

### Building from Source

1. **Install Rust** (nightly required):
   ```powershell
   rustup toolchain install nightly
   rustup default nightly
   ```

2. **Clone the repository**:
   ```powershell
   git clone https://github.com/MrDwarf7/logon_logger_rs.git
   cd logon_logger_rs
   ```

3. **Build with cargo-make**:
   ```powershell
   # Install cargo-make if not already installed
   cargo install cargo-make

   # Build release binary
   makers br
   # or
   cargo make build_release
   ```

   The binary will be available at `target/release/logon_logger.exe`

### Alternative: Direct Cargo Build

```powershell
cargo build --release
```

## Usage

### Deployment in a School Environment

#### 1. Prepare the Network Share

Create a network share accessible by all domain computers:

```
\\Server\LogonLogger$\Logs\
├── ComputerNEW\     # Workstation logs
└── UserNEW\         # User logs
```

**Permissions Required:**
- **Domain Computers**: Write access to log directories
- **IT Administrators**: Full control for analysis and maintenance

#### 2. Configure the Binary

Update the paths in `src/main.rs` (lines 44-46) if your network share differs:

```rust
const WS_BASE_PATH: &str = r"\\Server\LogonLogger$\Logs\ComputerNEW";
const USER_BASE_PATH: &str = r"\\Server\LogonLogger$\Logs\UserNEW";
```

Then rebuild:
```powershell
makers br
```

#### 3. Deploy via Group Policy

**Option A: Logon Script (Recommended)**

1. Open **Group Policy Management**
2. Create or edit a GPO linked to your student/staff OU
3. Navigate to: `User Configuration > Policies > Windows Settings > Scripts > Logon`
4. Add `logon_logger.exe` as a logon script

**Option B: Scheduled Task**

Create a scheduled task that runs at user logon:
- **Trigger**: At logon
- **Action**: Start a program → `logon_logger.exe`
- **Conditions**: Run only if network connection is available

#### 4. Verify Deployment

After a user logs on:
1. Check `\\Server\LogonLogger$\Logs\ComputerNEW\` for `workstation_log_YYYY-MM-DD.xlsx`
2. Check `\\Server\LogonLogger$\Logs\UserNEW\` for `user_log_YYYY-MM-DD.xlsx`
3. Open Excel files to verify data is being captured correctly

### Log File Format

Each Excel file contains:
- **Header Row** – Column names with bold formatting
- **Data Rows** – One row per logon event
- **Table Formatting** – Professional styling with filters enabled
- **Sorted Data** – Most recent logons appear first (sorted by datetime)
- **Auto-sized Columns** – Content fits without manual resizing

## Configuration

### School Periods

School periods are defined in `src/period.rs` and can be customized:

```rust
pub const PERIODS: [TimePeriod; 9] = [
    TimePeriod::new(hms(5, 0), hms(8, 45), false, "Before School"),
    TimePeriod::new(hms(8, 45), hms(8, 55), false, "Form"),
    TimePeriod::new(hms(8, 55), hms(10, 5), false, "Period 1"),
    // ... add or modify periods as needed
];
```

Modify times and labels to match your school's schedule, then rebuild.

### Network Paths

As mentioned in [Usage](#usage), update `WS_BASE_PATH` and `USER_BASE_PATH` in `main.rs` to match your infrastructure.

## Development

### Build Tasks (via cargo-make)

| Command | Description |
|---------|-------------|
| `makers b` or `cargo make build` | Build debug version |
| `makers br` | Build release version |
| `makers t` | Run tests |
| `makers f` | Format code |
| `makers cp` | Run Clippy lints |
| `makers r` | Run debug build |
| `makers rr` | Run release build |

### CI/CD

GitHub Actions workflows are configured for:
- **Build** – Compiles on Windows runners
- **Test** – Runs the test suite
- **Format Check** – Ensures code formatting compliance

See `.github/workflows/` for details.

## Technical Details

### Dependencies

| Crate | Purpose |
|-------|---------|
| `tokio` | Async runtime (features: full) |
| `chrono` | DateTime handling |
| `calamine` | Reading existing Excel files |
| `rust_xlsxwriter` | Creating and writing Excel files |
| `wmi` | Windows Management Instrumentation queries |
| `serde` | Serialization/deserialization |
| `eyre` | Error reporting |
| `thiserror` | Custom error types |
| `winreg` | Windows Registry access |
| `tracing` | Structured logging |

### Build Configuration

- **Edition**: Rust 2024
- **Toolchain**: Nightly (required for cranelift backend)
- **Codegen Backend**: Cranelift (faster compilation)
- **Unsafe Code**: Forbidden (enforced by lint)

### Performance

Typical execution time: **< 500ms** on modern hardware, including:
- WMI queries for hardware info
- Registry reads for OS info
- PowerShell AD queries
- Excel file I/O on network share

## Troubleshooting

### Common Issues

**"Failed to get DN" Error**
- Ensure the computer and user have Active Directory read permissions
- Verify PowerShell execution policy allows scripts

**Network Share Access Denied**
- Confirm domain computer accounts have write access to the share
- Check NTFS and share permissions

**Excel Files Corrupted**
- Ensure no other processes are locking the files
- Verify sufficient disk space on the server

## Contributing

Contributions are welcome! Please follow these guidelines:

1. **Fork** the repository
2. **Create a feature branch**: `git checkout -b feature/amazing-feature`
3. **Commit changes**: Use [Conventional Commits](https://www.conventionalcommits.org/)
   - Examples: `feat: add GPU info collection`, `fix: handle unicode usernames`
4. **Push to branch**: `git push origin feature/amazing-feature`
5. **Open a Pull Request**

### Development Setup

```powershell
# Install development tools
rustup component add rustfmt clippy
cargo install cargo-make

# Run full check suite
makers a  # Runs build, test, format, clippy
```

### Code Style

- Follow idiomatic Rust conventions
- Use `rustfmt` for formatting
- Address all Clippy warnings
- Add tests for new functionality
- Document public APIs with doc comments

## License

This project is licensed under the **MIT License** – see the [LICENSE](LICENSE) file for details.

```
Copyright (c) 2025 MrDwarf7

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
```

## Acknowledgments

- Original PowerShell script authors in the educational IT community
- The Rust async working group for Tokio
- Contributors to the `rust_xlsxwriter` and `calamine` crates

---

**Made with ❤️ for IT administrators everywhere**
