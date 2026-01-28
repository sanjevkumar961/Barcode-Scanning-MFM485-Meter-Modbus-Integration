# Barcode-Scanning-MFM485-Meter-Modbus-Integration

> **MotorNoLoadTester v1.2.0** - Industrial-grade motor testing and data logging system

This project integrates barcode scanning with Modbus RTU communication to test and record data from MFM485 meters. It automates motor testing workflows by reading electrical parameters (Voltage, Frequency, Power Factor, Watts, Amps) and logging results to Excel with comprehensive validation and reporting.

## 🎯 Features

- **📱 Barcode Scanning** - Automated serial number capture with pattern validation
- **⚡ Modbus Communication** - Direct RTU protocol connection to MFM485 meters with configurable retry logic
- **📊 Excel Logging** - Automatic workbook and sheet creation with real-time data logging
- **🔧 Dynamic Configuration** - Interactive setup wizard with config persistence in Excel
- **🔄 Retry Mechanism** - Configurable retry attempts for failed register reads (default: 5 retries)
- **⏹️ Graceful Shutdown** - ESC key for safe program termination with data sorting
- **🎨 Rich Terminal Output** - Color-coded messages and formatted data tables
- **✅ Quality Assurance** - Built-in amp range validation with pass/fail reporting

## 📋 Requirements

| Component | Version |
|-----------|---------|
| Python | 3.10+ (with type hints support) |
| pymodbus | Latest |
| openpyxl | Latest |
| keyboard | Latest |
| colorama | Latest |
| pandas | Latest (for Excel sorting) |
| prettytable | Latest |
| requests | Latest (for auto-updates) |

### Quick Install

```bash
pip install pymodbus openpyxl keyboard colorama pandas prettytable requests
```


## 🚀 Quick Start

### 1. Setup Hardware
- Connect MFM485 meter via Modbus RTU serial connection (typically COM3-COM5)
- Verify USB/serial adapter is recognized by Windows

### 2. Install & Configure
```bash
git clone https://github.com/sanjevkumar961/Barcode-Scanning-MFM485-Meter-Modbus-Integration
cd Barcode-Scanning-MFM485-Meter-Modbus-Integration
pip install -r requirements.txt
python MotorNoLoadTester_1.2.0.py
```

### 3. Initial Setup Wizard
On first run, you'll be prompted for:
- **Method**: `rtu` (default)
- **Port**: `COM3` (or your serial port)
- **Baudrate**: `9600` (typical)
- **Parity**: `E` (even)
- **Stopbits**: `1`
- **Bytesize**: `8`
- **Timeout**: `1` (second)

These settings are saved to the Excel file's Config sheet for future runs.

## 📖 Detailed Usage

### File Structure
```
MotorNoLoadReport.xlsx (created on Desktop)
├── Config          # Modbus settings
├── ModelAmp        # Motor models & amp ratings
├── Users           # Authorized fitters (encrypted)
└── [Serial]        # Data sheets (one per motor model)
```

### Workflow

1. **Start Program** → `python MotorNoLoadTester_1.2.0.py`
2. **Authenticate** → Scan/enter fitter barcode (encrypted)
3. **Scan Motor** → Present motor serial number barcode
4. **Validation** → System verifies serial pattern: `[A-Z]+[0-9]+[A-Z]?-[A-Z][0-9]{2}/[0-9]{3,4}`
5. **Model Lookup** → Retrieves min/max amp specs (or prompts to add new)
6. **Read Meter** → Acquires: Voltage, Frequency, PF, Watts, Amps
7. **Test Result** → Pass (green) if amps within range, Fail otherwise
8. **Log Data** → Writes complete record to Excel with timestamp
9. **Repeat** → Press ESC to exit (triggers auto-sorting of data)
## ⚙️ Configuration Details

### Modbus Registers
| Parameter | Register | Unit |
|-----------|----------|------|
| VLN (Voltage Line-to-Neutral) | 141 | V |
| PF (Power Factor) | 117 | — |
| Frequency | 157 | Hz |
| Watts | 101 | W |
| Amps | 149 | A |

### Excel Sheet Structure
- **Config**: Modbus connection parameters
- **ModelAmp**: Motor specifications (Serial, Model, MaxAmp, MinAmp, Capacitor)
- **Users**: Authorized fitter credentials (encrypted with Caesar cipher shift=3)
- **[Motor_Serial]**: Data logs sorted by Date and S.No

### Test Result Logic
```
IF Amps >= MinAmp AND Amps <= MaxAmp:
    Result = "PASS" (Green)
ELIF Amps < MinAmp:
    Result = "FAIL" (Red) - Reason: LowAmps
ELIF Amps > MaxAmp:
    Result = "FAIL" (Red) - Reason: HighAmps
```

## 🔍 Advanced Features

### Auto-Update System
- Checks GitHub releases on startup
- Downloads and verifies `.sha256` checksums
- Integrates with `updater.exe` for seamless updates

### Excel Data Sorting
- Automatically triggered on ESC key press
- Sorts each sheet by Date + S.No (ascending)
- Preserves Config, ModelAmp, Users sheets

### Barcode Pattern Validation
```regex
^[A-Za-z]+\d+[A-Za-z]?-[A-Za-z]\d{2}/\d{3,4}$
```
Examples: `F4R-24/0001`, `X4.2Y-25/0005`, `B1.5A-12/0001`

### Retry & Resilience
- **Register Read Retries**: Configurable (default 5 attempts)
- **Connection Retries**: Infinite with 3-5 second delays
- **Permission Handling**: Detects locked Excel files, prompts user to close

## 📊 Data Export

All data is logged to `C:\Users\[YourUsername]\Desktop\MotorNoLoadReport.xlsx`

### Sample Column Headers
| A | B | C | D | E | F | G | H | I | J | K | L |
|---|---|---|---|---|---|---|---|---|---|---|---|
| S.No | Model | Date | Time | Volt | Power Factor | Frequency | Watts | Amps | Capacitor | QC | Fitter Name |

## 🛠️ Code Quality (v1.2.0 Update)

### Recent Improvements
- ✅ **Type Hints**: Full Python 3.10+ type annotations across all functions
- ✅ **Code Organization**: PEP 8 compliant imports (stdlib, 3rd-party, local)
- ✅ **Variable Naming**: Improved readability with snake_case conventions
- ✅ **Path Handling**: Cross-platform `os.path.join()` instead of backslash concatenation
- ✅ **String Formatting**: Standardized f-strings throughout
- ✅ **Exception Handling**: Specific exception types (no bare `except:`)
- ✅ **Documentation**: Comprehensive docstrings for all functions

## 📜 License

This project is licensed under the **Apache License 2.0** - see the [LICENSE](LICENSE) file for details.

Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License. You may obtain a copy of the License at:

> http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.

## 📧 Support & Contact

For issues, feature requests, or questions:
- Open an issue on GitHub
- Check existing issues for solutions
- Include error logs and system information

## 📚 Additional Resources

- [Modbus RTU Protocol Documentation](https://en.wikipedia.org/wiki/Modbus)
- [Python Type Hints Reference](https://docs.python.org/3/library/typing.html)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)

---

**Version**: 1.2.0  
**Last Updated**: January 2026  
**Author**: SK96 (Sanjev Kumar)  
**Status**: ✅ Production Ready
