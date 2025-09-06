# Email Extractor - Standalone Executable

## 📦 Installation

### Option 1: Direct Use (Recommended)
1. Download `EmailExtractor.exe`
2. Double-click to run immediately - no installation required!

### Option 2: Portable Installation
1. Create a folder on your computer (e.g., `C:\EmailExtractor\`)
2. Copy `EmailExtractor.exe` to this folder
3. Create a desktop shortcut for easy access

## 🚀 How to Use

1. **Launch the Application**
   - Double-click `EmailExtractor.exe`
   - The GUI will open automatically

2. **Load Domains**
   - Click "Browse" to select a text file containing URLs/domains
   - Each URL should be on a separate line

3. **Start Extraction**
   - Click "Start Extraction" to begin
   - Use "Pause" and "Stop" buttons to control the process
   - Progress and statistics are shown in real-time

4. **View Results**
   - Valid emails appear in the left panel
   - Invalid/filtered emails appear in the right panel
   - Statistics show extraction speed and progress

5. **Export Results**
   - Click "Export to Excel" to save results
   - Files are saved with timestamp (e.g., `emails_2025-09-06_19-30-45.xlsx`)

## ✨ Features

- **Fast Async Processing**: Processes multiple domains simultaneously
- **Advanced Email Detection**: Finds regular and obfuscated emails
- **Smart Filtering**: Removes placeholder, invalid, and duplicate emails
- **Real-time Progress**: Live updates with speed and ETA
- **Export Options**: Excel format with detailed statistics
- **Modern GUI**: Clean, responsive interface

## 🔍 Supported Email Formats

### Regular Emails
- `contact@company.com`
- `info@business.org`
- `support@service.net`

### Obfuscated Emails (Automatically Detected)
- `contact[at]company[dot]com`
- `info(at)business(dot)org`
- `support at service dot net`
- `sales &#64; enterprise &#46; com`
- `admin{at}firm{dot}io`

## 🛡️ Smart Filtering

The application automatically filters out:
- Invalid email formats
- Placeholder emails (abc@xyz.com, test@demo.com)
- File extensions (.png@domain.com)
- Short/generic emails (a@b.com)
- Duplicate emails

## 💾 System Requirements

- **Operating System**: Windows 7/8/10/11 (64-bit)
- **Memory**: 4GB RAM minimum, 8GB recommended
- **Disk Space**: 100MB free space
- **Internet**: Required for domain processing

## 🔧 Troubleshooting

### Application Won't Start
- Ensure you have administrative privileges
- Check Windows Defender/antivirus settings
- Try running as administrator (right-click → "Run as administrator")

### Slow Performance
- Close other resource-intensive applications
- Use smaller batches of domains (100-500 at a time)
- Ensure stable internet connection

### No Emails Found
- Verify domain URLs are accessible
- Check if domains have email addresses visible
- Some sites may block automated access

## 📞 Support

For issues or questions:
- Check the GitHub repository: https://github.com/starsigns/extractemailsfromdomain
- Review the README.md file for detailed information
- Submit issues on GitHub for technical problems

## 📄 License

This software is released under the MIT License. See LICENSE file for details.

---

**Version**: 1.0  
**Build Date**: September 6, 2025  
**File Size**: ~48MB  
**Created with**: Python 3.13, PyQt5, PyInstaller
