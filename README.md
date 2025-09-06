# 🚀 Email Extractor Pro

**Advanced Domain Email Harvester with Modern GUI**

A powerful, high-performance Python application for extracting email addresses from domain lists with advanced validation, real-time processing controls, and professional reporting capabilities.

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![PyQt5](https://img.shields.io/badge/GUI-PyQt5-green.svg)
![Async](https://img.shields.io/badge/Processing-Async-orange.svg)
![License](https://img.shields.io/badge/License-MIT-red.svg)

## ✨ Features

### 🎯 Core Functionality
- **High-Performance Async Processing** - Concurrent domain processing with smart batch handling
- **Modern PyQt5 GUI** - Professional interface with real-time progress tracking
- **Advanced Email Validation** - Regex patterns + exclusion filters for data quality
- **Smart Deduplication** - Domain-specific email deduplication prevents duplicates
- **Auto-Save Functionality** - Automatic Excel exports every 500 processed domains

### ⭐ Advanced Features
- **🎮 Pause/Resume/Stop Controls** - Full process control with live speed monitoring
- **✅ Domain Validation** - Format checking, duplicate removal, exclude patterns
- **🌐 TLD Analytics** - Real-time categorization and statistics by domain extension
- **📊 Multiple Export Formats** - Excel with statistics + Plain text export
- **🚦 Live Progress Tracking** - Speed metrics (emails/min) and ETA calculations
- **🎨 Modern UI Design** - Sleek, compact interface with color-coded sections

## 🖥️ Screenshots

### Main Interface
- **Gradient Title Bar** with professional branding
- **Color-Coded Sections** for easy navigation
- **Side-by-Side Layout** for validation results and TLD distribution
- **Compact Controls** with icon buttons and hover effects

### Key Sections
1. **📁 Domain File Selection** - File browsing and validation controls
2. **📊 Validation Results** - Domain statistics and processing estimates  
3. **🌐 TLD Distribution** - Top-level domain analytics in two-column format
4. **🎮 Processing Controls** - Start, pause, resume, stop with visual feedback
5. **📈 Progress Tracking** - Real-time progress with speed and ETA
6. **📧 Results Table** - Live email discovery with source tracking
7. **💾 Export Options** - Excel and text export with comprehensive data

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- Windows 10/11 (optimized for Windows PowerShell)

### Setup Instructions

1. **Clone the Repository**
   ```bash
   git clone https://github.com/starsigns/extractemailsfromdomain.git
   cd extractemailsfromdomain
   ```

2. **Create Virtual Environment**
   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Application**
   ```bash
   python main.py
   ```

## 📋 Dependencies

```
aiohttp>=3.8.0
beautifulsoup4>=4.11.0
PyQt5>=5.15.0
openpyxl>=3.0.0
```

## 📖 Usage Guide

### 1. **Loading Domain List**
- Click **📂 Browse** to select a text file containing domains (one per line)
- Supports both plain domains (`example.com`) and full URLs (`https://example.com`)
- Click **✓ Validate** to check domain format and remove duplicates

### 2. **Domain Validation Features**
- **Format Validation** - Ensures proper domain structure
- **Duplicate Removal** - Automatically removes duplicate entries
- **Exclude Patterns** - Add regex patterns to skip unwanted domains
- **TLD Statistics** - View distribution of top-level domains

### 3. **Processing Controls**
- **▶️ Start** - Begin email extraction process
- **⏸️ Pause** - Temporarily halt processing (can resume)
- **▶️ Resume** - Continue from where you paused
- **⏹️ Stop** - Completely halt and enable export of partial results

### 4. **Real-Time Monitoring**
- **Progress Bar** - Visual progress indication
- **Speed Tracking** - Emails per minute processing rate
- **ETA Calculation** - Estimated time to completion
- **Live Results** - Table updates as emails are discovered

### 5. **Export Options**
- **📊 Excel Export** - Comprehensive spreadsheet with:
  - Email listings with source domains and URLs
  - Summary statistics (total emails, unique count, duplicates prevented)
  - Top domain rankings by email count
  - TLD distribution analysis (top 15 extensions)
- **📄 Text Export** - Clean text file with unique emails (one per line)

## ⚙️ Configuration

### Email Validation
The application includes advanced validation patterns to exclude:
- Generic/placeholder emails (`info@`, `admin@`, `contact@`)
- Invalid formats and suspicious patterns
- Configurable exclude patterns for custom filtering

### Performance Settings
- **Concurrent Processing** - 100-domain batches with 200 connection limit
- **Smart Link Filtering** - Prioritizes contact/about pages (5 links max per domain)
- **Memory Optimization** - 1MB response limit per page
- **GUI Responsiveness** - Batched updates every 500ms

## 🛠️ Technical Details

### Architecture
- **Async Worker Thread** - Background processing with PyQt signals
- **Main GUI Thread** - Responsive interface with live updates
- **Batch Processing** - Optimized concurrent domain handling
- **Smart Filtering** - Priority link selection for faster results

### Performance Features
- **Connection Pooling** - Reused HTTP connections for efficiency
- **DNS Caching** - 300-second TTL for faster lookups
- **Timeout Management** - 3-second total, 1-second connect timeouts
- **Error Handling** - Graceful handling of network issues and invalid domains

## 📊 Output Formats

### Excel Export Features
- **Email Listings** - Complete email database with metadata
- **Summary Statistics** - Processing metrics and data quality info
- **Domain Rankings** - Top domains by email count
- **TLD Analysis** - Distribution of domain extensions
- **Timestamps** - Processing date and time tracking

### Text Export Features
- **Unique Emails Only** - Deduplicated email list
- **One Email Per Line** - Clean format for import/processing
- **Instant Export** - Available even during partial processing

## 🚦 Status Indicators

### Visual Feedback
- **🟢 Green Buttons** - Safe actions (Start, Validate)
- **🟡 Yellow Buttons** - Caution actions (Pause)
- **🔵 Blue Buttons** - Information actions (Resume, Browse)
- **🔴 Red Buttons** - Stop actions (Stop processing)
- **📊 Color-Coded Sections** - Easy identification of different functions

## 🐛 Troubleshooting

### Common Issues
1. **Module Not Found** - Ensure virtual environment is activated
2. **Permission Errors** - Run PowerShell as Administrator if needed
3. **Network Timeouts** - Check internet connection and firewall settings
4. **Large Files** - Use smaller domain batches for very large lists

### Performance Tips
- Use SSD storage for better I/O performance
- Ensure stable internet connection for optimal speed
- Close unnecessary applications to free system resources
- Use exclude patterns to skip known problematic domains

## 📈 Performance Metrics

### Typical Processing Speeds
- **Small Lists** (1-100 domains) - ~2-3 seconds per domain
- **Medium Lists** (100-1000 domains) - ~3-4 seconds per domain  
- **Large Lists** (1000+ domains) - ~4-5 seconds per domain

### Factors Affecting Speed
- Domain response times
- Number of emails per domain
- Network connection quality
- System resources available

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **PyQt5** - For the excellent GUI framework
- **aiohttp** - For high-performance async HTTP processing
- **BeautifulSoup** - For reliable HTML parsing
- **openpyxl** - For Excel export functionality

## 📞 Support

For issues, questions, or feature requests:
- Create an issue on GitHub
- Check the troubleshooting section above
- Review the usage guide for common solutions

---

**Email Extractor Pro** - Professional domain email harvesting made simple and efficient! 🚀
