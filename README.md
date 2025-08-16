# 🚀 GeM Bidding Data Extractor

A modern, professional web application for extracting structured data from GeM (Government e-Marketplace) bidding PDF documents. Built with vanilla JavaScript and designed for seamless deployment on GitHub Pages.

## ✨ Features

- **📁 Drag & Drop Interface**: Intuitive file upload with modern UI
- **🔍 Advanced Text Extraction**: Sophisticated pattern matching for accurate data extraction
- **📊 Excel Generation**: Professional Excel output with proper formatting
- **🌐 Browser-Based**: Runs entirely in the browser - no server required
- **📱 Responsive Design**: Works perfectly on all devices
- **🚀 GitHub Pages Ready**: Deploy instantly without backend setup

## 🏗️ Architecture

This application uses a **pure frontend architecture**:

- **Frontend**: HTML5 + CSS3 + Vanilla JavaScript
- **PDF Processing**: PDF.js for text extraction
- **Data Processing**: Advanced regex patterns for data extraction
- **Excel Generation**: SheetJS for professional Excel output
- **Deployment**: GitHub Pages compatible

## 🚀 Quick Start

### Option 1: GitHub Pages Deployment (Recommended)

1. **Fork this repository** or create a new one
2. **Upload the files** to your repository
3. **Enable GitHub Pages** in repository settings
4. **Access your app** at `https://yourusername.github.io/repo-name/`

### Option 2: Local Development

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/gem-bidding-extractor.git
   cd gem-bidding-extractor
   ```

2. **Open `index.html`** in your web browser
3. **Start extracting data** from your PDFs!

## 📋 Data Extraction Capabilities

The application extracts **comprehensive data** from GeM bidding documents:

- **BID Number**: Unique bid identifier (e.g., GEM/2025/B/7487238)
- **Buyer Information**: Ministry/Department names with proper formatting
- **BID Start Date**: Document start date in DD-MM-YYYY format
- **BID Offer Validity**: Validity period in days
- **Geographic Location**: State and region information
- **Quantity Requirements**: Total quantity specifications
- **Item Categories**: Detailed item descriptions (multi-line support)
- **Technical Specifications**: Document references and specifications

## 🎯 How It Works

1. **PDF Upload**: Users upload GeM bidding PDF files
2. **Text Extraction**: PDF.js extracts text content from all pages
3. **Pattern Recognition**: Advanced regex patterns identify data fields
4. **Data Processing**: Extracted text is cleaned and structured
5. **Excel Generation**: Professional Excel file with organized data
6. **Download**: Users receive formatted Excel output

## 📁 Project Structure

```
gem-bidding-extractor/
├── index.html              # Main application interface
├── app.js                  # Core application logic
├── README.md              # Project documentation
└── .gitignore            # Git ignore rules
```

## 🔧 Technical Details

### **PDF Processing**
- **Library**: PDF.js (Mozilla's PDF rendering engine)
- **Text Extraction**: Full document text extraction
- **Page Support**: Multi-page PDF processing
- **Encoding**: Handles various PDF text encodings

### **Data Extraction**
- **Pattern Matching**: Sophisticated regex patterns
- **Text Cleaning**: Removal of artifacts and noise
- **Multi-line Support**: Complete content extraction
- **Fallback Logic**: Graceful handling of missing data

### **Excel Generation**
- **Library**: SheetJS (Professional Excel library)
- **Formatting**: Auto-sized columns and proper styling
- **Data Structure**: Organized, readable output
- **File Format**: Modern .xlsx format

## 🌐 Browser Compatibility

- ✅ **Chrome 80+** (Recommended)
- ✅ **Firefox 75+**
- ✅ **Safari 13+**
- ✅ **Edge 80+**

## 📊 Performance Characteristics

- **File Size Limit**: ~50MB per PDF (browser memory constraints)
- **Processing Speed**: Depends on PDF complexity and size
- **Memory Usage**: Optimized for browser environments
- **Concurrent Processing**: Single file processing for stability

## 🚨 Requirements & Limitations

### **PDF Requirements**
- **Text Content**: PDFs must contain extractable text (not scanned images)
- **File Format**: Standard PDF format
- **Encoding**: UTF-8 compatible text

### **Browser Requirements**
- **JavaScript**: Must be enabled
- **Memory**: Sufficient RAM for PDF processing
- **Network**: CDN access for required libraries

## 🔍 Troubleshooting

### **Common Issues**

1. **"PDF processing failed"**:
   - Ensure PDF contains extractable text
   - Check file size (keep under 50MB)
   - Verify PDF is not password-protected

2. **"Excel generation error"**:
   - Check browser compatibility
   - Ensure sufficient memory available
   - Verify JavaScript is enabled

3. **"Extraction incomplete"**:
   - PDF format may be non-standard
   - Text encoding issues
   - Complex document structure

### **Debug Tools**

- **Browser Console**: Check for JavaScript errors
- **Network Tab**: Monitor library loading
- **Built-in Debug**: Use debug toggle in the application

## 🎨 Customization

### **UI Customization**
- **Colors**: Modify CSS variables for branding
- **Layout**: Adjust CSS grid and flexbox properties
- **Typography**: Update font families and sizes

### **Functionality Extension**
- **New Fields**: Add extraction patterns in `app.js`
- **Data Processing**: Modify extraction logic
- **Output Format**: Customize Excel generation

## 🤝 Contributing

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature-name`
3. **Make your changes**
4. **Test thoroughly** with real PDFs
5. **Submit a pull request**

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

## 🆘 Support

- **GitHub Issues**: Report bugs and request features
- **Documentation**: Check this README for solutions
- **Testing**: Use real GeM bidding PDFs for verification

---

## 🎉 Ready to Extract Data?

Your professional GeM bidding data extractor is ready for deployment! 

**Deploy to GitHub Pages** and start extracting structured data from your bidding documents with a modern, professional interface.

**No backend required. No server setup. Just pure browser power!** 🚀
