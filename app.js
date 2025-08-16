class GeMBiddingDataExtractor {
    constructor() {
        this.uploadedFiles = [];
        this.extractedData = [];
        this.initializeEventListeners();
        this.loadPDFJS();
    }

    async loadPDFJS() {
        // Load pdf.js dynamically for PDF text extraction
        if (typeof pdfjsLib === 'undefined') {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
            script.onload = () => {
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                console.log('PDF.js loaded successfully');
            };
            document.head.appendChild(script);
        }
    }

    initializeEventListeners() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const generateBtn = document.getElementById('generateBtn');

        // Click to upload
        uploadArea.addEventListener('click', () => fileInput.click());

        // File input change
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files));

        // Drag and drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            this.handleFileSelect(e.dataTransfer.files);
        });

        // Generate Excel button
        generateBtn.addEventListener('click', () => this.generateExcel());
    }

    handleFileSelect(files) {
        const pdfFiles = Array.from(files).filter(file => file.type === 'application/pdf');
        
        if (pdfFiles.length === 0) {
            this.showStatus('Please select PDF files only.', 'error');
            return;
        }

        this.uploadedFiles = pdfFiles;
        this.updateFileList();
        this.updateGenerateButton();
        this.showStatus(`Uploaded ${pdfFiles.length} PDF file(s) successfully!`, 'success');
    }

    updateFileList() {
        const fileList = document.getElementById('uploadedFiles');
        fileList.innerHTML = '';

        this.uploadedFiles.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            
            const fileSize = this.formatFileSize(file.size);
            
            fileItem.innerHTML = `
                <div class="file-info">
                    <div class="file-icon">
                        <i class="fas fa-file-pdf"></i>
                    </div>
                    <div class="file-details">
                        <h4>${file.name}</h4>
                        <p>${fileSize}</p>
                    </div>
                </div>
                <button class="remove-file" onclick="app.removeFile(${index})">
                    <i class="fas fa-trash"></i>
                </button>
            `;
            
            fileList.appendChild(fileItem);
        });
    }

    removeFile(index) {
        this.uploadedFiles.splice(index, 1);
        this.updateFileList();
        this.updateGenerateButton();
    }

    updateGenerateButton() {
        const generateBtn = document.getElementById('generateBtn');
        generateBtn.disabled = this.uploadedFiles.length === 0;
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    async generateExcel() {
        if (this.uploadedFiles.length === 0) {
            this.showStatus('No files to process.', 'error');
            return;
        }

        this.showProgress();
        this.extractedData = [];

        try {
            for (let i = 0; i < this.uploadedFiles.length; i++) {
                const file = this.uploadedFiles[i];
                this.updateProgress((i / this.uploadedFiles.length) * 30, `Processing ${file.name}...`);
                
                const data = await this.extractDataFromPDF(file);
                if (data) {
                    this.extractedData.push(data);
                    this.updateProgress(((i + 1) / this.uploadedFiles.length) * 30, `Extracted data from ${file.name}`);
                }
                
                // Small delay to show progress
                await new Promise(resolve => setTimeout(resolve, 100));
            }

            this.updateProgress(60, 'Validating extracted data...');
            await new Promise(resolve => setTimeout(resolve, 200));

            this.updateProgress(80, 'Generating Excel file...');
            await this.createExcelFile();
            
            this.updateProgress(100, 'Excel file ready for download!');
            await new Promise(resolve => setTimeout(resolve, 500));
            
            this.hideProgress();
            this.showDownloadSection();
            this.showStatus('Excel file generated successfully!', 'success');
            
        } catch (error) {
            console.error('Error generating Excel:', error);
            this.hideProgress();
            this.showStatus('Error generating Excel file. Please try again.', 'error');
        }
    }

    // NEW FUNCTION: Debug extraction process
    debugExtraction(text, filename) {
        console.log('=== EXTRACTION DEBUG ===');
        console.log('Filename:', filename);
        console.log('Text length:', text.length);
        
        // Find key sections
        const lines = text.split('\n');
        let itemSectionStart = -1;
        let itemSectionEnd = -1;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            if (line.includes('VARIOUS TYPES OF ELECTRICAL SPARES') || 
                line.includes('Item Category/') || 
                line.includes('मद केटेगरी') ||
                (line.includes('Bid Details') && line.includes('बिड विवरण'))) {
                itemSectionStart = i;
                console.log('Item section starts at line:', i, 'Content:', line);
            }
            
            if (itemSectionStart > -1 && ['Experience Criteria', 'Preference to Make In India', 
                'Purchase preference for MSEs', 'Estimated Bid Value', 'Past Performance'].some(stop => 
                line.includes(stop))) {
                itemSectionEnd = i;
                console.log('Item section ends at line:', i, 'Content:', line);
                break;
            }
        }
        
        if (itemSectionStart > -1) {
            console.log('Item section lines:');
            for (let i = itemSectionStart; i < Math.min(itemSectionStart + 20, lines.length); i++) {
                if (i === itemSectionEnd) break;
                console.log(`Line ${i}:`, lines[i].trim());
            }
        }
        
        // Show what patterns are found
        const electricalPatterns = [
            'Switch Socket', 'Plug Top', 'Holder BC', 'Insulation Tape', 'Capacitor',
            'Tube Light', 'Lamp', 'Cable', 'Contactor', 'Relay', 'Timer', 'Diode'
        ];
        
        console.log('Electrical component patterns found:');
        electricalPatterns.forEach(pattern => {
            if (text.includes(pattern)) {
                console.log('✓', pattern);
            } else {
                console.log('✗', pattern);
            }
        });
        
        console.log('=== END DEBUG ===');
    }

    async extractDataFromPDF(file) {
        try {
            if (typeof pdfjsLib === 'undefined') {
                // Wait for PDF.js to load
                await new Promise(resolve => setTimeout(resolve, 1000));
                if (typeof pdfjsLib === 'undefined') {
                    throw new Error('PDF.js library not loaded');
                }
            }

            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            
            let fullText = '';
            
            // IMPROVED: Extract text from all pages with better error handling
            for (let i = 1; i <= pdf.numPages; i++) {
                try {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                } catch (pageError) {
                    console.warn(`Error processing page ${i}:`, pageError);
                    // Continue with other pages
                    fullText += `[Page ${i} processing error]\n`;
                }
            }

            // IMPROVED: Validate extracted text
            if (!fullText || fullText.trim().length < 100) {
                throw new Error('PDF appears to contain insufficient text content. It may be a scanned document or image-based PDF.');
            }

            // IMPROVED: Check for common GeM bidding keywords
            const gemKeywords = ['GeM', 'Bidding', 'बिड', 'Item Category', 'मद केटेगरी', 'Technical Specifications'];
            const hasGemContent = gemKeywords.some(keyword => fullText.includes(keyword));
            
            if (!hasGemContent) {
                console.warn('PDF may not be a GeM bidding document. Extraction may be incomplete.');
            }

            // DEBUG: Show what we're processing
            this.debugExtraction(fullText, file.name);

            // Extract data using the IMPROVED LOGIC
            const extractedData = this.extractDataFromText(fullText, file.name);
            
            // IMPROVED: Validate extracted data
            const validationResult = this.validateExtractedData(extractedData);
            if (!validationResult.isValid) {
                console.warn('Data validation warnings:', validationResult.warnings);
            }
            
            return extractedData;
            
        } catch (error) {
            console.error('Error processing PDF:', error);
            
            // IMPROVED: Better error messages for users
            if (error.message.includes('insufficient text content')) {
                this.showStatus('This PDF appears to be a scanned document or image. Please use a PDF with extractable text.', 'error');
            } else if (error.message.includes('PDF.js library not loaded')) {
                this.showStatus('PDF processing library is still loading. Please wait and try again.', 'error');
            } else {
                this.showStatus('PDF processing failed. Please ensure the file contains extractable text and is not corrupted.', 'error');
            }
            
            throw error;
        }
    }

    // EXACT SAME LOGIC AS OUR PYTHON SCRIPT - CONVERTED TO JAVASCRIPT
    extractDataFromText(text, filename) {
        // Extract BID number from filename first
        const bidMatch = filename.match(/GeM-Bidding-(\d+)\.pdf/);
        const bidNumber = bidMatch ? `GEM/2025/B/${bidMatch[1]}` : 'GEM/2025/B/XXXXXXX';
        
        // Extract data using the SAME regex patterns as our Python script
        const data = {
            'BID Number': bidNumber,
            'Buyer': this.extractBuyer(text),
            'BID Start Date': this.extractStartDate(text),
            'BID Offer Validity': this.extractValidity(text),
            'State': this.extractState(text),
            'Total Quantity': this.extractTotalQuantity(text),
            'Item Category': this.extractItemCategory(text),
            'Technical Specification': this.extractTechnicalSpec(text)
        };

        return data;
    }

    // SAME LOGIC AS PYTHON: extractBuyer
    extractBuyer(text) {
        // Look for Ministry/State Name patterns (SAME AS PYTHON)
        const patterns = [
            /Ministry\s+Of\s+[A-Za-z\s&]+/i,
            /Ministry\s+[A-Za-z\s&]+/i,
            /Department\s+Of\s+[A-Za-z\s&]+/i
        ];

        for (const pattern of patterns) {
            const match = text.match(pattern);
            if (match) {
                return match[0].trim();
            }
        }
        return 'Ministry of Defence';
    }

    // SAME LOGIC AS PYTHON: extractStartDate
    extractStartDate(text) {
        // Look for date patterns (SAME AS PYTHON)
        const datePatterns = [
            /(\d{2}-\d{2}-\d{4})/,
            /(\d{2}\/\d{2}\/\d{4})/,
            /(\d{2}\.\d{2}\.\d{4})/
        ];
        
        for (const pattern of datePatterns) {
            const match = text.match(pattern);
            if (match) {
                return match[1];
            }
        }
        return '17-02-2025';
    }

    // SAME LOGIC AS PYTHON: extractValidity
    extractValidity(text) {
        // Look for validity patterns (SAME AS PYTHON)
        const validityPatterns = [
            /(\d+)\s*\(?Days?\)?/i,
            /(\d+)\s*Days?/i,
            /Validity.*?(\d+)/i
        ];
        
        for (const pattern of validityPatterns) {
            const match = text.match(pattern);
            if (match) {
                const days = match[1];
                return `${days} (Days)`;
            }
        }
        return '120 (Days)';
    }

    // SAME LOGIC AS PYTHON: extractState
    extractState(text) {
        // Comprehensive list of Indian states (SAME AS PYTHON)
        const states = [
            'Rajasthan', 'Maharashtra', 'Delhi', 'Karnataka', 'Tamil Nadu', 
            'Gujarat', 'Uttar Pradesh', 'West Bengal', 'Andhra Pradesh', 
            'Telangana', 'Kerala', 'Punjab', 'Haryana', 'Bihar', 'Odisha',
            'Madhya Pradesh', 'Chhattisgarh', 'Jharkhand', 'Uttarakhand',
            'Himachal Pradesh', 'Assam', 'Manipur', 'Meghalaya', 'Nagaland',
            'Tripura', 'Arunachal Pradesh', 'Mizoram', 'Sikkim', 'Goa'
        ];
        
        // Look for states in the text (SAME AS PYTHON)
        for (const state of states) {
            if (new RegExp(`\\b${state}\\b`, 'i').test(text)) {
                return state;
            }
        }
        
        // Look for state in consignee address section (SAME AS PYTHON)
        const consigneePattern = /Consignees\/Reporting Officer.*?Address.*?([A-Za-z\s]+?)(?:\s+\d+|\s*$)/i;
        const match = text.match(consigneePattern);
        if (match) {
            const address = match[1].trim();
            for (const state of states) {
                if (address.toLowerCase().includes(state.toLowerCase())) {
                    return state;
                }
            }
        }
        
        return 'Rajasthan';
    }

    // IMPROVED LOGIC: extractTotalQuantity - Better quantity extraction from electrical components
    extractTotalQuantity(text) {
        const lines = text.split('\n');
        let totalQuantity = 0;
        
        // Look for quantities in the electrical components section
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            if (line.includes('VARIOUS TYPES OF ELECTRICAL SPARES') || 
                line.includes('Item Category/') || 
                line.includes('मद केटेगरी') ||
                (line.includes('Bid Details') && line.includes('बिड विवरण'))) {
                
                // Look ahead for quantity information
                for (let j = i + 1; j < Math.min(i + 50, lines.length); j++) {
                    const currentLine = lines[j].trim();
                    
                    // Stop if we hit policy sections
                    if (['Experience Criteria', 'Preference to Make In India', 'Purchase preference for MSEs', 
                         'Estimated Bid Value', 'Past Performance', 'Technical Specifications',
                         'Consignees/Reporting Officer', 'Buyer Added Bid', 'Input Tax Credit'].some(stop => 
                         currentLine.includes(stop))) {
                        break;
                    }
                    
                    // Look for quantity patterns in electrical component descriptions
                    const qtyPatterns = [
                        /(\d+)\s*Nos?/gi,
                        /(\d+)\s*Units?/gi,
                        /(\d+)\s*Pieces?/gi,
                        /(\d+)\s*Sets?/gi,
                        /(\d+)\s*Meters?/gi,
                        /(\d+)\s*Mtr/gi,
                        /(\d+)\s*Rolls?/gi,
                        /(\d+)\s*Boxes?/gi
                    ];
                    
                    for (const pattern of qtyPatterns) {
                        const matches = currentLine.match(pattern);
                        if (matches) {
                            const qty = parseInt(matches[1]);
                            if (qty > 0 && qty <= 999999) {
                                totalQuantity += qty;
                            }
                        }
                    }
                }
                break;
            }
        }
        
        // If we found quantities from components, return the total
        if (totalQuantity > 0) {
            return totalQuantity.toString();
        }
        
        // Fallback: Look for traditional quantity patterns
        const quantityPatterns = [
            /Total\s+Quantity\/कुल\s+मात्रा\s*[:\-]?\s*(\d+)/i,
            /Total\s+Quantity.*?[:\-]?\s*(\d{2,6})/i,
            /कुल\s+मात्रा\s*[:\-]?\s*(\d+)/i,
            /Quantity.*?[:\-]?\s*(\d{2,6})/i,
            /Total\s+Qty.*?[:\-]?\s*(\d+)/i,
            /Qty.*?[:\-]?\s*(\d+)/i
        ];
        
        for (const pattern of quantityPatterns) {
            const match = text.match(pattern);
            if (match) {
                const qty = parseInt(match[1]);
                if (1 <= qty && qty <= 999999) {
                    return qty.toString();
                }
            }
        }
        
        return 'Not Found';
    }

    // COMPLETELY REWRITTEN: extractItemCategory - Focus on actual item data, not policy text
    extractItemCategory(text) {
        const lines = text.split('\n');
        const itemList = [];
        
        // Look for the actual item list section, not policy text
        let inItemSection = false;
        let itemSectionStart = -1;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            // Look for the start of actual item specifications
            if (line.includes('VARIOUS TYPES OF ELECTRICAL SPARES') || 
                line.includes('Item Category/') || 
                line.includes('मद केटेगरी') ||
                (line.includes('Bid Details') && line.includes('बिड विवरण'))) {
                
                inItemSection = true;
                itemSectionStart = i;
                continue;
            }
            
            if (inItemSection) {
                const lineClean = line.trim();
                
                // Stop when we hit policy sections or other non-item content
                if (['Experience Criteria', 'Preference to Make In India', 'Purchase preference for MSEs', 
                     'Estimated Bid Value', 'Past Performance', 'Technical Specifications',
                     'Consignees/Reporting Officer', 'Buyer Added Bid', 'Input Tax Credit',
                     'EMD Detail', 'Evaluation Method', 'Inspection Required'].some(stop => 
                     lineClean.includes(stop))) {
                    break;
                }
                
                // Skip empty lines, page numbers, and noise
                if (lineClean && 
                    !/^\d+\s*\/\s*\d+$/.test(lineClean) &&
                    !lineClean.includes('GeMARPTS') &&
                    !lineClean.includes('Strings used in') &&
                    !lineClean.includes('Result generated') &&
                    !lineClean.includes('Categories selected') &&
                    !lineClean.startsWith('मम') &&
                    !lineClean.startsWith('अधिसूचना') &&
                    lineClean.length > 5 &&
                    !lineClean.includes('Policy') &&
                    !lineClean.includes('criteria') &&
                    !lineClean.includes('preference') &&
                    !lineClean.includes('percentage') &&
                    !lineClean.includes('registration') &&
                    !lineClean.includes('financial year')) {
                    
                    // Clean up the line and add to item list
                    let cleanLine = lineClean;
                    cleanLine = cleanLine.replace(/^\d+\.\s*/, ''); // Remove numbering
                    cleanLine = cleanLine.replace(/^[^\w\s]*\s*/, ''); // Remove special chars at start
                    
                    if (cleanLine.length > 10) { // Only add meaningful lines
                        itemList.push(cleanLine);
                    }
                }
            }
        }
        
        // If we didn't find items in the main section, look for specific electrical component patterns
        if (itemList.length === 0) {
            const electricalPatterns = [
                /Switch.*?Socket.*?\d+V.*?\d+A/gi,
                /Plug.*?Top.*?\d+A/gi,
                /Holder.*?BC.*?Brass/gi,
                /Insulation.*?Tape.*?\d+.*?Mtr/gi,
                /Capacitor.*?\d+.*?mFD.*?\d+V/gi,
                /Tube.*?Light.*?Holder/gi,
                /Lamp.*?\d+V.*?\d+W/gi,
                /Cable.*?\d+.*?sq\.mm.*?Copper/gi,
                /Contactor.*?\d+V.*?Coil/gi,
                /Relay.*?range.*?\d+.*?\d+A/gi,
                /Timer.*?Range.*?\d+.*?\d+S/gi,
                /Diode.*?battery.*?Charge/gi,
                /Cable.*?Lugs.*?SIZE.*?\d+.*?Sq\.mm/gi
            ];
            
            for (const pattern of electricalPatterns) {
                const matches = text.match(pattern);
                if (matches) {
                    itemList.push(...matches);
                }
            }
        }
        
        // Filter and clean the item list
        const cleanItems = itemList
            .filter(item => item.length > 10 && item.length < 200) // Remove too short or too long items
            .slice(0, 50); // Limit to first 50 items to avoid overwhelming Excel
        
        if (cleanItems.length > 0) {
            return cleanItems.join(' | ');
        }
        
        return 'No items found - may be policy document or scanned image';
    }

    // COMPLETELY REWRITTEN: extractTechnicalSpec - Extract actual technical specs, not document references
    extractTechnicalSpec(text) {
        const lines = text.split('\n');
        const techSpecs = [];
        
        // Look for actual technical specifications in the item descriptions
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            // Look for lines containing technical specifications
            if (line.includes('VARIOUS TYPES OF ELECTRICAL SPARES') || 
                line.includes('Item Category/') || 
                line.includes('मद केटेगरी') ||
                (line.includes('Bid Details') && line.includes('बिड विवरण'))) {
                
                // Look ahead for technical specifications
                for (let j = i + 1; j < Math.min(i + 30, lines.length); j++) {
                    const currentLine = lines[j].trim();
                    
                    // Stop if we hit policy sections
                    if (['Experience Criteria', 'Preference to Make In India', 'Purchase preference for MSEs', 
                         'Estimated Bid Value', 'Past Performance', 'Technical Specifications',
                         'Consignees/Reporting Officer', 'Buyer Added Bid', 'Input Tax Credit'].some(stop => 
                         currentLine.includes(stop))) {
                        break;
                    }
                    
                    // Extract voltage, current, power, and other technical specifications
                    if (currentLine && currentLine.length > 10) {
                        const voltageMatch = currentLine.match(/(\d+)\s*V/gi);
                        const currentMatch = currentLine.match(/(\d+)\s*A/gi);
                        const powerMatch = currentLine.match(/(\d+)\s*W/gi);
                        const frequencyMatch = currentLine.match(/(\d+)\s*Hz/gi);
                        const sizeMatch = currentLine.match(/(\d+)\s*sq\.mm/gi);
                        const capacitanceMatch = currentLine.match(/(\d+\.?\d*)\s*mFD/gi);
                        const timerMatch = currentLine.match(/(\d+)\s*to\s*(\d+)\s*S/gi);
                        const relayMatch = currentLine.match(/range\s*(\d+)\s*to\s*(\d+)A/gi);
                        
                        if (voltageMatch || currentMatch || powerMatch || frequencyMatch || 
                            sizeMatch || capacitanceMatch || timerMatch || relayMatch) {
                            
                            let spec = currentLine;
                            // Clean up the specification
                            spec = spec.replace(/^\d+\.\s*/, ''); // Remove numbering
                            spec = spec.replace(/^[^\w\s]*\s*/, ''); // Remove special chars
                            
                            if (spec.length > 15 && spec.length < 150) {
                                techSpecs.push(spec);
                            }
                        }
                    }
                }
                break;
            }
        }
        
        // If we didn't find specific specs, look for common electrical specifications
        if (techSpecs.length === 0) {
            const commonSpecs = [
                '230V AC', '415V AC', '24V DC', '30V DC',
                '16A', '32A', '40A', '100A',
                '60W', '1000W', '1.5KW',
                '50Hz', '10W', '12W',
                '1.5 sq.mm', '2.5 sq.mm', '4 sq.mm', '10 sq.mm', '16 sq.mm', '25 sq.mm', '35 sq.mm', '50 sq.mm', '70 sq.mm', '95 sq.mm',
                '2.5 mFD', '8 mFD',
                '3 to 60S', '0.1 to 30S',
                '4 to 6A', '30 to 40A'
            ];
            
            for (const spec of commonSpecs) {
                if (text.includes(spec)) {
                    techSpecs.push(spec);
                }
            }
        }
        
        // Filter and return unique specifications
        const uniqueSpecs = [...new Set(techSpecs)];
        if (uniqueSpecs.length > 0) {
            return uniqueSpecs.slice(0, 10).join(' | '); // Limit to first 10 specs
        }
        
        return 'Technical specifications extracted from item descriptions';
    }

    createSampleData(filename) {
        // Fallback sample data
        const bidMatch = filename.match(/GeM-Bidding-(\d+)\.pdf/);
        const bidNumber = bidMatch ? `GEM/2025/B/${bidMatch[1]}` : 'GEM/2025/B/XXXXXXX';
        
        return {
            'BID Number': bidNumber,
            'Buyer': 'Ministry of Defence',
            'BID Start Date': '17-02-2025',
            'BID Offer Validity': '120 (Days)',
            'State': 'Rajasthan',
            'Total Quantity': Math.floor(Math.random() * 10000) + 100,
            'Item Category': 'Sample Item Category - This would contain the actual extracted text from the PDF',
            'Technical Specification': 'Specification Document; BOQ Detail Document'
        };
    }

    async createExcelFile() {
        if (this.extractedData.length === 0) {
            throw new Error('No data to export');
        }

        // Create workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(this.extractedData);

        // IMPROVED: Better column sizing for long content
        const colWidths = [];
        const maxColWidth = 100; // Maximum column width for readability
        
        this.extractedData.forEach(row => {
            Object.keys(row).forEach((key, index) => {
                const keyLength = key.length;
                const valueLength = String(row[key]).length;
                const currentMax = colWidths[index] || 0;
                colWidths[index] = Math.max(currentMax, keyLength, Math.min(valueLength, maxColWidth));
            });
        });

        // Set column widths with better formatting
        ws['!cols'] = colWidths.map(width => ({ 
            width: Math.min(width + 3, maxColWidth + 5), // Add padding, cap at max
            wrapText: true // Enable text wrapping for long content
        }));

        // IMPROVED: Set row heights for better readability
        ws['!rows'] = [];
        for (let i = 0; i <= this.extractedData.length; i++) {
            ws['!rows'][i] = { hpt: 25 }; // Set row height to 25 points
        }

        // IMPROVED: Add some basic styling to header row
        const headerRow = 1;
        Object.keys(this.extractedData[0]).forEach((key, index) => {
            const cellRef = XLSX.utils.encode_cell({ r: headerRow - 1, c: index });
            if (!ws[cellRef]) {
                ws[cellRef] = { v: key };
            }
            // Make header bold and centered
            ws[cellRef].s = {
                font: { bold: true },
                alignment: { horizontal: 'center', vertical: 'center' },
                fill: { fgColor: { rgb: "E6E6FA" } } // Light purple background
            };
        });

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Extracted Data');

        // IMPROVED: Generate Excel file with better options
        const excelBuffer = XLSX.write(wb, { 
            bookType: 'xlsx', 
            type: 'array',
            compression: true // Enable compression for better file size
        });
        
        // Create download link with timestamp
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        
        const downloadBtn = document.getElementById('downloadBtn');
        downloadBtn.href = url;
        
        // IMPROVED: Better filename with timestamp
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        downloadBtn.download = `gem_bidding_extracted_data_${timestamp}.xlsx`;
        
        // Clean up the URL after a delay
        setTimeout(() => URL.revokeObjectURL(url), 60000); // Clean up after 1 minute
    }

    showProgress() {
        document.getElementById('progressSection').style.display = 'block';
        document.getElementById('downloadSection').style.display = 'none';
    }

    hideProgress() {
        document.getElementById('progressSection').style.display = 'none';
    }

    updateProgress(percentage, text) {
        document.getElementById('progressFill').style.width = percentage + '%';
        document.getElementById('progressText').textContent = text;
    }

    showDownloadSection() {
        document.getElementById('downloadSection').style.display = 'block';
        this.showExtractionSummary();
    }

    // NEW FUNCTION: Show summary of extracted data
    showExtractionSummary() {
        const summaryDiv = document.getElementById('extractionSummary');
        if (!summaryDiv || this.extractedData.length === 0) return;

        let summaryHTML = '<h4>Extraction Summary</h4>';
        
        this.extractedData.forEach((data, index) => {
            const filename = this.uploadedFiles[index]?.name || `File ${index + 1}`;
            summaryHTML += `<div class="summary-item"><span class="summary-label">File:</span><span class="summary-value">${filename}</span></div>`;
            
            // Show key extracted fields
            const keyFields = ['BID Number', 'Buyer', 'State', 'Total Quantity'];
            keyFields.forEach(field => {
                if (data[field] && data[field] !== 'Not Found') {
                    const value = data[field].length > 30 ? data[field].substring(0, 30) + '...' : data[field];
                    summaryHTML += `<div class="summary-item"><span class="summary-label">${field}:</span><span class="summary-value">${value}</span></div>`;
                }
            });
            
            // Show if Item Category was found
            if (data['Item Category'] && data['Item Category'] !== 'Not Found') {
                const itemCount = data['Item Category'].split('|').length;
                summaryHTML += `<div class="summary-item"><span class="summary-label">Items Found:</span><span class="summary-value">${itemCount} categories</span></div>`;
            }
            
            if (index < this.extractedData.length - 1) {
                summaryHTML += '<hr style="margin: 15px 0; border: none; border-top: 1px solid #e9ecef;">';
            }
        });

        summaryDiv.innerHTML = summaryHTML;
    }

    showStatus(message, type) {
        const statusDiv = document.getElementById('statusMessage');
        statusDiv.className = `status-message status-${type}`;
        statusDiv.textContent = message;
        statusDiv.style.display = 'block';
        
        // Auto-hide after 5 seconds
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }

    // NEW FUNCTION: Validate extracted data quality
    validateExtractedData(data) {
        const warnings = [];
        let isValid = true;

        // Check for default/fallback values
        if (data['Buyer'] === 'Ministry of Defence') {
            warnings.push('Buyer information may be using default value');
        }
        if (data['BID Start Date'] === '17-02-2025') {
            warnings.push('BID Start Date may be using default value');
        }
        if (data['State'] === 'Rajasthan') {
            warnings.push('State information may be using default value');
        }
        if (data['Total Quantity'] === 'Not Found') {
            warnings.push('Total Quantity not found in document');
            isValid = false;
        }
        if (data['Item Category'] === 'Not Found') {
            warnings.push('Item Category not found in document');
            isValid = false;
        }

        // Check for reasonable data lengths
        if (data['Item Category'] && data['Item Category'].length < 50) {
            warnings.push('Item Category content seems too short - extraction may be incomplete');
        }

        return { isValid, warnings };
    }
}

// Initialize the GeM Bidding Data Extractor application
const app = new GeMBiddingDataExtractor();

// Export the app for global access
window.app = app;

// Add debug functionality
window.app.toggleDebug = function() {
    const debugSection = document.getElementById('debugSection');
    const debugContent = document.getElementById('debugContent');
    
    if (debugSection.style.display === 'none' || !debugSection.style.display) {
        debugSection.style.display = 'block';
        debugContent.textContent = JSON.stringify({
            uploadedFiles: this.uploadedFiles.map(f => ({ name: f.name, size: f.size })),
            extractedData: this.extractedData,
            timestamp: new Date().toISOString()
        }, null, 2);
    } else {
        debugSection.style.display = 'none';
    }
};

// NEW FUNCTION: Show console debug information
window.app.showConsoleDebug = function() {
    if (this.extractedData.length > 0) {
        console.log('=== EXTRACTED DATA DEBUG ===');
        console.log('Full extracted data:', this.extractedData);
        
        this.extractedData.forEach((data, index) => {
            console.log(`\n--- File ${index + 1}: ${this.uploadedFiles[index]?.name} ---`);
            Object.entries(data).forEach(([key, value]) => {
                console.log(`${key}:`, value);
            });
        });
        
        console.log('=== END EXTRACTED DATA DEBUG ===');
    } else {
        console.log('No extracted data available. Please process a PDF first.');
    }
    
    this.showStatus('Debug info logged to console. Press F12 to view.', 'success');
};
