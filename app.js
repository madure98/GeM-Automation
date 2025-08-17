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

    // FIXED: extractBuyer - Remove "Department Name" and clean up buyer information
    extractBuyer(text) {
        // Look for Ministry/State Name patterns (IMPROVED)
        const patterns = [
            /Ministry\s+Of\s+[A-Za-z\s&]+/i,
            /Ministry\s+[A-Za-z\s&]+/i,
            /Department\s+Of\s+[A-Za-z\s&]+/i
        ];

        for (const pattern of patterns) {
            const match = text.match(pattern);
            if (match) {
                let buyer = match[0].trim();
                
                // Clean up the buyer name - remove "Department Name" and similar
                buyer = buyer.replace(/\s*Department\s+Name\s*$/i, '');
                buyer = buyer.replace(/\s*Organisation\s+Name\s*$/i, '');
                buyer = buyer.replace(/\s*Office\s+Name\s*$/i, '');
                buyer = buyer.replace(/\s*Buyer\s+Email\s*$/i, '');
                
                // Remove any trailing commas or extra spaces
                buyer = buyer.replace(/[,\s]+$/, '').trim();
                
                if (buyer.length > 5) {
                    console.log('Found buyer:', buyer);
                    return buyer;
                }
            }
        }
        
        // Fallback: look for specific ministry patterns
        const ministryPatterns = [
            'Ministry of Coal',
            'Ministry of Defence', 
            'Ministry of Petroleum and Natural Gas',
            'Ministry of Commerce and Industry'
        ];
        
        for (const ministry of ministryPatterns) {
            if (text.includes(ministry)) {
                console.log('Found ministry from pattern:', ministry);
                return ministry;
            }
        }
        
        console.log('No buyer found, using default');
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

    // COMPLETELY REWRITTEN: extractItemCategory - Extract actual items from GeM bidding PDFs
    extractItemCategory(text) {
        const lines = text.split('\n');
        const items = [];
        
        // Look for the actual item list section
        let inItemSection = false;
        let foundItems = false;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            // Look for the start of actual item specifications
            if (line.includes('Item Category/') || 
                line.includes('मद केटेगरी') ||
                line.includes('VARIOUS TYPES OF ELECTRICAL SPARES') ||
                line.includes('CIRCUIT BREAKER') ||
                line.includes('Bid Details') && line.includes('बिड विवरण')) {
                
                inItemSection = true;
                console.log('Found item section at line:', i, 'Content:', line);
                continue;
            }
            
            if (inItemSection) {
                const lineClean = line.trim();
                
                // Stop when we hit policy sections or other non-item content
                if (['Experience Criteria', 'Preference to Make In India', 'Purchase preference for MSEs', 
                     'Estimated Bid Value', 'Past Performance', 'Technical Specifications',
                     'Consignees/Reporting Officer', 'Buyer Added Bid', 'Input Tax Credit',
                     'EMD Detail', 'Evaluation Method', 'Inspection Required', 'Advisory',
                     'Searched Strings', 'Relevant Categories', 'MSE Exemption'].some(stop => 
                     lineClean.includes(stop))) {
                    console.log('Stopping at line:', i, 'due to:', lineClean);
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
                    lineClean.length > 10 &&
                    !lineClean.includes('Policy') &&
                    !lineClean.includes('criteria') &&
                    !lineClean.includes('preference') &&
                    !lineClean.includes('percentage') &&
                    !lineClean.includes('registration') &&
                    !lineClean.includes('financial year') &&
                    !lineClean.includes('Advisory') &&
                    !lineClean.includes('Department Name') &&
                    !lineClean.includes('Organisation Name') &&
                    !lineClean.includes('Office Name') &&
                    !lineClean.includes('Buyer Email')) {
                    
                    // Clean up the line
                    let cleanLine = lineClean;
                    cleanLine = cleanLine.replace(/^\d+\.\s*/, ''); // Remove numbering
                    cleanLine = cleanLine.replace(/^[^\w\s]*\s*/, ''); // Remove special chars at start
                    cleanLine = cleanLine.replace(/\s+/g, ' ').trim(); // Normalize spaces
                    
                    if (cleanLine.length > 10 && cleanLine.length < 300) {
                        // Check if this looks like an actual item (not policy text)
                        const hasItemIndicators = 
                            cleanLine.includes('PART NO.') ||
                            cleanLine.includes('CIRCUIT BREAKER') ||
                            cleanLine.includes('Switch') ||
                            cleanLine.includes('Socket') ||
                            cleanLine.includes('Plug') ||
                            cleanLine.includes('Holder') ||
                            cleanLine.includes('Cable') ||
                            cleanLine.includes('Lamp') ||
                            cleanLine.includes('Light') ||
                            cleanLine.includes('Contactor') ||
                            cleanLine.includes('Relay') ||
                            cleanLine.includes('Timer') ||
                            cleanLine.includes('Diode') ||
                            cleanLine.includes('Insulation') ||
                            cleanLine.includes('Capacitor') ||
                            cleanLine.includes('Volt') ||
                            cleanLine.includes('Amp') ||
                            cleanLine.includes('Watt') ||
                            cleanLine.includes('sq.mm') ||
                            cleanLine.includes('mFD') ||
                            cleanLine.includes('Hz') ||
                            cleanLine.includes('Phase') ||
                            cleanLine.includes('Pole') ||
                            cleanLine.includes('Range') ||
                            cleanLine.includes('Make') ||
                            cleanLine.includes('Brand') ||
                            cleanLine.includes('Model') ||
                            cleanLine.includes('Type');
                        
                        if (hasItemIndicators) {
                            items.push(cleanLine);
                            foundItems = true;
                            console.log('Found item:', cleanLine);
                        }
                    }
                }
            }
        }
        
        // If we didn't find items in the main section, look for specific patterns
        if (!foundItems) {
            console.log('No categorized items found, looking for specific patterns...');
            
            // Look for circuit breaker patterns
            const circuitBreakerPatterns = [
                /CIRCUIT\s+BREAKER\s+[A-Z\s]+PART\s+NO\.\s+[A-Z0-9]+/gi,
                /Circuit\s+Breaker\s+[A-Za-z\s]+Part\s+No\.\s+[A-Z0-9]+/gi,
                /[A-Z]+\s+PART\s+NO\.\s+[A-Z0-9]+/gi
            ];
            
            for (const pattern of circuitBreakerPatterns) {
                const matches = text.match(pattern);
                if (matches) {
                    items.push(...matches);
                    foundItems = true;
                    console.log('Found circuit breaker pattern:', matches);
                }
            }
            
            // Look for other electrical component patterns
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
                    items.push(...matches);
                    foundItems = true;
                }
            }
        }
        
        // Format the found items
        if (items.length > 0) {
            // Clean and deduplicate items
            const cleanItems = [...new Set(items)]
                .filter(item => item.length > 10 && item.length < 300)
                .slice(0, 20); // Limit to first 20 items
            
            console.log('Found items:', cleanItems.length);
            return cleanItems.join(' | ');
        }
        
        console.log('No items found');
        return 'No items found - may be policy document or scanned image';
    }

    // IMPROVED: extractTechnicalSpec - Show full specification details, not truncated text
    extractTechnicalSpec(text) {
        const lines = text.split('\n');
        const techSpecs = [];
        
        // Look for Technical Specifications section
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            if (line.includes('Technical Specifications/') || 
                line.includes('Technical Specifications') ||
                line.includes('तकनीकी विशिष्टियाँ')) {
                
                console.log('Found Technical Specifications section at line:', i);
                
                // Look ahead for specification details
                for (let j = i + 1; j < Math.min(i + 100, lines.length); j++) {
                    const currentLine = lines[j].trim();
                    
                    // Stop if we hit next major section
                    if (['Item Category', 'मद केटेगरी', 'Experience Criteria', 'Preference to Make In India', 
                         'Purchase preference for MSEs', 'Estimated Bid Value', 'Past Performance',
                         'Consignees/Reporting Officer', 'Buyer Added Bid', 'Input Tax Credit',
                         'Advisory', 'Bid Details', 'बिड विवरण'].some(stop => 
                         currentLine.includes(stop))) {
                        console.log('Stopping at line:', j, 'due to:', currentLine);
                        break;
                    }
                    
                    // Look for specification content
                    if (currentLine && currentLine.length > 10) {
                        // Check for technical specification patterns
                        const hasSpecIndicators = 
                            currentLine.includes('Specification') ||
                            currentLine.includes('Document') ||
                            currentLine.includes('BOQ') ||
                            currentLine.includes('View File') ||
                            currentLine.includes('Download') ||
                            currentLine.includes('Click here') ||
                            currentLine.includes('Open') ||
                            currentLine.includes('Access') ||
                            currentLine.includes('IS:') ||
                            currentLine.includes('IEC') ||
                            currentLine.includes('conforming to') ||
                            currentLine.includes('as per') ||
                            currentLine.includes('Standard') ||
                            currentLine.includes('Type') ||
                            currentLine.includes('Range') ||
                            currentLine.includes('Make') ||
                            currentLine.includes('Brand') ||
                            currentLine.includes('Model') ||
                            currentLine.includes('Part No.') ||
                            currentLine.includes('Voltage') ||
                            currentLine.includes('Current') ||
                            currentLine.includes('Power') ||
                            currentLine.includes('Frequency') ||
                            currentLine.includes('Phase') ||
                            currentLine.includes('Pole');
                        
                        if (hasSpecIndicators) {
                            // Clean up the specification
                            let spec = currentLine;
                            spec = spec.replace(/^\d+\.\s*/, ''); // Remove numbering
                            spec = spec.replace(/^[^\w\s]*\s*/, ''); // Remove special chars
                            spec = spec.replace(/\s+/g, ' ').trim(); // Normalize spaces
                            
                            if (spec.length > 15 && spec.length < 500) { // Increased length limit
                                techSpecs.push(spec);
                                console.log('Found tech spec:', spec);
                            }
                        }
                    }
                }
                break;
            }
        }
        
        // If we didn't find specs in the Technical Specifications section,
        // look for them in the entire document
        if (techSpecs.length === 0) {
            console.log('No tech specs found in Technical Specifications section, searching entire document...');
            
            // Look for specification patterns throughout the text
            const specPatterns = [
                /IS:\s*\d+[^,]*/gi,
                /IEC\s*\d+[^,]*/gi,
                /conforming\s+to[^,]*/gi,
                /as\s+per[^,]*/gi,
                /Type\s+[A-Z][^,]*/gi,
                /Range\s+\d+[^,]*/gi,
                /Make\s+[A-Z][^,]*/gi,
                /Brand\s+[A-Z][^,]*/gi,
                /Model\s+[A-Z][^,]*/gi,
                /Part\s+No\.\s+[A-Z0-9]+/gi
            ];
            
            for (const pattern of specPatterns) {
                const matches = text.match(pattern);
                if (matches) {
                    techSpecs.push(...matches);
                }
            }
        }
        
        // Return the found technical specifications
        if (techSpecs.length > 0) {
            const uniqueSpecs = [...new Set(techSpecs)];
            const result = uniqueSpecs.slice(0, 10).join(' | '); // Limit to first 10 specs
            
            console.log('Final tech specs found:', uniqueSpecs.length);
            return result;
        }
        
        console.log('No technical specifications found');
        return 'Technical specifications available in document';
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

    // NEW FUNCTION: Categorize bids as active or expired
    categorizeBids(data) {
        const now = new Date();
        const activeBids = [];
        const expiredBids = [];
        
        data.forEach(bid => {
            try {
                // Parse start date
                const startDateStr = bid['BID Start Date'];
                let startDate = null;
                
                if (startDateStr && startDateStr !== 'Not Found') {
                    // Handle different date formats
                    if (startDateStr.includes('-')) {
                        const [day, month, year] = startDateStr.split('-');
                        startDate = new Date(year, month - 1, day);
                    } else if (startDateStr.includes('/')) {
                        const [day, month, year] = startDateStr.split('/');
                        startDate = new Date(year, month - 1, day);
                    }
                }
                
                // Parse validity period
                const validityStr = bid['BID Offer Validity'];
                let validityDays = 0;
                
                if (validityStr && validityStr !== 'Not Found') {
                    const match = validityStr.match(/(\d+)/);
                    if (match) {
                        validityDays = parseInt(match[1]);
                    }
                }
                
                // Calculate end date
                let endDate = null;
                if (startDate && validityDays > 0) {
                    endDate = new Date(startDate);
                    endDate.setDate(endDate.getDate() + validityDays);
                }
                
                // Categorize bid
                if (endDate && endDate > now) {
                    activeBids.push({
                        ...bid,
                        'BID End Date': endDate.toLocaleDateString('en-GB'),
                        'Status': 'Active'
                    });
                } else {
                    expiredBids.push({
                        ...bid,
                        'BID End Date': endDate ? endDate.toLocaleDateString('en-GB') : 'Unknown',
                        'Status': 'Expired'
                    });
                }
                
            } catch (error) {
                console.warn('Error categorizing bid:', error);
                // If we can't categorize, put in expired by default
                expiredBids.push({
                    ...bid,
                    'BID End Date': 'Unknown',
                    'Status': 'Unknown'
                });
            }
        });
        
        return { activeBids, expiredBids };
    }

    async createExcelFile() {
        if (this.extractedData.length === 0) {
            throw new Error('No data to export');
        }

        // Categorize bids
        const { activeBids, expiredBids } = this.categorizeBids(this.extractedData);

        // Create workbook
        const wb = XLSX.utils.book_new();
        
        // Create main data worksheet
        const wsMain = XLSX.utils.json_to_sheet(this.extractedData);
        this.formatWorksheet(wsMain, 'Main Data');
        XLSX.utils.book_append_sheet(wb, wsMain, 'All Bids');
        
        // Create active bids worksheet
        if (activeBids.length > 0) {
            const wsActive = XLSX.utils.json_to_sheet(activeBids);
            this.formatWorksheet(wsActive, 'Active Bids');
            XLSX.utils.book_append_sheet(wb, wsActive, 'Active Bids');
        }
        
        // Create expired bids worksheet
        if (expiredBids.length > 0) {
            const wsExpired = XLSX.utils.json_to_sheet(expiredBids);
            this.formatWorksheet(wsExpired, 'Expired Bids');
            XLSX.utils.book_append_sheet(wb, wsExpired, 'Expired Bids');
        }
        
        // Create summary worksheet
        const summaryData = [
            { 'Category': 'Total Bids', 'Count': this.extractedData.length },
            { 'Category': 'Active Bids', 'Count': activeBids.length },
            { 'Category': 'Expired Bids', 'Count': expiredBids.length },
            { 'Category': 'Success Rate', 'Count': `${((activeBids.length / this.extractedData.length) * 100).toFixed(1)}%` }
        ];
        const wsSummary = XLSX.utils.json_to_sheet(summaryData);
        this.formatWorksheet(wsSummary, 'Summary');
        XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

        // Generate Excel file with better options
        const excelBuffer = XLSX.write(wb, { 
            bookType: 'xlsx', 
            type: 'array',
            compression: true
        });
        
        // Create download link with timestamp
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        
        const downloadBtn = document.getElementById('downloadBtn');
        downloadBtn.href = url;
        
        // Better filename with timestamp
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        downloadBtn.download = `gem_bidding_extracted_data_${timestamp}.xlsx`;
        
        // Clean up the URL after a delay
        setTimeout(() => URL.revokeObjectURL(url), 60000);
    }

    // NEW FUNCTION: Format worksheet with consistent styling
    formatWorksheet(ws, title) {
        // Auto-size columns
        const colWidths = [];
        const maxColWidth = 100;
        
        // Get all data rows
        const dataRows = [];
        for (const cell in ws) {
            if (cell[0] === '!') continue; // Skip special properties
            const row = parseInt(cell.replace(/[A-Z]/g, ''));
            if (!dataRows.includes(row)) dataRows.push(row);
        }
        
        dataRows.forEach(row => {
            Object.keys(ws).forEach(cell => {
                if (cell[0] === '!') return;
                const col = cell.replace(/\d/g, '');
                const colIndex = XLSX.utils.decode_col(col);
                const cellValue = ws[cell].v || '';
                const valueLength = String(cellValue).length;
                
                colWidths[colIndex] = Math.max(colWidths[colIndex] || 0, 
                    Math.min(valueLength, maxColWidth));
            });
        });

        // Set column widths with better formatting
        ws['!cols'] = colWidths.map(width => ({ 
            width: Math.min(width + 3, maxColWidth + 5),
            wrapText: true
        }));

        // Set row heights for better readability
        ws['!rows'] = [];
        const maxRow = Math.max(...dataRows);
        for (let i = 0; i <= maxRow; i++) {
            ws['!rows'][i] = { hpt: 25 };
        }

        // Add styling to header row
        const headerRow = 1;
        Object.keys(ws).forEach(cell => {
            if (cell[0] === '!') return;
            const row = parseInt(cell.replace(/[A-Z]/g, ''));
            if (row === headerRow) {
                ws[cell].s = {
                    font: { bold: true },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    fill: { fgColor: { rgb: "E6E6FA" } }
                };
            }
        });
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
    }

    // IMPROVED: Show summary of extracted data with bid categorization
    showExtractionSummary() {
        const summaryDiv = document.getElementById('extractionSummary');
        if (!summaryDiv || this.extractedData.length === 0) return;

        // Categorize bids for summary
        const { activeBids, expiredBids } = this.categorizeBids(this.extractedData);

        let summaryHTML = '<h4>Extraction Summary</h4>';
        
        // Overall statistics
        summaryHTML += `
            <div class="summary-item">
                <span class="summary-label">Total Bids Processed:</span>
                <span class="summary-value">${this.extractedData.length}</span>
            </div>
            <div class="summary-item">
                <span class="summary-label">Active Bids:</span>
                <span class="summary-value" style="color: #28a745;">${activeBids.length}</span>
            </div>
            <div class="summary-item">
                <span class="summary-label">Expired Bids:</span>
                <span class="summary-value" style="color: #dc3545;">${expiredBids.length}</span>
            </div>
            <hr style="margin: 15px 0; border: none; border-top: 1px solid #e9ecef;">
        `;
        
        // File details
        this.extractedData.forEach((data, index) => {
            const filename = this.uploadedFiles[index]?.name || `File ${index + 1}`;
            const isActive = activeBids.some(bid => 
                bid['BID Number'] === data['BID Number'] || 
                bid['BID Start Date'] === data['BID Start Date']
            );
            const statusColor = isActive ? '#28a745' : '#dc3545';
            const statusText = isActive ? 'Active' : 'Expired';
            
            summaryHTML += `
                <div class="summary-item">
                    <span class="summary-label">File:</span>
                    <span class="summary-value">${filename}</span>
                </div>
                <div class="summary-item">
                    <span class="summary-label">Status:</span>
                    <span class="summary-value" style="color: ${statusColor}; font-weight: 600;">${statusText}</span>
                </div>
            `;
            
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
                const itemCount = data['Item Category'].split('\n').length;
                summaryHTML += `<div class="summary-item"><span class="summary-label">Categories Found:</span><span class="summary-value">${itemCount} groups</span></div>`;
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
try {
    console.log('Initializing GeM Bidding Data Extractor...');
    const app = new GeMBiddingDataExtractor();
    console.log('App created successfully:', app);

    // Export the app for global access
    window.app = app;
    console.log('App exported to window.app:', window.app);
    
    // Verify the app is accessible
    if (window.app && typeof window.app.extractDataFromPDF === 'function') {
        console.log('✅ App is properly initialized and accessible');
    } else {
        console.error('❌ App initialization failed - methods not available');
    }
    
} catch (error) {
    console.error('❌ Error initializing app:', error);
    console.error('Stack trace:', error.stack);
}

// Add debug functionality
if (window.app) {
    window.app.toggleDebug = function() {
        const debugSection = document.getElementById('debugSection');
        const debugContent = document.getElementById('debugContent');
        
        if (debugSection && debugContent) {
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
        } else {
            console.warn('Debug elements not found');
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
} else {
    console.error('❌ Cannot add debug functions - app not available');
}
