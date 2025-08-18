class GeMBiddingDataExtractor {
    constructor() {
        this.uploadedFiles = [];
        this.extractedData = [];
        this.isProcessingFiles = false;
        this.isGeneratingExcel = false;
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

        // Prevent multiple event listeners by checking if already initialized
        if (uploadArea.dataset.initialized === 'true') {
            console.log('Event listeners already initialized, skipping...');
            return;
        }

        console.log('Initializing event listeners...');

        // Click to upload - simplified and direct
        uploadArea.addEventListener('click', (e) => {
            // Only handle clicks on the upload area itself, not its children
            if (e.target === uploadArea || e.target.closest('.upload-icon') || e.target.closest('.upload-text') || e.target.closest('.upload-subtext')) {
                console.log('Upload area clicked, opening file selector...');
                fileInput.click();
            }
        });

        // File input change - with validation
        fileInput.addEventListener('change', (e) => {
            console.log('File input change event triggered');
            if (e.target.files && e.target.files.length > 0) {
                this.handleFileSelect(e.target.files);
            }
        });

        // Drag and drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                this.handleFileSelect(e.dataTransfer.files);
            }
        });

        // Generate Excel button - with validation
        generateBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            console.log('Generate button clicked');
            if (!generateBtn.disabled) {
                this.generateExcel();
            }
        });

        // Mark as initialized to prevent duplicate listeners
        uploadArea.dataset.initialized = 'true';
        console.log('Event listeners initialized successfully');
        
        // Add some debugging info
        console.log('Upload area element:', uploadArea);
        console.log('File input element:', fileInput);
        console.log('Generate button element:', generateBtn);
    }

    handleFileSelect(files) {
        // Prevent multiple simultaneous calls
        if (this.isProcessingFiles) {
            console.log('File processing already in progress, ignoring new selection');
            return;
        }
        
        console.log('handleFileSelect called with', files.length, 'files');
        
        this.isProcessingFiles = true;
        
        try {
            const pdfFiles = Array.from(files).filter(file => file.type === 'application/pdf');
            
            if (pdfFiles.length === 0) {
                this.showStatus('Please select PDF files only.', 'error');
                return;
            }

            this.uploadedFiles = pdfFiles;
            this.updateFileList();
            this.updateGenerateButton();
            this.showStatus(`Uploaded ${pdfFiles.length} PDF file(s) successfully!`, 'success');
            
            console.log('Files processed successfully:', pdfFiles.length);
        } catch (error) {
            console.error('Error in handleFileSelect:', error);
            this.showStatus('Error processing files. Please try again.', 'error');
        } finally {
            this.isProcessingFiles = false;
        }
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
        // Prevent multiple simultaneous executions
        if (this.isGeneratingExcel) {
            console.log('Excel generation already in progress, ignoring request');
            return;
        }
        
        if (this.uploadedFiles.length === 0) {
            this.showStatus('No files to process.', 'error');
            return;
        }

        this.isGeneratingExcel = true;
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
        } finally {
            this.isGeneratingExcel = false;
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
            let links = [];
            
            // Extract both text and annotations from all pages
            for (let i = 1; i <= pdf.numPages; i++) {
                try {
                const page = await pdf.getPage(i);
                    
                    // Get text content
                    const pageTextContent = await page.getTextContent();
                    const pageTextStr = pageTextContent.items.map(item => item.str).join(' ');
                    fullText += pageTextStr + '\n';
                    
                    // Helper function to get text context around a position
                    const getContextText = async (rect) => {
                        const linkTextContent = await page.getTextContent();
                        const items = linkTextContent.items;
                        const contextChars = 200; // Increased context range to better catch specification text
                        let context = '';
                        
                        // Find items near the link vertically
                        const nearbyItems = items.filter(item => {
                            const itemY = item.transform[5];
                            return Math.abs(itemY - rect[1]) < 100; // Increased vertical range to catch more context
                        });
                        
                        // Sort by X position and join
                        context = nearbyItems
                            .sort((a, b) => a.transform[4] - b.transform[4])
                            .map(item => item.str)
                            .join(' ')
                            .trim();
                            
                        // Limit context length
                        if (context.length > contextChars) {
                            const startIdx = Math.max(0, Math.floor(context.length / 2) - contextChars / 2);
                            context = context.substr(startIdx, contextChars);
                        }
                        
                        return context;
                    };

                    // Helper function to categorize links
                    const categorizeLink = (url, contextText) => {
                        const lowerContext = contextText.toLowerCase();
                        const lowerUrl = url.toLowerCase();
                        
                        // BOQ Detail Document - check both context and URL patterns
                        if (lowerContext.includes('boq detail document') || 
                            lowerContext.includes('boq document') ||
                            (lowerContext.includes('boq') && lowerContext.includes('view file')) ||
                            lowerUrl.includes('boqdocument') ||
                            lowerUrl.includes('boq_detail_document')) {
                            return 'BOQ Detail Document';
                        }
                        
                        // Buyer Specification Document - check both context and URL patterns
                        if (lowerContext.includes('buyer specification document') || 
                            lowerContext.includes('buyer spec') ||
                            (lowerContext.includes('specification') && lowerContext.includes('download')) ||
                            lowerUrl.includes('buyer_specification') ||
                            lowerUrl.includes('buyerspecification') ||
                            lowerUrl.includes('spec_document')) {
                            return 'Buyer Specification Document';
                        }
                        
                        // GeM Category Specification - check both context and URL patterns
                        if (lowerContext.includes('as per gem category specification') ||
                            lowerContext.includes('जेम केटेगरी विशिष्टि के अनुसार') ||  // Hindi text
                            lowerContext.includes('gem category spec') ||
                            (lowerContext.includes('as per') && lowerContext.includes('gem')) ||
                            (lowerContext.includes('as per') && lowerContext.includes('category')) ||
                            lowerUrl.includes('category_specification') ||
                            lowerUrl.includes('categoryspec') ||
                            lowerUrl.includes('catalogattrs') ||
                            lowerUrl.includes('catalog_support') ||
                            (lowerUrl.includes('catalog') && lowerUrl.includes('specification'))) {
                            console.log('Found GeM Category Specification by context/URL match');
                            return 'GeM Category Specification';
                        }
                        
                        // Additional check for GeM Category Specification link text
                        const gemSpecText = 'As per GeM Category Specification';
                        if (contextText.includes(gemSpecText) && contextText.indexOf(gemSpecText) < 100) {
                            console.log('Found GeM Category Specification by exact text match');
                            return 'GeM Category Specification';
                        }
                        
                        // If URL contains obvious specification indicators, categorize accordingly
                        if (lowerUrl.includes('tech_spec') || 
                            lowerUrl.includes('technical_specification') ||
                            lowerUrl.includes('specification_document')) {
                            if (lowerUrl.includes('boq')) {
                                return 'BOQ Detail Document';
                            }
                            return 'Buyer Specification Document';
                        }
                        
                        return 'Other';
                    };

                    // Get annotations (which include links)
                    const annotations = await page.getAnnotations();
                    const pageLinks = [];
                    
                    // Process each annotation
                    for (const annot of annotations) {
                        if (annot.subtype === 'Link') {
                            // Extract URL from various possible locations
                            let url = annot.url || annot.A?.URI || 
                                    (annot.action && (annot.action.URI || annot.action.Url));
                            
                            // If this is a text link without URL, try to get URL from the text content
                            if (!url && annot.rect) {
                                const linkText = await getContextText(annot.rect);
                                if (linkText.includes('As per GeM Category Specification') ||
                                    linkText.includes('जेम केटेगरी विशिष्टि के अनुसार')) {
                                    // Look for a URL in nearby annotations
                                    for (const otherAnnot of annotations) {
                                        if (otherAnnot.subtype === 'Link' && 
                                            otherAnnot.url && 
                                            Math.abs(otherAnnot.rect[1] - annot.rect[1]) < 50) {
                                            url = otherAnnot.url;
                                            console.log('Found URL near GeM Category Specification text:', url);
                                            break;
                                        }
                                    }
                                }
                            }
                            
                            if (url) {
                                // Get surrounding text context
                                const contextText = await getContextText(annot.rect);
                                
                                // Categorize and store the link
                                pageLinks.push({
                                    url: url,
                                    type: categorizeLink(url, contextText),
                                    pageNum: i
                                });
                            }
                        }
                    }
                    
                    // Add regex-based URL extraction as fallback
                    const regexTextContent = await page.getTextContent();
                    const regexPageText = regexTextContent.items.map(item => item.str).join(' ');
                    const urlRegex = /(https?:\/\/[^\s)]+)/g;
                    let match;
                    
                    while ((match = urlRegex.exec(regexPageText)) !== null) {
                        const url = match[1];
                        // Check if URL was already found in annotations
                        if (!pageLinks.some(link => link.url === url)) {
                            // Get context around the URL
                            const start = Math.max(0, match.index - 50);
                            const end = Math.min(regexPageText.length, match.index + url.length + 50);
                            const contextText = regexPageText.substring(start, end);
                            
                            pageLinks.push({
                                url: url,
                                type: categorizeLink(url, contextText),
                                pageNum: i
                            });
                        }
                    }
                    
                    links.push(...pageLinks);
                    
                    // Get structured content (might include link elements)
                    const structTree = await page.getStructTree();
                    if (structTree) {
                        const extractLinks = (node) => {
                            if (node.url) {
                                links.push({
                                    url: node.url,
                                    pageNum: i
                                });
                            }
                            if (node.kids) {
                                node.kids.forEach(extractLinks);
                            }
                        };
                        structTree.forEach(extractLinks);
                    }
                    
                } catch (pageError) {
                    console.warn(`Error processing page ${i}:`, pageError);
                    fullText += `[Page ${i} processing error]\n`;
                }
            }
            
            // Store links in the instance for use in extractTechnicalSpec
            this.pdfLinks = links;

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

    // NEW FUNCTION: extractBIDNumber - Extract bid number from PDF content
    extractBIDNumber(text) {
        // Look for bid number patterns like "GEM/2025/B/5918811"
        const bidNumberPatterns = [
            /Bid\s+Number[:\s]*([A-Z]+\/\d{4}\/[A-Z]+\/\d+)/i,
            /बोली\s+क्रमांक[:\s]*([A-Z]+\/\d{4}\/[A-Z]+\/\d+)/i,
            /बिड\s+संख्या[:\s]*([A-Z]+\/\d{4}\/[A-Z]+\/\d+)/i,
            /GEM\/\d{4}\/[A-Z]+\/\d+/gi,
            /([A-Z]+\/\d{4}\/[A-Z]+\/\d+)/g
        ];
        
        for (const pattern of bidNumberPatterns) {
            const match = text.match(pattern);
            if (match) {
                const bidNumber = match[1] || match[0];
                console.log('Found BID Number:', bidNumber);
                return bidNumber;
            }
        }
        
        // Fallback: look for any GEM pattern
        const gemPattern = /GEM\/\d{4}\/[A-Z]+\/\d+/i;
        const gemMatch = text.match(gemPattern);
        if (gemMatch) {
            console.log('Found GEM pattern:', gemMatch[0]);
            return gemMatch[0];
        }
        
        console.log('No BID Number found, using default');
        return 'GEM/2025/B/XXXXXXX'; // Default fallback
    }

    // NEW FUNCTION: calculateBIDEndDate - Calculate end date from start date + validity
    calculateBIDEndDate(startDateStr, validityStr) {
        try {
            // Parse the start date (Dated field)
            if (!startDateStr || startDateStr === 'Not Found') {
                return 'Not Found';
            }
            
            const startDateParts = startDateStr.split('-');
            if (startDateParts.length !== 3) {
                return 'Invalid Start Date';
            }
            
            const startDate = new Date(
                parseInt(startDateParts[2]), // Year
                parseInt(startDateParts[1]) - 1, // Month (0-indexed)
                parseInt(startDateParts[0]) // Day
            );
            
            if (isNaN(startDate.getTime())) {
                return 'Invalid Start Date';
            }
            
            // Parse validity period
            let validityDays = 0;
            if (validityStr && validityStr !== 'Not Found') {
                const match = validityStr.match(/(\d+)/);
                if (match) {
                    validityDays = parseInt(match[1]);
                }
            }
            
            if (validityDays <= 0) {
                return 'Invalid Validity';
            }
            
            // Calculate end date
            const endDate = new Date(startDate);
            endDate.setDate(endDate.getDate() + validityDays);
            
            // Format as DD-MM-YYYY
            const day = String(endDate.getDate()).padStart(2, '0');
            const month = String(endDate.getMonth() + 1).padStart(2, '0');
            const year = endDate.getFullYear();
            
            const endDateStr = `${day}-${month}-${year}`;
            console.log(`Calculated BID End Date: ${startDateStr} + ${validityDays} days = ${endDateStr}`);
            return endDateStr;
            
        } catch (error) {
            console.warn('Error calculating BID End Date:', error);
            return 'Calculation Error';
        }
    }

    // UPDATED: extractDataFromText - Extract bid number from PDF content and add filename column
    extractDataFromText(text, filename) {
        // Extract BID number from PDF content (not filename)
        const bidNumber = this.extractBIDNumber(text);
        
        // Extract data using the SAME regex patterns as our Python script
        const data = {
            'BID Number': bidNumber,
            'Buyer': this.extractBuyer(text),
            'BID Start Date': this.extractBIDStartDate(text),
            'BID Offer Validity': this.extractValidity(text),
            'BID End Date': this.calculateBIDEndDate(this.extractBIDStartDate(text), this.extractValidity(text)),
            'State': this.extractState(text),
            'Total Quantity': this.extractTotalQuantity(text),
            'Item Category': this.extractItemCategory(text),
            'Technical Specification': this.extractTechnicalSpec(text),
            'Filename': filename // NEW: Add filename as last column
        };

        console.log('Extracted data:', data);
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

    // FIXED: extractBIDStartDate - Extract the "Dated" field from top right of PDF
    extractBIDStartDate(text) {
        // Look for the "Dated" field (top right of PDF header) with multiple formats
        const datedPatterns = [
            /Dated\s*:\s*(\d{2}-\d{2}-\d{4})/i,
            /Dated\/\s*दिनांक\s*:\s*(\d{2}-\d{2}-\d{4})/i,
            /दिनांक\s*:\s*(\d{2}-\d{2}-\d{4})/i,
            /Date\s*:\s*(\d{2}-\d{2}-\d{4})/i,
            /Dated\s*\/\s*[^\s]*\s*:\s*(\d{2}-\d{2}-\d{4})/i  // Dated/anything : date
        ];
        
        for (const pattern of datedPatterns) {
            const match = text.match(pattern);
            if (match) {
                const date = match[1];
                console.log('Found BID Start Date (Dated field):', date);
                return date;
            }
        }
        
        // Look for the specific format from the PDF: "Dated/दिनांक : 12-02-2025"
        const specificPattern = /Dated\s*\/\s*[^\s]*\s*[^\s]*\s*:\s*(\d{2}-\d{2}-\d{4})/i;
        const specificMatch = text.match(specificPattern);
        if (specificMatch) {
            const date = specificMatch[1];
            console.log('Found BID Start Date (specific Dated/दिनांक format):', date);
            return date;
        }
        
        // Look for date patterns near "Dated" or "दिनांक" keywords
        const lines = text.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            if (line.includes('Dated') || line.includes('दिनांक')) {
                // Look for DD-MM-YYYY pattern in this line or nearby lines
                const dateMatch = line.match(/(\d{2}-\d{2}-\d{4})/);
                if (dateMatch) {
                    const date = dateMatch[1];
                    console.log('Found BID Start Date near Dated keyword:', date);
                    return date;
                }
                
                // Check next few lines for date
                for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
                    const nextLine = lines[j];
                    const nextDateMatch = nextLine.match(/(\d{2}-\d{2}-\d{4})/);
                    if (nextDateMatch) {
                        const date = nextDateMatch[1];
                        console.log('Found BID Start Date in next line after Dated:', date);
                        return date;
                    }
                }
            }
        }
        
        // Fallback: look for any DD-MM-YYYY format in the first 20 lines (header area)
        for (let i = 0; i < Math.min(20, lines.length); i++) {
            const line = lines[i];
            const dateMatch = line.match(/(\d{2}-\d{2}-\d{4})/);
            if (dateMatch) {
                const date = dateMatch[1];
                console.log('Found fallback date in header area:', date);
                return date;
            }
        }
        
        console.log('No BID Start Date found, using default');
        return '01-03-2025'; // Default fallback
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

    // IMPROVED: extractState - Better state extraction with bid opening date handling
    extractState(text) {
        const lines = text.split('\n');
        
        // Look for state information in the document
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // Look for bid opening date/time line which contains state code
            if (line.includes('Bid Opening Date/Time') || line.includes('बिड खुलने की तारीख')) {
                // Extract the state code (e.g., "03-2025" from the date string)
                const stateMatch = line.match(/(\d{2})-\d{4}\s+\d{2}:\d{2}:\d{2}/);
                if (stateMatch) {
                    const stateCode = stateMatch[1];
                    // Map state code to state name
                    const stateMap = {
                        '01': 'Jammu & Kashmir',
                        '02': 'Himachal Pradesh',
                        '03': 'Punjab',
                        '04': 'Chandigarh',
                        '05': 'Uttarakhand',
                        '06': 'Haryana',
                        '07': 'Delhi',
                        '08': 'Rajasthan',
                        '09': 'Uttar Pradesh',
                        '10': 'Bihar',
                        '11': 'Sikkim',
                        '12': 'Arunachal Pradesh',
                        '13': 'Nagaland',
                        '14': 'Manipur',
                        '15': 'Mizoram',
                        '16': 'Tripura',
                        '17': 'Meghalaya',
                        '18': 'Assam',
                        '19': 'West Bengal',
                        '20': 'Jharkhand',
                        '21': 'Odisha',
                        '22': 'Chhattisgarh',
                        '23': 'Madhya Pradesh',
                        '24': 'Gujarat',
                        '25': 'Daman & Diu',
                        '26': 'Dadra & Nagar Haveli',
                        '27': 'Maharashtra',
                        '28': 'Andhra Pradesh',
                        '29': 'Karnataka',
                        '30': 'Goa',
                        '31': 'Lakshadweep',
                        '32': 'Kerala',
                        '33': 'Tamil Nadu',
                        '34': 'Puducherry',
                        '35': 'Andaman & Nicobar Islands',
                        '36': 'Telangana',
                        '37': 'Ladakh'
                    };
                    
                    if (stateMap[stateCode]) {
                        console.log('Found state from code:', stateCode, stateMap[stateCode]);
                        return stateMap[stateCode];
                    }
                }
            }
            
            // Backup: Look for explicit state mentions
            if (line.includes('State:') || line.includes('राज्य:')) {
                const stateMatch = line.match(/(?:State|राज्य):\s*([^,;\n]+)/);
                if (stateMatch && stateMatch[1].trim()) {
                    const state = stateMatch[1].trim();
                    console.log('Found explicit state mention:', state);
                    return state;
                }
            }
        }
        
        console.log('No state information found');
        return 'Not Found';
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

    // IMPROVED: extractItemCategory - Better extraction of specific product names like "Schneider Electric Compact"
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
                line.includes('Schneider Electric') ||
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
                            cleanLine.includes('Schneider Electric') ||
                            cleanLine.includes('Compact') ||
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
                            cleanLine.includes('Type') ||
                            cleanLine.includes('pieces') ||
                            cleanLine.includes('units') ||
                            cleanLine.includes('nos') ||
                            cleanLine.includes('quantity') ||
                            cleanLine.includes('qty');
                        
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
            
            // Look for Schneider Electric patterns
            const schneiderPatterns = [
                /Schneider\s+Electric\s+[A-Za-z\s]+\(\s*\d+\s*pieces?\s*\)/gi,
                /Schneider\s+Electric\s+[A-Za-z\s]+/gi,
                /Compact\s*\(\s*\d+\s*pieces?\s*\)/gi
            ];
            
            for (const pattern of schneiderPatterns) {
                const matches = text.match(pattern);
                if (matches) {
                    items.push(...matches);
                    foundItems = true;
                    console.log('Found Schneider Electric pattern:', matches);
                }
            }
            
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

    // Extract technical specification URL from PDF text
    extractTechnicalSpec(text) {
        console.log('Extracting technical specification URL...');
        
        // If we have no extracted links, return early
        if (!this.pdfLinks || this.pdfLinks.length === 0) {
            console.log('No links found in PDF');
            return 'No technical specification URLs found';
        }
        
        // Priority order for link types
        const linkPriority = [
            'BOQ Detail Document',
            'Buyer Specification Document',
            'GeM Category Specification'
        ];
        
        // Look for links in priority order
        for (const priority of linkPriority) {
            const link = this.pdfLinks.find(l => l.type === priority);
            if (link) {
                console.log(`Found ${priority} URL:`, link.url);
                return link.url;
            }
        }
        
        console.log('No technical specification URLs found');
        return 'No technical specification URLs found';
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
            'BID End Date': '17-06-2025',
            'State': 'Rajasthan',
            'Total Quantity': Math.floor(Math.random() * 10000) + 100,
            'Item Category': 'Sample Item Category - This would contain the actual extracted text from the PDF',
            'Technical Specification': 'Specification Document; BOQ Detail Document'
        };
    }

    // UPDATED: categorizeBids - Use calculated BID End Date for bid categorization
    categorizeBids(data) {
        const active = [];
        const expired = [];
        const currentDate = new Date();
        
        data.forEach(bid => {
            try {
                // Use calculated BID End Date for categorization
                const endDateStr = bid['BID End Date'];
                if (!endDateStr || endDateStr === 'Not Found' || endDateStr === 'Invalid Start Date' || 
                    endDateStr === 'Invalid Validity' || endDateStr === 'Calculation Error') {
                    // If no valid end date, consider as expired
                    expired.push(bid);
                    return;
                }
                
                // Parse the end date (assuming DD-MM-YYYY format)
                const endDateParts = endDateStr.split('-');
                if (endDateParts.length !== 3) {
                    // Invalid date format, consider as expired
                    expired.push(bid);
                    return;
                }
                
                const endDate = new Date(
                    parseInt(endDateParts[2]), // Year
                    parseInt(endDateParts[1]) - 1, // Month (0-indexed)
                    parseInt(endDateParts[0]) // Day
                );
                
                // Check if end date is valid
                if (isNaN(endDate.getTime())) {
                    expired.push(bid);
                    return;
                }
                
                // Categorize based on end date
                // If end date is today or in the future, consider as active
                // If end date is in the past, consider as expired
                if (endDate >= currentDate) {
                    active.push(bid);
                } else {
                    expired.push(bid);
                }
                
            } catch (error) {
                console.warn('Error categorizing bid:', error);
                // If there's an error, consider as expired
                expired.push(bid);
            }
        });
        
        console.log(`Bids categorized using BID End Date: ${active.length} active, ${expired.length} expired`);
        return { active, expired };
    }

    createExcelFile() {
        try {
            console.log('Creating Excel file with', this.extractedData.length, 'records');
            
            // Categorize bids into active and expired
            const categorizedBids = this.categorizeBids(this.extractedData);
            
            // Create workbook
        const wb = XLSX.utils.book_new();
            
            // Create main data array with headers
            const headers = [
                'BID Number', 'Buyer', 'BID Start Date', 'BID Offer Validity', 
                'BID End Date', 'State', 'Total Quantity', 'Item Category', 'Technical Specification',
                'Filename'
            ];
            
            const data = [headers];
            
            // Add data rows with hyperlink formulas
        this.extractedData.forEach(row => {
                const techSpec = row['Technical Specification'];
                // Check if it's a URL and create a proper Excel HYPERLINK formula
                // Create direct hyperlink without formula
                // Just insert the URL directly - Excel will recognize it as a link
                const techSpecCell = techSpec.startsWith('http') ? techSpec : techSpec || 'Not Found';
                
                data.push([
                    row['BID Number'] || 'Not Found',
                    row['Buyer'] || 'Not Found',
                    row['BID Start Date'] || 'Not Found',
                    row['BID Offer Validity'] || 'Not Found',
                    row['BID End Date'] || 'Not Found',
                    row['State'] || 'Not Found',
                    row['Total Quantity'] || 'Not Found',
                    row['Item Category'] || 'Not Found',
                    techSpecCell, // Use the cell object with formula
                    row['Filename'] || 'Not Found'
                ]);
            });
            
            // Create 'All Bids' worksheet
            const wsAll = XLSX.utils.aoa_to_sheet(data);
            this.formatWorksheet(wsAll, 'All Bids');
            XLSX.utils.book_append_sheet(wb, wsAll, 'All Bids');
            
            // Create 'Active Bids' worksheet
            if (categorizedBids.active.length > 0) {
                const activeData = [headers];
                categorizedBids.active.forEach(row => {
                    const techSpec = row['Technical Specification'];
                    // Just use the direct URL
                    const techSpecCell = techSpec.startsWith('http') ? techSpec : techSpec || 'Not Found';
                    
                    activeData.push([
                    row['BID Number'] || 'Not Found',
                    row['Buyer'] || 'Not Found',
                    row['BID Start Date'] || 'Not Found',
                    row['BID Offer Validity'] || 'Not Found',
                    row['BID End Date'] || 'Not Found',
                    row['State'] || 'Not Found',
                    row['Total Quantity'] || 'Not Found',
                    row['Item Category'] || 'Not Found',
                        techSpecCell,
                        row['Filename'] || 'Not Found'
                    ]);
                });
                const wsActive = XLSX.utils.aoa_to_sheet(activeData);
                this.formatWorksheet(wsActive, 'Active Bids');
                XLSX.utils.book_append_sheet(wb, wsActive, 'Active Bids');
            }
            
            // Create 'Expired Bids' worksheet
            if (categorizedBids.expired.length > 0) {
                const expiredData = [headers];
                categorizedBids.expired.forEach(row => {
                    const techSpec = row['Technical Specification'];
                    // For hyperlinks, just use the raw Excel formula string
                    // Just use the direct URL
                    const techSpecCell = techSpec.startsWith('http') ? techSpec : techSpec || 'Not Found';
                    
                    expiredData.push([
                    row['BID Number'] || 'Not Found',
                    row['Buyer'] || 'Not Found',
                    row['BID Start Date'] || 'Not Found',
                    row['BID Offer Validity'] || 'Not Found',
                    row['BID End Date'] || 'Not Found',
                    row['State'] || 'Not Found',
                    row['Total Quantity'] || 'Not Found',
                    row['Item Category'] || 'Not Found',
                        techSpecCell,
                        row['Filename'] || 'Not Found'
                    ]);
                });
                const wsExpired = XLSX.utils.aoa_to_sheet(expiredData);
                this.formatWorksheet(wsExpired, 'Expired Bids');
                XLSX.utils.book_append_sheet(wb, wsExpired, 'Expired Bids');
            }
            
            // Create summary worksheet
            const summaryData = [
                ['Summary Statistics'],
                [''],
                ['Total Bids', this.extractedData.length],
                ['Active Bids', categorizedBids.active.length],
                ['Expired Bids', categorizedBids.expired.length],
                [''],
                ['Generated on', new Date().toLocaleString()],
                ['Total Files Processed', this.uploadedFiles.length]
            ];
            
            const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
            this.formatWorksheet(wsSummary, 'Summary');
            XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

        // Generate Excel file
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // Create download link
        const url = URL.createObjectURL(blob);
        const downloadBtn = document.getElementById('downloadBtn');
        downloadBtn.href = url;
            downloadBtn.download = `GeM_Bidding_Data_${new Date().toISOString().split('T')[0]}.xlsx`;
            
            console.log('Excel file created successfully with', wb.SheetNames.length, 'worksheets');
            return true;
            
        } catch (error) {
            console.error('Error creating Excel file:', error);
            this.showStatus('Error creating Excel file: ' + error.message, 'error');
            return false;
        }
    }

    // Format worksheet with consistent styling and proper hyperlinks
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
        
        // Process each cell for formatting
        dataRows.forEach(row => {
            Object.keys(ws).forEach(cell => {
                if (cell[0] === '!') return;
                
                const cellData = ws[cell];
                const col = cell.replace(/\d/g, '');
                const colIndex = XLSX.utils.decode_col(col);
                
                // Calculate column width - skip formatting for URLs
                if (!(typeof cellData.v === 'string' && cellData.v.startsWith('http'))) {
                    const cellValue = cellData.v || '';
                    const valueLength = String(cellValue).length;
                colWidths[colIndex] = Math.max(colWidths[colIndex] || 0, 
                    Math.min(valueLength, maxColWidth));
                }
            });
        });

        // Set column widths
        ws['!cols'] = colWidths.map(width => ({ 
            width: Math.min(width + 3, maxColWidth + 5),
            wrapText: true
        }));

        // Set row heights
        ws['!rows'] = [];
        const maxRow = Math.max(...dataRows);
        for (let i = 0; i <= maxRow; i++) {
            ws['!rows'][i] = { hpt: 25 };
        }

        // Style header row
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
        if (data['BID Number'] === 'GEM/2025/B/XXXXXXX') {
            warnings.push('BID Number may be using default value');
        }
        if (data['Buyer'] === 'Ministry of Defence') {
            warnings.push('Buyer information may be using default value');
        }
        if (data['BID Start Date'] === '01-03-2025') {
            warnings.push('BID Start Date may be using default value');
        }
        if (data['BID End Date'] === 'Not Found' || data['BID End Date'] === 'Invalid Start Date' || 
            data['BID End Date'] === 'Invalid Validity' || data['BID End Date'] === 'Calculation Error') {
            warnings.push('BID End Date calculation failed');
            isValid = false;
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