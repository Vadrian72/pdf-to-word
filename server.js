const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const officegen = require('officegen');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Suppress officegen warnings
process.on('warning', (warning) => {
    if (warning.name === 'ExperimentalWarning' || 
        warning.message.includes('TT: undefined function')) {
        return; // Ignore these warnings
    }
    console.warn(warning);
});

// Middleware
app.use(cors());
app.use(express.static('public'));
app.use(express.json());

// Ensure uploads directory exists
const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});

const upload = multer({
    storage: storage,
    limits: {
        fileSize: 10 * 1024 * 1024, // 10MB limit
    },
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/pdf') {
            cb(null, true);
        } else {
            cb(new Error('Only PDF files are allowed!'), false);
        }
    }
});

// Helper function to clean text
function cleanText(text) {
    if (!text) return 'No text content found in PDF';
    
    // Remove excessive whitespace and normalize line breaks
    return text
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/\n{3,}/g, '\n\n')
        .replace(/[ \t]{2,}/g, ' ')
        .trim();
}

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/convert', upload.single('pdfFile'), async (req, res) => {
    let tempFilePath = null;
    
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        tempFilePath = req.file.path;
        const outputFileName = `converted-${Date.now()}.docx`;
        const outputPath = path.join(outputDir, outputFileName);

        console.log('Processing PDF:', req.file.originalname);

        // Extract text from PDF with error handling
        let extractedText = '';
        try {
            const pdfBuffer = fs.readFileSync(tempFilePath);
            const pdfData = await pdfParse(pdfBuffer, {
                normalizeWhitespace: true,
                disableCombineTextItems: false
            });
            extractedText = cleanText(pdfData.text);
        } catch (pdfError) {
            console.error('PDF parsing error:', pdfError.message);
            extractedText = 'Error: Could not extract text from PDF. The file may be image-based or corrupted.';
        }

        // Create Word document with better formatting
        const docx = officegen({
            type: 'docx',
            orientation: 'portrait',
            margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
        });

        // Set document properties
        docx.setDocTitle('Converted from PDF');
        docx.setDocSubject('PDF to Word Conversion');
        docx.setDocKeywords('pdf, word, conversion');

        // Add content
        const title = docx.createP({ align: 'center' });
        title.addText('Document Converted from PDF', { 
            bold: true, 
            font_size: 18,
            color: '2F5496'
        });
        
        const subtitle = docx.createP({ align: 'center' });
        subtitle.addText(`Original file: ${req.file.originalname}`, { 
            italic: true, 
            font_size: 12,
            color: '7F7F7F'
        });
        
        // Add some spacing
        docx.createP().addLineBreak();
        
        // Add extracted text with proper formatting
        const textLines = extractedText.split('\n');
        textLines.forEach(line => {
            if (line.trim()) {
                const p = docx.createP();
                p.addText(line.trim(), { font_size: 11 });
            } else {
                docx.createP().addLineBreak();
            }
        });

        // Save the document
        return new Promise((resolve, reject) => {
            const output = fs.createWriteStream(outputPath);
            
            output.on('error', (err) => {
                console.error('File write error:', err);
                reject(new Error('Failed to create Word document'));
            });

            output.on('close', () => {
                console.log('Word document created successfully');
                resolve({
                    success: true,
                    message: 'PDF converted successfully',
                    downloadUrl: `/download/${outputFileName}`,
                    originalName: req.file.originalname
                });
            });

            docx.generate(output);
        }).then(result => {
            res.json(result);
        }).catch(error => {
            throw error;
        });

    } catch (error) {
        console.error('Conversion error:', error);
        res.status(500).json({ 
            error: 'Failed to convert PDF to Word',
            details: error.message 
        });
    } finally {
        // Clean up uploaded file
        if (tempFilePath && fs.existsSync(tempFilePath)) {
            try {
                fs.unlinkSync(tempFilePath);
                console.log('Temporary file cleaned up');
            } catch (cleanupError) {
                console.error('Cleanup error:', cleanupError);
            }
        }
    }
});

app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(outputDir, filename);
    
    if (fs.existsSync(filePath)) {
        const downloadName = `converted-document-${Date.now()}.docx`;
        
        res.download(filePath, downloadName, (err) => {
            if (err) {
                console.error('Download error:', err);
                res.status(500).json({ error: 'Error downloading file' });
            } else {
                console.log('File downloaded successfully');
                // Clean up file after download with longer delay
                setTimeout(() => {
                    if (fs.existsSync(filePath)) {
                        try {
                            fs.unlinkSync(filePath);
                            console.log('Output file cleaned up');
                        } catch (cleanupError) {
                            console.error('Output cleanup error:', cleanupError);
                        }
                    }
                }, 120000); // Delete after 2 minutes
            }
        });
    } else {
        res.status(404).json({ error: 'File not found' });
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'OK', 
        timestamp: new Date().toISOString(),
        uptime: process.uptime()
    });
});

// Error handling middleware
app.use((error, req, res, next) => {
    console.error('Express error:', error);
    
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({ 
                error: 'File too large. Maximum size is 10MB.' 
            });
        }
        if (error.code === 'LIMIT_UNEXPECTED_FILE') {
            return res.status(400).json({ 
                error: 'Unexpected file field. Please select a PDF file.' 
            });
        }
    }
    
    res.status(500).json({ 
        error: 'Internal server error',
        message: error.message 
    });
});

// 404 handler
app.use((req, res) => {
    res.status(404).json({ error: 'Endpoint not found' });
});

app.listen(PORT, () => {
    console.log(`PDF to Word Converter running on port ${PORT}`);
    console.log(`Health check available at: http://localhost:${PORT}/health`);
});
