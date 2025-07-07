const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const officegen = require('officegen');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Graceful shutdown handling
let server;
let isShuttingDown = false;

// Suppress warnings
const originalConsoleWarn = console.warn;
console.warn = function(message) {
    if (typeof message === 'string' && 
        (message.includes('TT: undefined function') || 
         message.includes('ExperimentalWarning'))) {
        return;
    }
    originalConsoleWarn.apply(console, arguments);
};

// Middleware
app.use(cors());
app.use(express.static('public'));
app.use(express.json({ limit: '15mb' }));
app.use(express.urlencoded({ extended: true, limit: '15mb' }));

// Request timeout middleware
app.use((req, res, next) => {
    res.setTimeout(30000, () => {
        if (!res.headersSent) {
            res.status(408).json({ error: 'Request timeout' });
        }
    });
    next();
});

// Health check - must be before other routes
app.get('/health', (req, res) => {
    res.status(200).json({ 
        status: 'OK', 
        timestamp: new Date().toISOString(),
        uptime: Math.floor(process.uptime()),
        memory: process.memoryUsage()
    });
});

// Create directories
const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

try {
    if (!fs.existsSync(uploadsDir)) {
        fs.mkdirSync(uploadsDir, { recursive: true });
    }
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }
} catch (error) {
    console.error('Failed to create directories:', error);
}

// Configure multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadsDir);
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});

const upload = multer({
    storage: storage,
    limits: {
        fileSize: 10 * 1024 * 1024, // 10MB
        files: 1
    },
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/pdf') {
            cb(null, true);
        } else {
            cb(new Error('Only PDF files are allowed'));
        }
    }
});

// Utility functions
function cleanText(text) {
    if (!text || typeof text !== 'string') {
        return 'No readable text found in this PDF document.';
    }
    
    return text
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/\n{3,}/g, '\n\n')
        .replace(/[ \t]{2,}/g, ' ')
        .replace(/^\s+|\s+$/gm, '')
        .trim() || 'Document appears to be empty or contains only images.';
}

function safeFileCleanup(filePath) {
    if (filePath && fs.existsSync(filePath)) {
        try {
            fs.unlinkSync(filePath);
            console.log('File cleaned up:', path.basename(filePath));
        } catch (error) {
            console.error('Cleanup failed:', error.message);
        }
    }
}

// Routes
app.get('/', (req, res) => {
    if (isShuttingDown) {
        return res.status(503).json({ error: 'Server is shutting down' });
    }
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/convert', (req, res) => {
    if (isShuttingDown) {
        return res.status(503).json({ error: 'Server is shutting down' });
    }

    upload.single('pdfFile')(req, res, async (uploadError) => {
        let tempFilePath = null;
        
        try {
            if (uploadError) {
                throw uploadError;
            }

            if (!req.file) {
                throw new Error('No file uploaded');
            }

            tempFilePath = req.file.path;
            const outputFileName = `converted-${Date.now()}.docx`;
            const outputPath = path.join(outputDir, outputFileName);

            console.log('Converting:', req.file.originalname, `(${(req.file.size / 1024).toFixed(1)}KB)`);

            // Extract text from PDF
            let extractedText;
            try {
                const pdfBuffer = fs.readFileSync(tempFilePath);
                const pdfData = await pdfParse(pdfBuffer, {
                    normalizeWhitespace: true,
                    disableCombineTextItems: false,
                    max: 0 // No limit on pages
                });
                extractedText = cleanText(pdfData.text);
            } catch (pdfError) {
                console.error('PDF parsing error:', pdfError.message);
                extractedText = `Error extracting text from PDF: ${pdfError.message}\n\nThis may be an image-based PDF or the file may be corrupted.`;
            }

            // Create Word document
            const docx = officegen({
                type: 'docx',
                orientation: 'portrait',
                margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
            });

            // Handle officegen errors
            docx.on('error', (err) => {
                console.error('Officegen error:', err);
            });

            // Document properties
            docx.setDocTitle('PDF Conversion Result');
            docx.setDocSubject(`Converted from ${req.file.originalname}`);

            // Add content
            const header = docx.createP({ align: 'center' });
            header.addText('PDF to Word Conversion', { 
                bold: true, 
                font_size: 16,
                color: '2F5496'
            });

            const fileInfo = docx.createP({ align: 'center' });
            fileInfo.addText(`Source: ${req.file.originalname}`, { 
                italic: true, 
                font_size: 10,
                color: '808080'
            });

            docx.createP().addLineBreak();

            // Add text content
            const paragraphs = extractedText.split('\n\n');
            paragraphs.forEach(paragraph => {
                if (paragraph.trim()) {
                    const p = docx.createP();
                    p.addText(paragraph.trim(), { font_size: 11 });
                } else {
                    docx.createP().addLineBreak();
                }
            });

            // Save document
            const output = fs.createWriteStream(outputPath);
            
            output.on('error', (err) => {
                throw new Error(`Failed to write document: ${err.message}`);
            });

            output.on('close', () => {
                console.log('Conversion completed successfully');
                res.json({
                    success: true,
                    message: 'PDF converted successfully',
                    downloadUrl: `/download/${outputFileName}`,
                    originalName: req.file.originalname
                });
            });

            docx.generate(output);

        } catch (error) {
            console.error('Conversion error:', error.message);
            
            let errorMessage = 'Failed to convert PDF';
            if (error.code === 'LIMIT_FILE_SIZE') {
                errorMessage = 'File too large. Maximum size is 10MB';
            } else if (error.message.includes('Only PDF files')) {
                errorMessage = 'Please select a valid PDF file';
            } else if (error.message.includes('No file uploaded')) {
                errorMessage = 'No file was uploaded';
            }

            res.status(400).json({ 
                error: errorMessage,
                details: process.env.NODE_ENV === 'development' ? error.message : undefined
            });
        } finally {
            // Cleanup
            if (tempFilePath) {
                setTimeout(() => safeFileCleanup(tempFilePath), 1000);
            }
        }
    });
});

app.get('/download/:filename', (req, res) => {
    if (isShuttingDown) {
        return res.status(503).json({ error: 'Server is shutting down' });
    }

    const filename = req.params.filename;
    
    // Validate filename
    if (!/^converted-\d+\.docx$/.test(filename)) {
        return res.status(400).json({ error: 'Invalid filename' });
    }

    const filePath = path.join(outputDir, filename);
    
    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'File not found or expired' });
    }

    const downloadName = `document-${Date.now()}.docx`;
    
    res.download(filePath, downloadName, (err) => {
        if (err) {
            console.error('Download error:', err.message);
            if (!res.headersSent) {
                res.status(500).json({ error: 'Download failed' });
            }
        } else {
            console.log('File downloaded:', downloadName);
            // Cleanup after 2 minutes
            setTimeout(() => safeFileCleanup(filePath), 120000);
        }
    });
});

// Error handling
app.use((error, req, res, next) => {
    console.error('Express error:', error.message);
    
    if (res.headersSent) {
        return next(error);
    }

    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({ error: 'File too large (max 10MB)' });
        }
    }
    
    res.status(500).json({ error: 'Internal server error' });
});

// 404 handler
app.use((req, res) => {
    res.status(404).json({ error: 'Not found' });
});

// Graceful shutdown
function gracefulShutdown(signal) {
    console.log(`Received ${signal}. Starting graceful shutdown...`);
    isShuttingDown = true;
    
    if (server) {
        server.close((err) => {
            if (err) {
                console.error('Error during server shutdown:', err);
                process.exit(1);
            }
            console.log('Server closed gracefully');
            process.exit(0);
        });
        
        // Force shutdown after 10 seconds
        setTimeout(() => {
            console.log('Forcing shutdown');
            process.exit(1);
        }, 10000);
    } else {
        process.exit(0);
    }
}

// Handle shutdown signals
process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
process.on('SIGINT', () => gracefulShutdown('SIGINT'));
process.on('SIGHUP', () => gracefulShutdown('SIGHUP'));

// Handle uncaught exceptions
process.on('uncaughtException', (error) => {
    console.error('Uncaught Exception:', error);
    gracefulShutdown('UNCAUGHT_EXCEPTION');
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

// Start server
server = app.listen(PORT, '0.0.0.0', () => {
    console.log(`PDF to Word Converter running on port ${PORT}`);
    console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
    console.log(`Memory usage: ${Math.round(process.memoryUsage().heapUsed / 1024 / 1024)}MB`);
});
