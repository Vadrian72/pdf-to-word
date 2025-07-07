document.addEventListener('DOMContentLoaded', function() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('pdfFile');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const removeFile = document.getElementById('removeFile');
    const convertBtn = document.getElementById('convertBtn');
    const progressSection = document.getElementById('progressSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const downloadSection = document.getElementById('downloadSection');
    const downloadBtn = document.getElementById('downloadBtn');
    const errorMessage = document.getElementById('errorMessage');
    const errorText = document.getElementById('errorText');

    let selectedFile = null;
    let downloadUrl = null;

    // Upload area click handler
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // File input change handler
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            handleFile(file);
        }
    });

    // Drag and drop handlers
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (file.type === 'application/pdf') {
                handleFile(file);
            } else {
                showError('Please select only PDF files');
            }
        }
    });

    // Remove file handler
    removeFile.addEventListener('click', function() {
        selectedFile = null;
        fileInput.value = '';
        fileInfo.style.display = 'none';
        uploadArea.style.display = 'block';
        convertBtn.disabled = true;
        hideError();
        hideProgress();
        hideDownload();
    });

    // Convert button handler
    convertBtn.addEventListener('click', function() {
        if (selectedFile) {
            convertFile();
        }
    });

    // Download button handler
    downloadBtn.addEventListener('click', function() {
        if (downloadUrl) {
            window.location.href = downloadUrl;
        }
    });

    function handleFile(file) {
        if (file.type !== 'application/pdf') {
            showError('Please select only PDF files');
            return;
        }

        if (file.size > 10 * 1024 * 1024) { // 10MB
            showError('File is too large. Maximum size allowed is 10MB');
            return;
        }

        selectedFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        
        uploadArea.style.display = 'none';
        fileInfo.style.display = 'flex';
        convertBtn.disabled = false;
        
        hideError();
        hideProgress();
        hideDownload();
    }

    function convertFile() {
        const formData = new FormData();
        formData.append('pdfFile', selectedFile);

        convertBtn.disabled = true;
        showProgress();
        hideError();
        hideDownload();

        // Simulate progress
        let progress = 0;
        const progressInterval = setInterval(() => {
            progress += Math.random() * 15;
            if (progress > 90) {
                clearInterval(progressInterval);
                progress = 90;
            }
            updateProgress(progress);
        }, 200);

        fetch('/convert', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            clearInterval(progressInterval);
            
            if (data.success) {
                updateProgress(100);
                downloadUrl = data.downloadUrl;
                
                setTimeout(() => {
                    hideProgress();
                    showDownload();
                }, 1000);
            } else {
                hideProgress();
                showError(data.error || 'Error converting file');
            }
        })
        .catch(error => {
            clearInterval(progressInterval);
            hideProgress();
            showError('Error connecting to server');
            console.error('Error:', error);
        })
        .finally(() => {
            convertBtn.disabled = false;
        });
    }

    function showProgress() {
        progressSection.style.display = 'block';
        updateProgress(0);
    }

    function hideProgress() {
        progressSection.style.display = 'none';
    }

    function updateProgress(percent) {
        progressFill.style.width = percent + '%';
        progressText.textContent = `Processing... ${Math.round(percent)}%`;
    }

    function showDownload() {
        downloadSection.style.display = 'block';
    }

    function hideDownload() {
        downloadSection.style.display = 'none';
    }

    function showError(message) {
        errorText.textContent = message;
        errorMessage.style.display = 'block';
    }

    function hideError() {
        errorMessage.style.display = 'none';
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
});
