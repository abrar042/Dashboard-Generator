const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const path = require('path');
const cors = require('cors');
const { exec } = require('child_process');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Storage setup
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = './uploads';
        fs.ensureDirSync(uploadDir);
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        cb(null, `data_${Date.now()}_${file.originalname}`);
    }
});

const upload = multer({ 
    storage: storage,
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.xlsx', '.xls', '.csv', '.primavera', '.xml'];
        const ext = path.extname(file.originalname).toLowerCase();
        if (allowedTypes.includes(ext) || file.originalname.includes('primavera')) {
            cb(null, true);
        } else {
            cb(new Error('Unsupported file type'), false);
        }
    }
});

// Data cleaning function - AUTO CLEANS ANY FILE
function autoCleanData(data) {
    // Remove empty rows
    data = data.filter(row => Object.values(row).some(val => val !== '' && val !== undefined && val !== null));
    
    // Standardize column names
    const cleanColumns = {};
    Object.keys(data[0] || {}).forEach(key => {
        let clean = key.toLowerCase()
            .replace(/[^a-z0-9]/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_|_$/g, '');
        cleanColumns[key] = clean || 'column';
    });
    
    data = data.map(row => {
        const newRow = {};
        Object.keys(row).forEach(key => {
            const newKey = cleanColumns[key] || key;
            let value = row[key];
            // Auto convert types
            if (typeof value === 'string') {
                const num = parseFloat(value.replace(/[$,%]/g, ''));
                if (!isNaN(num)) {
                    newRow[newKey] = num;
                } else if (value.toLowerCase() === 'true' || value.toLowerCase() === 'false') {
                    newRow[newKey] = value.toLowerCase() === 'true';
                } else if (value.toLowerCase() === 'null' || value.toLowerCase() === 'undefined') {
                    newRow[newKey] = null;
                } else {
                    newRow[newKey] = value.trim();
                }
            } else {
                newRow[newKey] = value;
            }
        });
        return newRow;
    });
    
    return data;
}

// Generate Dashboard HTML - BEAUTIFUL INTERACTIVE DASHBOARD
function generateDashboardHTML(data, filename) {
    const columns = Object.keys(data[0] || {});
    const numericCols = columns.filter(col => 
        typeof data[0]?.[col] === 'number' || !isNaN(parseFloat(data[0]?.[col]))
    );
    
    // Auto generate KPIs
    const firstNumeric = numericCols[0] || columns[0];
    const total = data.reduce((sum, row) => sum + (parseFloat(row[firstNumeric]) || 0), 0);
    const avg = (total / data.length).toFixed(2);
    const max = Math.max(...data.map(row => parseFloat(row[firstNumeric]) || 0));
    const min = Math.min(...data.map(row => parseFloat(row[firstNumeric]) || 0));
    
    return `
    <!DOCTYPE html>
    <html>
    <head>
        <title>Auto Dashboard - ${filename}</title>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background: #f0f4f8;
                padding: 20px;
            }
            .container { max-width: 1400px; margin: 0 auto; }
            .header {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 30px;
                border-radius: 15px;
                margin-bottom: 30px;
            }
            .kpi-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 20px;
                margin-bottom: 30px;
            }
            .kpi-card {
                background: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            .kpi-card h3 { color: #666; font-size: 14px; margin-bottom: 10px; }
            .kpi-card .value { font-size: 28px; font-weight: bold; color: #333; }
            .chart-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
                gap: 20px;
                margin-bottom: 30px;
            }
            .chart-card {
                background: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            .chart-card h3 { color: #333; margin-bottom: 15px; }
            .table-container {
                background: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                overflow-x: auto;
                max-height: 400px;
                overflow-y: auto;
            }
            table { width: 100%; border-collapse: collapse; }
            th {
                background: #667eea;
                color: white;
                padding: 12px;
                text-align: left;
                position: sticky;
                top: 0;
                z-index: 10;
            }
            td { padding: 10px; border-bottom: 1px solid #eee; }
            tr:hover { background: #f8f9fa; }
            .badge {
                background: #4CAF50;
                color: white;
                padding: 5px 10px;
                border-radius: 20px;
                font-size: 12px;
            }
            @media (max-width: 768px) {
                .chart-grid { grid-template-columns: 1fr; }
                .kpi-grid { grid-template-columns: 1fr 1fr; }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>📊 Auto Dashboard - ${filename}</h1>
                <p>Data cleaned and analyzed automatically | ${data.length} records loaded</p>
                <p><span class="badge">Auto-Generated</span> <span class="badge">${columns.length} columns</span></p>
            </div>
            
            <div class="kpi-grid">
                <div class="kpi-card">
                    <h3>Total ${firstNumeric}</h3>
                    <div class="value">${total.toLocaleString()}</div>
                </div>
                <div class="kpi-card">
                    <h3>Average ${firstNumeric}</h3>
                    <div class="value">${avg}</div>
                </div>
                <div class="kpi-card">
                    <h3>Maximum ${firstNumeric}</h3>
                    <div class="value">${max.toLocaleString()}</div>
                </div>
                <div class="kpi-card">
                    <h3>Minimum ${firstNumeric}</h3>
                    <div class="value">${min.toLocaleString()}</div>
                </div>
            </div>
            
            <div class="chart-grid">
                <div class="chart-card">
                    <h3>Top 10 ${firstNumeric} Distribution</h3>
                    <canvas id="chart1"></canvas>
                </div>
                <div class="chart-card">
                    <h3>Column Distribution</h3>
                    <canvas id="chart2"></canvas>
                </div>
            </div>
            
            <div class="table-container">
                <h3 style="margin-bottom:15px;">📋 Cleaned Data Preview</h3>
                <table>
                    <thead>
                        <tr>${columns.map(col => `<th>${col}</th>`).join('')}</tr>
                    </thead>
                    <tbody>
                        ${data.slice(0, 50).map(row => `
                            <tr>${columns.map(col => `<td>${row[col] !== undefined ? row[col] : ''}</td>`).join('')}</tr>
                        `).join('')}
                    </tbody>
                </table>
                ${data.length > 50 ? `<p style="margin-top:10px;color:#666;">Showing 50 of ${data.length} records</p>` : ''}
            </div>
        </div>
        
        <script>
            const data = ${JSON.stringify(data.slice(0, 20))};
            const columns = ${JSON.stringify(columns)};
            const numericCols = ${JSON.stringify(numericCols)};
            
            // Chart 1
            const ctx1 = document.getElementById('chart1').getContext('2d');
            const labels1 = data.map((d, i) => 'Item ' + (i+1));
            const values1 = data.map(d => parseFloat(d['${firstNumeric}']) || 0);
            
            new Chart(ctx1, {
                type: 'bar',
                data: {
                    labels: labels1,
                    datasets: [{
                        label: '${firstNumeric}',
                        data: values1,
                        backgroundColor: 'rgba(102, 126, 234, 0.6)',
                        borderColor: '#667eea',
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    plugins: {
                        legend: { display: false }
                    }
                }
            });
            
            // Chart 2 - Column distribution
            const ctx2 = document.getElementById('chart2').getContext('2d');
            const colTypes = columns.map(col => {
                const firstVal = data[0]?.[col];
                if (typeof firstVal === 'number') return 'Number';
                if (firstVal instanceof Date) return 'Date';
                if (typeof firstVal === 'boolean') return 'Boolean';
                return 'Text';
            });
            
            const typeCount = {};
            colTypes.forEach(type => { typeCount[type] = (typeCount[type] || 0) + 1; });
            
            new Chart(ctx2, {
                type: 'doughnut',
                data: {
                    labels: Object.keys(typeCount),
                    datasets: [{
                        data: Object.values(typeCount),
                        backgroundColor: ['#667eea', '#764ba2', '#f093fb', '#4facfe']
                    }]
                },
                options: {
                    responsive: true,
                    plugins: {
                        legend: { position: 'bottom' }
                    }
                }
            });
        </script>
    </body>
    </html>
    `;
}

// API endpoint - UPLOAD AND AUTO GENERATE
app.post('/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const filePath = req.file.path;
        const ext = path.extname(req.file.originalname).toLowerCase();
        let data = [];
        
        // Read file based on type
        if (ext === '.csv') {
            const csv = require('csv-parser');
            const fs = require('fs');
            const results = [];
            await new Promise((resolve, reject) => {
                fs.createReadStream(filePath)
                    .pipe(csv())
                    .on('data', (row) => results.push(row))
                    .on('end', resolve)
                    .on('error', reject);
            });
            data = results;
        } else {
            const workbook = XLSX.readFile(filePath);
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            data = XLSX.utils.sheet_to_json(firstSheet);
        }
        
        if (!data || data.length === 0) {
            return res.status(400).json({ error: 'No data found in file' });
        }
        
        // AUTO CLEAN DATA
        const cleanedData = autoCleanData(data);
        
        // Generate dashboard
        const filename = path.basename(req.file.originalname, ext);
        const dashboardHTML = generateDashboardHTML(cleanedData, filename);
        
        // Save dashboard
        const dashboardDir = './dashboards';
        fs.ensureDirSync(dashboardDir);
        const dashboardFile = path.join(dashboardDir, `dashboard_${Date.now()}.html`);
        fs.writeFileSync(dashboardFile, dashboardHTML);
        
        // Also save cleaned data as Excel
        const cleanWorkbook = XLSX.utils.book_new();
        const cleanSheet = XLSX.utils.json_to_sheet(cleanedData);
        XLSX.utils.book_append_sheet(cleanWorkbook, cleanSheet, 'Cleaned Data');
        const cleanFile = path.join(dashboardDir, `cleaned_${Date.now()}.xlsx`);
        XLSX.writeFile(cleanWorkbook, cleanFile);
        
        res.json({
            success: true,
            message: 'Dashboard generated successfully!',
            records: cleanedData.length,
            columns: Object.keys(cleanedData[0] || {}),
            dashboardUrl: `/dashboard/${path.basename(dashboardFile)}`,
            cleanedFile: `/download/${path.basename(cleanFile)}`,
            preview: cleanedData.slice(0, 5)
        });
        
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

// Serve dashboards
app.get('/dashboard/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'dashboards', req.params.filename);
    if (fs.existsSync(filePath)) {
        res.sendFile(path.resolve(filePath));
    } else {
        res.status(404).send('Dashboard not found');
    }
});

// Download cleaned data
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'dashboards', req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).send('File not found');
    }
});

// Serve upload page
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Auto Dashboard Generator</title>
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body {
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    min-height: 100vh;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    padding: 20px;
                }
                .card {
                    background: white;
                    padding: 50px;
                    border-radius: 20px;
                    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                    max-width: 600px;
                    width: 100%;
                    text-align: center;
                }
                h1 { color: #333; margin-bottom: 10px; }
                p { color: #666; margin-bottom: 30px; }
                .dropzone {
                    border: 3px dashed #667eea;
                    border-radius: 15px;
                    padding: 50px;
                    margin: 20px 0;
                    cursor: pointer;
                    transition: all 0.3s;
                }
                .dropzone:hover { background: #f8f9ff; }
                .dropzone input { display: none; }
                .dropzone .icon { font-size: 50px; margin-bottom: 15px; }
                button {
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    border: none;
                    padding: 15px 40px;
                    border-radius: 50px;
                    font-size: 18px;
                    cursor: pointer;
                    transition: transform 0.3s;
                    margin-top: 20px;
                }
                button:hover { transform: scale(1.05); }
                #status { margin-top: 20px; color: #333; }
                .file-info {
                    background: #f0f4f8;
                    padding: 10px;
                    border-radius: 10px;
                    margin-top: 10px;
                    display: none;
                }
                .loader {
                    border: 4px solid #f3f3f3;
                    border-top: 4px solid #667eea;
                    border-radius: 50%;
                    width: 40px;
                    height: 40px;
                    animation: spin 1s linear infinite;
                    margin: 20px auto;
                    display: none;
                }
                @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
                .supported-formats {
                    display: flex;
                    gap: 10px;
                    justify-content: center;
                    flex-wrap: wrap;
                    margin: 15px 0;
                }
                .format-badge {
                    background: #e8ecf1;
                    padding: 5px 15px;
                    border-radius: 20px;
                    font-size: 12px;
                    color: #555;
                }
            </style>
        </head>
        <body>
            <div class="card">
                <h1>🚀 Auto Dashboard Generator</h1>
                <p>Upload any file, get an interactive dashboard instantly!</p>
                <div class="supported-formats">
                    <span class="format-badge">📊 Excel (.xlsx, .xls)</span>
                    <span class="format-badge">📄 CSV</span>
                    <span class="format-badge">📋 Primavera</span>
                    <span class="format-badge">📁 XML</span>
                </div>
                <div class="dropzone" id="dropzone">
                    <div class="icon">📤</div>
                    <p><strong>Drag & Drop</strong> your file here</p>
                    <p style="font-size:14px;color:#999;">or click to browse</p>
                    <input type="file" id="fileInput" accept=".xlsx,.xls,.csv,.xml,.primavera">
                </div>
                <div id="fileInfo" class="file-info"></div>
                <button onclick="uploadFile()">⚡ Generate Dashboard</button>
                <div id="loader" class="loader"></div>
                <div id="status"></div>
            </div>

            <script>
                const dropzone = document.getElementById('dropzone');
                const fileInput = document.getElementById('fileInput');
                let selectedFile = null;

                dropzone.addEventListener('click', () => fileInput.click());
                dropzone.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    dropzone.style.borderColor = '#764ba2';
                });
                dropzone.addEventListener('dragleave', () => {
                    dropzone.style.borderColor = '#667eea';
                });
                dropzone.addEventListener('drop', (e) => {
                    e.preventDefault();
                    if (e.dataTransfer.files.length) {
                        selectedFile = e.dataTransfer.files[0];
                        handleFile(selectedFile);
                    }
                });
                fileInput.addEventListener('change', (e) => {
                    if (e.target.files.length) {
                        selectedFile = e.target.files[0];
                        handleFile(selectedFile);
                    }
                });

                function handleFile(file) {
                    const info = document.getElementById('fileInfo');
                    info.style.display = 'block';
                    info.innerHTML = \`
                        <strong>\${file.name}</strong><br>
                        Size: \${(file.size / 1024).toFixed(1)} KB
                    \`;
                    document.getElementById('status').innerHTML = '✅ Ready to generate!';
                }

                async function uploadFile() {
                    if (!selectedFile) {
                        document.getElementById('status').innerHTML = '⚠️ Please select a file first!';
                        return;
                    }

                    const formData = new FormData();
                    formData.append('file', selectedFile);

                    document.getElementById('loader').style.display = 'block';
                    document.getElementById('status').innerHTML = '🔄 Processing your data... This may take a moment';

                    try {
                        const response = await fetch('/upload', {
                            method: 'POST',
                            body: formData
                        });

                        const result = await response.json();

                        if (result.success) {
                            document.getElementById('loader').style.display = 'none';
                            document.getElementById('status').innerHTML = \`
                                ✅ Dashboard generated!<br>
                                📊 \${result.records} records cleaned<br>
                                📋 \${result.columns.length} columns found<br>
                                <a href="\${result.dashboardUrl}" target="_blank" style="color:#667eea;font-weight:bold;">🔗 View Dashboard</a>
                                &nbsp;|&nbsp;
                                <a href="\${result.cleanedFile}" style="color:#667eea;font-weight:bold;">📥 Download Cleaned Data</a>
                            \`;
                            document.getElementById('status').innerHTML += \`
                                <div style="margin-top:15px;background:#f0f4f8;padding:15px;border-radius:10px;text-align:left;font-size:14px;">
                                    <strong>Preview (first 5 rows):</strong><br>
                                    <pre style="white-space:pre-wrap;font-size:12px;">\${JSON.stringify(result.preview, null, 2)}</pre>
                                </div>
                            \`;
                        } else {
                            document.getElementById('loader').style.display = 'none';
                            document.getElementById('status').innerHTML = \`❌ Error: \${result.error}\`;
                        }
                    } catch (error) {
                        document.getElementById('loader').style.display = 'none';
                        document.getElementById('status').innerHTML = \`❌ Error: \${error.message}\`;
                    }
                }
            </script>
        </body>
        </html>
    `);
});

// Start server
app.listen(PORT, () => {
    console.log(`🚀 Dashboard Generator running at http://localhost:${PORT}`);
    console.log(`📊 Upload any file to auto-generate interactive dashboards`);
});