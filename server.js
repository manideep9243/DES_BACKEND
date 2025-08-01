const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const axios = require('axios');

const app = express();
const port = 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use('/uploads', express.static(path.join(__dirname, 'Uploads')));

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        const uploadDir = 'Uploads';
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir);
        }
        cb(null, uploadDir);
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            file.mimetype === 'application/vnd.ms-excel') {
            cb(null, true);
        } else {
            cb(null, false);
            return cb(new Error('Only Excel files are allowed!'));
        }
    }
});

// Store questions in memory
let questionBank = null;

// Function to convert Google Drive sharing URL to direct image URL
function getDirectImageURL(url) {
    const driveRegex = /https:\/\/drive\.google\.com\/file\/d\/([^/]+)\/view/;
    const match = url.match(driveRegex);
    return match ? `https://drive.google.com/uc?export=view&id=${match[1]}` : url;
}

// Proxy endpoint to fetch image and return base64 data
app.get('/api/image-proxy-base64', async (req, res) => {
    const { url } = req.query;
    if (!url) {
        console.error('No URL provided to /api/image-proxy-base64');
        return res.status(400).json({ error: 'No URL provided' });
    }

    const directUrl = getDirectImageURL(url);
    console.log(`Fetching image from: ${directUrl}`);

    try {
        const response = await axios.get(directUrl, {
            responseType: 'arraybuffer',
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Referer': 'https://drive.google.com'
            }
        });

        const contentType = response.headers['content-type'];
        if (!contentType.startsWith('image/')) {
            console.error(`Invalid content type from ${directUrl}: ${contentType}`);
            return res.status(400).json({ error: 'URL does not point to an image', contentType });
        }

        const base64Data = Buffer.from(response.data).toString('base64');
        const dataUrl = `data:${contentType};base64,${base64Data}`;
        console.log(`Successfully fetched image from ${directUrl}, data URL length: ${dataUrl.length}, starts with: ${dataUrl.substring(0, 50)}...`);
        
        res.json({ dataUrl });
    } catch (error) {
        console.error(`Image proxy error for ${directUrl}:`, error.message, error.response?.status, error.response?.data?.toString());
        const placeholder = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAvElEQVR4nO3YQQqDMBAF0L/KnW+/Q6+xu1oSLeI4DAgAAAAAAAAA7rZpm7Zt2/9eNpvNZrPZdrsdANxut9vt9nq9PgAwGo1Go9FoNBr9MabX6/U2m01mM5vNZnO5XC6X+wDAXC6Xy+VyuVwul8sFAKPRaDQajUaj0Wg0Go1Goz8A8Hg8Ho/H4/F4PB6Px+MBgMFoNBqNRqPRaDQajUaj0Wg0Go1Goz8AAAAAAAAA7rYBAK3eVREcAAAAAElFTkSuQmCC';
        res.json({ dataUrl: placeholder });
    }
});

// API Endpoint to Upload and Process Excel File
app.post('/api/upload', upload.single('excelFile'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        
        console.log('Raw Excel Data (first 5 rows):', jsonData.slice(0, 5));
        
        questionBank = processExcelData(jsonData);
        console.log('Processed questionBank (first 5 entries):', questionBank.slice(0, 5));
        console.log('Total questions:', questionBank.length);

        // Validate question bank
        if (questionBank.length < 17) {
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: `Insufficient questions: got ${questionBank.length}, need at least 17 (5 for Part A, 12 for Part B)` });
        }

        // Check available questions by unit and BTL
        const partAQuestions = questionBank.filter(q => q.btLevel === '1');
        const partBQuestions = questionBank.filter(q => q.btLevel !== '1');
        console.log(`Part A questions (BTL L1): ${partAQuestions.length}`);
        console.log(`Part B questions (BTL L2-L6): ${partBQuestions.length}`);

        const unitCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        const btlCounts = { '1': 0, '2': 0, '3': 0, '4': 0, '5': 0, '6': 0 };
        questionBank.forEach(q => {
            unitCounts[q.unit]++;
            btlCounts[q.btLevel]++;
        });
        console.log('Questions per unit:', unitCounts);
        console.log('Questions per BTL:', btlCounts);

        // Validate required fields
        const requiredFields = ['subjectCode', 'subject', 'branch', 'regulation', 'year', 'semester'];
        const sampleQuestion = questionBank[0];
        for (const field of requiredFields) {
            if (!sampleQuestion[field] || sampleQuestion[field] === '') {
                fs.unlinkSync(req.file.path);
                return res.status(400).json({ error: `Missing or empty field '${field}' in Excel data` });
            }
        }

        // Validate BTL levels and units
        const validBTLevels = ['1', '2', '3', '4', '5', '6'];
        const validUnits = [1, 2, 3, 4, 5];
        const invalidQuestions = questionBank.filter(
            q => !validBTLevels.includes(q.btLevel) || !validUnits.includes(q.unit)
        );
        if (invalidQuestions.length > 0) {
            console.log('Invalid questions:', invalidQuestions);
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: `Invalid BTL levels or units in ${invalidQuestions.length} questions` });
        }

        fs.unlinkSync(req.file.path);
        res.json({
            message: 'File processed successfully',
            questionCount: questionBank.length
        });
    } catch (error) {
        console.error('Error processing file:', error);
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        res.status(500).json({ error: 'Error processing file: ' + error.message });
    }
});

// Helper Function to Convert Roman Numerals to Integers
function romanToInt(roman) {
    const romanMap = { 'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5 };
    return romanMap[String(roman).toUpperCase()] || 0;
}

// Helper Function to Convert Excel Date to Readable Format
function excelDateToString(excelDate) {
    if (!excelDate || isNaN(excelDate)) return '';
    const date = new Date((excelDate - 25569) * 86400 * 1000); // Excel epoch: Jan 1, 1900
    return date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
}

// Helper Function to Process Excel Data with Line Break Handling
function processExcelData(data) {
    return data.map((row, index) => {
        const btLevelRaw = String(row['B.T Level'] || '').trim();
        const btLevel = btLevelRaw.replace(/^L/i, '');
        
        let questionText = String(row.Question || '');
        if (questionText) {
            questionText = questionText.replace(/\"<br>\"/g, '<br>');
            if (!questionText.includes('<br>')) {
                questionText = questionText.replace(/(\d+\.\s|[a-z]\)\s)/g, '$1<br>');
            }
        }

        const unit = romanToInt(row.Unit);
        const month = excelDateToString(row.Month);
        
        return {
            id: index + 1,
            unit: unit,
            question: questionText,
            btLevel: btLevel || '0',
            subjectCode: String(row['Subject Code'] || ''),
            subject: String(row.Subject || ''),
            branch: String(row.Branch || ''),
            regulation: String(row.Regulation || ''),
            year: String(row.Year || ''),
            semester: String(row.Sem || ''),
            month: month,
            imageUrl: row['Image Url'] ? getDirectImageURL(String(row['Image Url'])) : '',
            sno: String(row['S.NO'] || '')
        };
    }).filter(q => q.unit >= 1 && q.unit <= 5 && q.btLevel !== '0');
}

// Function to generate questions for Part A and Part B
function generateQuestions(paperType) {
    if (!questionBank || questionBank.length < 17) {
        throw new Error(`Insufficient questions in question bank: got ${questionBank.length}, need at least 17 (5 for Part A, 12 for Part B)`);
    }

    // Step 1: Filter questions by BTL and Unit for Part A and Part B
    const partAQuestions = questionBank.filter(q => q.btLevel === '1' && q.unit >= 1 && q.unit <= 5);
    const partBQuestions = questionBank.filter(q => q.btLevel !== '1' && q.unit >= 1 && q.unit <= 5);

    // Step 2: Validate available questions for the specific paper type
    let unitRequirements;
    let partAUnitRequirements;
    if (paperType === 'mid1') {
        partAUnitRequirements = [
            { unit: 1, count: 2 },
            { unit: 2, count: 2 },
            { unit: 3, count: 1 }
        ];
        unitRequirements = [
            { unit: 1, minCount: 5, maxCount: 5 },
            { unit: 2, minCount: 5, maxCount: 5 },
            { unit: 3, minCount: 2, maxCount: 2 }
        ];
    } else if (paperType === 'mid2') {
        partAUnitRequirements = [
            { unit: 3, count: 1 },
            { unit: 4, count: 2 },
            { unit: 5, count: 2 }
        ];
        unitRequirements = [
            { unit: 3, minCount: 2, maxCount: 2 },
            { unit: 4, minCount: 5, maxCount: 5 },
            { unit: 5, minCount: 5, maxCount: 5 }
        ];
    } else {
        throw new Error('Invalid paper type');
    }

    // Check Part A requirements
    for (const req of partAUnitRequirements) {
        const count = partAQuestions.filter(q => q.unit === req.unit).length;
        if (count < req.count) {
            throw new Error(`Insufficient BTL L1 questions for Unit ${req.unit}: got ${count}, need ${req.count}`);
        }
    }

    // Check Part B requirements
    for (const req of unitRequirements) {
        const count = partBQuestions.filter(q => q.unit === req.unit).length;
        if (count < req.minCount) {
            throw new Error(`Insufficient BTL L2-L6 questions for Unit ${req.unit}: got ${count}, need ${req.minCount}`);
        }
    }

    // Step 3: Assess available questions by unit and BTL for Part B
    const availableByUnitAndBTL = {};
    const availableBTLs = new Set();
    for (let unit = 1; unit <= 5; unit++) {
        availableByUnitAndBTL[unit] = {};
        partBQuestions.filter(q => q.unit === unit).forEach(q => {
            if (!availableByUnitAndBTL[unit][q.btLevel]) {
                availableByUnitAndBTL[unit][q.btLevel] = [];
            }
            availableByUnitAndBTL[unit][q.btLevel].push(q);
            availableBTLs.add(q.btLevel);
        });
    }
    console.log('Available questions for Part B by unit and BTL:', availableByUnitAndBTL);
    console.log('Unique BTLs for Part B:', [...availableBTLs]);

    // Step 4: Determine maximum BTL level for Part B
    const btLevels = partBQuestions.map(q => parseInt(q.btLevel) || 0).filter(btl => btl > 0);
    if (btLevels.length === 0) {
        throw new Error('No valid BTL levels found in Part B question bank');
    }
    const maxBTL = Math.max(...btLevels);
    console.log('Max BTL for Part B:', maxBTL);

    // Step 5: Define BTL requirements for Part B
    let btlRequirements;
    if (maxBTL === 6) {
        btlRequirements = [
            { level: '2', count: 4 },
            { level: '3', count: 4 },
            { level: '4', count: 2 },
            { level: 'random', options: ['5', '6'], count: 2 }
        ];
    } else if (maxBTL === 5) {
        btlRequirements = [
            { level: '2', count: 4 },
            { level: '3', count: 4 },
            { level: '4', count: 2 },
            { level: 'random', options: ['5', '3'], count: 2 }
        ];
    } else if (maxBTL === 4) {
        btlRequirements = [
            { level: '2', count: 4 },
            { level: '3', count: 4 },
            { level: '4', count: 2 },
            { level: 'random', options: ['3', '4'], count: 2 }
        ];
    } else if (maxBTL === 3) {
        btlRequirements = [
            { level: '2', count: 5 },
            { level: '3', count: 5 },
            { level: 'random', options: ['2', '3'], count: 2 }
        ];
    } else if (maxBTL === 2 && availableBTLs.has('2')) {
        btlRequirements = [
            { level: '2', count: 12 }
        ];
    } else if (availableBTLs.size === 1) {
        btlRequirements = [{ level: [...availableBTLs][0], count: 12 }];
    } else {
        throw new Error(`Unsupported case: Max BTL = ${maxBTL} with BTLs (${[...availableBTLs]}).`);
    }
    console.log('BTL Requirements for Part B:', btlRequirements);

    // Step 6: Define unit requirements and question labels based on paper type
    let questionLabels;
    let partALabels;
    if (paperType === 'mid1') {
        questionLabels = [
            { label: '2a', unit: 1 },
            { label: '2b', unit: 1 },
            { label: '3a', unit: 1 },
            { label: '3b', unit: 1 },
            { label: '6b', unit: 1 },
            { label: '4a', unit: 2 },
            { label: '4b', unit: 2 },
            { label: '5a', unit: 2 },
            { label: '5b', unit: 2 },
            { label: '7b', unit: 2 },
            { label: '6a', unit: 3 },
            { label: '7a', unit: 3 }
        ];
        partALabels = [
            { label: '1', unit: 1 },
            { label: '2', unit: 1 },
            { label: '3', unit: 2 },
            { label: '4', unit: 2 },
            { label: '5', unit: 3 }
        ];
    } else if (paperType === 'mid2') {
        questionLabels = [
            { label: '2a', unit: 3 },
            { label: '2b', unit: 3 },
            { label: '3a', unit: 4 },
            { label: '3b', unit: 4 },
            { label: '4a', unit: 4 },
            { label: '4b', unit: 4 },
            { label: '5a', unit: 4 },
            { label: '5b', unit: 5 },
            { label: '6a', unit: 5 },
            { label: '6b', unit: 5 },
            { label: '7a', unit: 5 },
            { label: '7b', unit: 5 }
        ];
        partALabels = [
            { label: '1', unit: 3 },
            { label: '2', unit: 4 },
            { label: '3', unit: 4 },
            { label: '4', unit: 5 },
            { label: '5', unit: 5 }
        ];
    } else {
        throw new Error('Invalid paper type');
    }
    console.log('Unit Requirements for Part B:', unitRequirements);
    console.log('Question Labels for Part B:', questionLabels);
    console.log('Unit Requirements for Part A:', partAUnitRequirements);
    console.log('Question Labels for Part A:', partALabels);

    // Step 7: Select questions for Part A
    const selectPartAQuestions = (unitReqs, labels) => {
        let selectedQuestions = [];
        let unitCount = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        let remainingQuestions = [...partAQuestions];

        const pickQuestionFromUnit = (unit) => {
            let unitQuestions = remainingQuestions.filter(q => q.unit === unit && q.btLevel === '1');
            if (unitQuestions.length === 0) {
                throw new Error(`No BTL L1 questions available for Unit ${unit}`);
            }
            const idx = Math.floor(Math.random() * unitQuestions.length);
            const q = unitQuestions[idx];
            remainingQuestions = remainingQuestions.filter(r => r.id !== q.id);
            unitCount[q.unit]++;
            return q;
        };

        for (const label of labels) {
            const q = pickQuestionFromUnit(label.unit);
            selectedQuestions.push({ ...q, label: label.label, part: 'A' });
        }

        // Validate unit requirements
        for (const req of unitReqs) {
            if (unitCount[req.unit] !== req.count) {
                throw new Error(`Unit ${req.unit} has ${unitCount[req.unit]} questions, needs exactly ${req.count}`);
            }
        }

        console.log('Selected Part A Questions:', selectedQuestions.map(q => `Label ${q.label}, Unit ${q.unit}, BTL ${q.btLevel}`));
        console.log('Part A Unit Count:', unitCount);
        return selectedQuestions;
    };

    // Step 8: Select questions for Part B
    const selectPartBQuestions = (btlReqs, unitReqs, labels) => {
        let selectedQuestions = [];
        let unitCount = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        let btlCount = {};
        let remainingQuestions = [...partBQuestions];

        const pickQuestionFromUnit = (btl, unit) => {
            let unitQuestions = remainingQuestions.filter(q => q.unit === unit);
            let btlMatches = unitQuestions.filter(q => q.btLevel === btl);
            if (btlMatches.length === 0 || unitCount[unit] >= unitReqs.find(r => r.unit === unit).maxCount) {
                btlMatches = unitQuestions;
            }
            if (btlMatches.length === 0) {
                throw new Error(`No questions available for Unit ${unit}`);
            }
            const idx = Math.floor(Math.random() * btlMatches.length);
            const q = btlMatches[idx];
            remainingQuestions = remainingQuestions.filter(r => r.id !== q.id);
            unitCount[q.unit]++;
            btlCount[q.btLevel] = (btlCount[q.btLevel] || 0) + 1;
            return q;
        };

        for (const label of labels) {
            let btl = null;
            for (const req of btlReqs) {
                if (req.count > 0) {
                    btl = req.level === 'random' ? req.options[Math.floor(Math.random() * req.options.length)] : req.level;
                    req.count--;
                    break;
                }
            }
            if (!btl) {
                btl = [...availableBTLs][Math.floor(Math.random() * availableBTLs.size)];
            }
            const q = pickQuestionFromUnit(btl, label.unit);
            selectedQuestions.push({ ...q, label: label.label, part: 'B' });
        }

        // Validate unit requirements
        for (const req of unitReqs) {
            if (unitCount[req.unit] < req.minCount) {
                throw new Error(`Unit ${req.unit} has ${unitCount[req.unit]} questions, needs at least ${req.minCount}`);
            }
        }

        // Sort by unit and label
        selectedQuestions.sort((a, b) => {
            if (a.unit !== b.unit) return a.unit - b.unit;
            const labelOrder = ['2a', '2b', '3a', '3b', '4a', '4b', '5a', '5b', '6a', '6b', '7a', '7b'];
            return labelOrder.indexOf(a.label) - labelOrder.indexOf(b.label);
        });

        console.log('Selected Part B Questions:', selectedQuestions.map(q => `Label ${q.label}, Unit ${q.unit}, BTL ${q.btLevel}`));
        console.log('Part B Unit Count:', unitCount);
        console.log('Part B BTL Count:', btlCount);
        return selectedQuestions;
    };

    const partASelected = selectPartAQuestions(partAUnitRequirements, partALabels);
    const partBSelected = selectPartBQuestions(btlRequirements, unitRequirements, questionLabels);

    if (partASelected.length !== 5) {
        throw new Error('Failed to select exactly 5 questions for Part A');
    }
    if (partBSelected.length !== 12) {
        throw new Error('Failed to select exactly 12 questions for Part B');
    }

    return { partA: partASelected, partB: partBSelected };
}

// API Endpoint to Generate Questions
app.post('/api/generate', (req, res) => {
    try {
        if (!questionBank) {
            return res.status(400).json({ error: 'No questions available. Please upload an Excel file first.' });
        }

        const { paperType } = req.body;
        if (!['mid1', 'mid2'].includes(paperType)) {
            return res.status(400).json({ error: 'Invalid paper type' });
        }

        const { partA, partB } = generateQuestions(paperType);
        console.log('Generated Questions:');
        console.log('Part A:');
        partA.forEach((q, index) => {
            console.log(`Question ${q.label}:`);
            console.log(`  Question: ${q.question}`);
            console.log(`  Unit: ${q.unit}`);
            console.log(`  BTL: ${q.btLevel}`);
            console.log(`  Subject: ${q.subject}`);
            console.log(`  Subject Code: ${q.subjectCode}`);
            console.log(`  Year: ${q.year}`);
            console.log('------------------------');
        });
        console.log('Part B:');
        partB.forEach((q, index) => {
            console.log(`Question ${q.label}:`);
            console.log(`  Question: ${q.question}`);
            console.log(`  Unit: ${q.unit}`);
            console.log(`  BTL: ${q.btLevel}`);
            console.log(`  Subject: ${q.subject}`);
            console.log(`  Subject Code: ${q.subjectCode}`);
            console.log(`  Year: ${q.year}`);
            console.log('------------------------');
        });

        // Extract paper details from the first question
        const paperDetails = {
            subjectCode: partA[0]?.subjectCode || partB[0]?.subjectCode || '',
            subject: partA[0]?.subject || partB[0]?.subject || '',
            branch: partA[0]?.branch || partB[0]?.branch || '',
            regulation: partA[0]?.regulation || partB[0]?.regulation || '',
            year: partA[0]?.year || partB[0]?.year || '',
            semester: partA[0]?.semester || partB[0]?.semester || '',
            month: partA[0]?.month || partB[0]?.month || ''
        };

        // Validate paper details
        const requiredFields = ['subjectCode', 'subject', 'branch', 'regulation', 'year', 'semester'];
        for (const field of requiredFields) {
            if (!paperDetails[field] || paperDetails[field] === '') {
                return res.status(400).json({ error: `Missing or empty field '${field}' in paper details` });
            }
        }

        res.json({
            partA: partA.map(q => ({
                question: q.question,
                imageUrl: q.imageUrl,
                btLevel: q.btLevel,
                unit: q.unit,
                label: q.label
            })),
            partB: partB.map(q => ({
                question: q.question,
                imageUrl: q.imageUrl,
                btLevel: q.btLevel,
                unit: q.unit,
                label: q.label
            })),
            paperDetails
        });
    } catch (error) {
        console.error('Error generating questions:', error.message);
        res.status(500).json({ error: 'Error generating questions: ' + error.message });
    }
});

// Start the Server
app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});