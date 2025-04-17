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
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        const uploadDir = 'uploads';
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
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        questionBank = processExcelData(jsonData);
        console.log('Processed questionBank:', questionBank);

        fs.unlinkSync(req.file.path);

        res.json({
            message: 'File processed successfully',
            questionCount: questionBank.length
        });
    } catch (error) {
        console.error('Error processing file:', error);
        res.status(500).json({ error: 'Error processing file' });
    }
});

// Helper Function to Process Excel Data with Line Break Handling
function processExcelData(data) {
    return data.map((row, index) => {
        const btLevelRaw = String(row['B.T Level'] || '').trim();
        const btLevel = btLevelRaw.replace(/^L/i, '');
        
        let questionText = row.Question || '';
        if (typeof questionText === 'string') {
            questionText = questionText.replace(/\"<br>\"/g, '<br>');
            if (!questionText.includes('<br>')) {
                questionText = questionText.replace(/(\d+\.\s|[a-z]\)\s)/g, '$1<br>');
            }
        }

        return {
            id: index + 1,
            unit: parseInt(row.Unit) || 0,
            question: questionText,
            btLevel: btLevel || '0',
            subjectCode: row['Subject Code'] || '',
            subject: row.Subject || '',
            branch: row.Branch || '',
            regulation: row.Regulation || '',
            year: row.Year || '',
            semester: row.Sem || '',
            month: row.Month || '',
            imageUrl: row['Image Url'] ? getDirectImageURL(row['Image Url']) : ''
        };
    }).filter(q => q.unit >= 1 && q.unit <= 5 && q.btLevel !== '0');
}

// Function to generate questions with strict BTL enforcement and specific unit order
function generateQuestions(paperType) {
    if (!questionBank || questionBank.length < 6) {
        throw new Error('Insufficient questions in question bank. At least 6 questions are required.');
    }

    // Step 1: Assess available questions by unit and BTL
    const availableByUnitAndBTL = {};
    const availableBTLs = new Set();
    for (let unit = 1; unit <= 5; unit++) {
        availableByUnitAndBTL[unit] = {};
        questionBank.filter(q => q.unit === unit).forEach(q => {
            if (!availableByUnitAndBTL[unit][q.btLevel]) {
                availableByUnitAndBTL[unit][q.btLevel] = [];
            }
            availableByUnitAndBTL[unit][q.btLevel].push(q);
            availableBTLs.add(q.btLevel);
        });
    }
    console.log('Available questions by unit and BTL:', availableByUnitAndBTL);
    console.log('Unique BTLs:', [...availableBTLs]);

    // Step 2: Determine maximum BTL level
    const btLevels = questionBank.map(q => parseInt(q.btLevel) || 0).filter(btl => btl > 0);
    if (btLevels.length === 0) {
        throw new Error('No valid BTL levels found in question bank');
    }
    const maxBTL = Math.max(...btLevels);
    console.log('Max BTL:', maxBTL);

    // Step 3: Define BTL requirements (strict enforcement)
    let btlRequirements;
    if (maxBTL === 6) {
        btlRequirements = [
            { level: '1', count: 1 },
            { level: '2', count: 2 },
            { level: '3', count: 1 },
            { level: '4', count: 1 },
            { level: 'random', options: ['5', '6'], count: 1 }
        ];
    } else if (maxBTL === 5) {
        btlRequirements = [
            { level: '1', count: 1 },
            { level: '2', count: 2 },
            { level: '3', count: 1 },
            { level: '4', count: 1 },
            { level: 'random', options: ['5', '3'], count: 1 }
        ];
    } else if (maxBTL === 4) {
        btlRequirements = [
            { level: '1', count: 1 },
            { level: '2', count: 2 },
            { level: '3', count: 1 },
            { level: '4', count: 1 },
            { level: 'random', options: ['3', '4'], count: 1 }
        ];
    } else if (maxBTL === 3) {
        btlRequirements = [
            { level: '1', count: 1 },
            { level: '2', count: 2 },
            { level: '3', count: 2 },
            { level: 'random', options: ['2', '3'], count: 1 }
        ];
    } else if (maxBTL === 2 && availableBTLs.has('2') && availableBTLs.has('1')) {
        btlRequirements = [
            { level: '1', count: 1 },
            { level: '2', count: 5 }
        ];
    } else if (availableBTLs.size === 1) {
        btlRequirements = [{ level: [...availableBTLs][0], count: 6 }];
    } else {
        throw new Error(`Unsupported case: Max BTL = ${maxBTL} with BTLs (${[...availableBTLs]}). Only Case i (max BTL = 6), Case ii (max BTL = 5), Case iii (max BTL = 4), Case iv (max BTL = 3 with BTL 2 & 3), or Case v (single BTL) are supported.`);
    }
    console.log('BTL Requirements:', btlRequirements);

    // Step 4: Define unit requirements based on paper type
    let unitRequirements;
    switch (paperType) {
        case 'mid1':
            unitRequirements = [
                { unit: 1, minCount: 2, maxCount: 3 },
                { unit: 2, minCount: 2, maxCount: 3 },
                { unit: 3, minCount: 1, maxCount: 1 }
            ];
            break;
        case 'mid2':
            unitRequirements = [
                { unit: 3, minCount: 1, maxCount: 1 },
                { unit: 4, minCount: 2, maxCount: 3 },
                { unit: 5, minCount: 2, maxCount: 3 }
            ];
            break;
        case 'special':
            unitRequirements = [
                { unit: 1, minCount: 1, maxCount: 2 },
                { unit: 2, minCount: 1, maxCount: 2 },
                { unit: 3, minCount: 1, maxCount: 2 },
                { unit: 4, minCount: 1, maxCount: 2 },
                { unit: 5, minCount: 1, maxCount: 2 }
            ];
            break;
        default:
            throw new Error('Invalid paper type');
    }
    console.log('Unit Requirements:', unitRequirements);

    // Step 5: Select questions with BTL priority and specific unit order
    const selectQuestions = (btlReqs, unitReqs) => {
        let selectedQuestions = [];
        let unitCount = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        let btlCount = {};
        let remainingQuestions = [...questionBank];

        // Helper to pick a question by BTL, adjusting units
        const pickQuestion = (btl) => {
            const allowedUnits = unitReqs.map(r => r.unit).sort((a, b) => {
                const aDiff = unitCount[a] - unitReqs.find(r => r.unit === a).minCount;
                const bDiff = unitCount[b] - unitReqs.find(r => r.unit === b).minCount;
                return aDiff - bDiff; // Prefer units below minCount
            });
            const available = remainingQuestions.filter(q => 
                allowedUnits.includes(q.unit) &&
                q.btLevel === btl &&
                unitCount[q.unit] < unitReqs.find(r => r.unit === q.unit).maxCount
            );
            if (available.length === 0) return null;
            const idx = Math.floor(Math.random() * available.length);
            const q = available[idx];
            remainingQuestions = remainingQuestions.filter(r => r.id !== q.id);
            unitCount[q.unit]++;
            btlCount[q.btLevel] = (btlCount[q.btLevel] || 0) + 1;
            return q;
        };

        // Enforce BTL requirements exactly
        for (const req of btlReqs) {
            let count = req.count;
            if (req.level === 'random') {
                const randomBTLs = req.options.filter(btl => 
                    unitReqs.some(u => (availableByUnitAndBTL[u.unit]?.[btl]?.length || 0) > (btlCount[btl] || 0))
                );
                if (randomBTLs.length === 0) {
                    throw new Error(`No questions available for BTL [${req.options.join(', ')}] in required units`);
                }
                while (count > 0) {
                    const btl = randomBTLs[Math.floor(Math.random() * randomBTLs.length)];
                    const q = pickQuestion(btl);
                    if (q) {
                        selectedQuestions.push(q);
                        count--;
                    } else {
                        throw new Error(`Insufficient questions for BTL ${btl} in required units`);
                    }
                }
            } else {
                const availableCount = unitReqs.reduce((sum, u) => 
                    sum + (availableByUnitAndBTL[u.unit]?.[req.level]?.length || 0) - (btlCount[req.level] || 0), 0);
                if (availableCount < req.count) {
                    throw new Error(`Insufficient questions for BTL ${req.level} (need ${req.count}, found ${availableCount}) in required units`);
                }
                while (count > 0) {
                    const q = pickQuestion(req.level);
                    if (q) {
                        selectedQuestions.push(q);
                        count--;
                    } else {
                        throw new Error(`Failed to pick BTL ${req.level} despite availability`);
                    }
                }
            }
        }

        // Validate unit minimums
        for (const req of unitReqs) {
            if (unitCount[req.unit] < req.minCount) {
                throw new Error(`Unit ${req.unit} has ${unitCount[req.unit]} questions, needs at least ${req.minCount}`);
            }
        }

        // Sort questions to match desired order based on paperType
        const sortedQuestions = [];
        if (paperType === 'mid1') {
            // Mid1 order: Q1-2 (Unit 1), Q3-4 (Unit 2), Q5 (Unit 3), Q6 (Unit 1 or 2)
            const unit1Questions = selectedQuestions.filter(q => q.unit === 1);
            const unit2Questions = selectedQuestions.filter(q => q.unit === 2);
            const unit3Questions = selectedQuestions.filter(q => q.unit === 3);

            if (unit1Questions.length < 2 || unit2Questions.length < 2 || unit3Questions.length < 1) {
                throw new Error('Insufficient questions to meet Mid1 unit distribution');
            }

            sortedQuestions.push(unit1Questions.shift()); // Q1: Unit 1
            sortedQuestions.push(unit1Questions.shift()); // Q2: Unit 1
            sortedQuestions.push(unit2Questions.shift()); // Q3: Unit 2
            sortedQuestions.push(unit2Questions.shift()); // Q4: Unit 2
            sortedQuestions.push(unit3Questions.shift()); // Q5: Unit 3

            // Q6: Unit 1 or 2
            const remainingMid1 = [...unit1Questions, ...unit2Questions];
            if (remainingMid1.length > 0) {
                sortedQuestions.push(remainingMid1[Math.floor(Math.random() * remainingMid1.length)]);
            } else {
                throw new Error('No remaining questions for Q6 from Unit 1 or 2');
            }
        } else if (paperType === 'mid2') {
            // Mid2 order: Q1 (Unit 3), Q2-3 (Unit 4), Q4-5 (Unit 5), Q6 (Unit 4 or 5)
            const unit3Questions = selectedQuestions.filter(q => q.unit === 3);
            const unit4Questions = selectedQuestions.filter(q => q.unit === 4);
            const unit5Questions = selectedQuestions.filter(q => q.unit === 5);

            if (unit3Questions.length < 1 || unit4Questions.length < 2 || unit5Questions.length < 2) {
                throw new Error('Insufficient questions to meet Mid2 unit distribution');
            }

            sortedQuestions.push(unit3Questions.shift()); // Q1: Unit 3
            sortedQuestions.push(unit4Questions.shift()); // Q2: Unit 4
            sortedQuestions.push(unit4Questions.shift()); // Q3: Unit 4
            sortedQuestions.push(unit5Questions.shift()); // Q4: Unit 5
            sortedQuestions.push(unit5Questions.shift()); // Q5: Unit 5

            // Q6: Unit 4 or 5
            const remainingMid2 = [...unit4Questions, ...unit5Questions];
            if (remainingMid2.length > 0) {
                sortedQuestions.push(remainingMid2[Math.floor(Math.random() * remainingMid2.length)]);
            } else {
                throw new Error('No remaining questions for Q6 from Unit 4 or 5');
            }
        } else if (paperType === 'special') {
            // Special order: One from each unit if possible, then sort as per Mid1 for simplicity
            const unit1Questions = selectedQuestions.filter(q => q.unit === 1);
            const unit2Questions = selectedQuestions.filter(q => q.unit === 2);
            const unit3Questions = selectedQuestions.filter(q => q.unit === 3);
            const unit4Questions = selectedQuestions.filter(q => q.unit === 4);
            const unit5Questions = selectedQuestions.filter(q => q.unit === 5);

            if (unit1Questions.length < 2 || unit2Questions.length < 2 || unit3Questions.length < 1) {
                throw new Error('Insufficient questions to meet Special unit distribution');
            }

            sortedQuestions.push(unit1Questions.shift()); // Q1: Unit 1
            sortedQuestions.push(unit1Questions.shift()); // Q2: Unit 1
            sortedQuestions.push(unit2Questions.shift()); // Q3: Unit 2
            sortedQuestions.push(unit2Questions.shift()); // Q4: Unit 2
            sortedQuestions.push(unit3Questions.shift()); // Q5: Unit 3

            const remainingSpecial = [...unit1Questions, ...unit2Questions, ...unit4Questions, ...unit5Questions];
            if (remainingSpecial.length > 0) {
                sortedQuestions.push(remainingSpecial[Math.floor(Math.random() * remainingSpecial.length)]);
            } else {
                throw new Error('No remaining questions for Q6 in Special');
            }
        }

        console.log('Selected Questions:', sortedQuestions.map(q => `Unit ${q.unit}, BTL ${q.btLevel}`));
        console.log('Unit Count:', unitCount);
        console.log('BTL Count:', btlCount);
        return sortedQuestions;
    };

    const selected = selectQuestions(btlRequirements, unitRequirements);

    if (selected.length !== 6) {
        throw new Error('Failed to select exactly 6 questions with required BTL and unit constraints');
    }

    return selected;
}

// API Endpoint to Generate Questions
app.post('/api/generate', (req, res) => {
    try {
        if (!questionBank) {
            return res.status(400).json({ error: 'No questions available. Please upload an Excel file first.' });
        }

        const { paperType } = req.body;
        const selectedQuestions = generateQuestions(paperType);
        console.log('Generated Questions (Limited Info):');
        selectedQuestions.forEach((q, index) => {
            console.log(`Question ${index + 1}:`);
            console.log(`  Question: ${q.question}`);
            console.log(`  Unit: ${q.unit}`);
            console.log(`  Subject: ${q.subject}`);
            console.log(`  Year: ${q.year}`);
            console.log('------------------------');
        });

        res.json({
            questions: selectedQuestions.map(q => ({
                question: q.question,
                imageUrl: q.imageUrl,
                btLevel: q.btLevel,
                unit: q.unit
            })),
            paperDetails: {
                subject: selectedQuestions[0].subject,
                subjectCode: selectedQuestions[0].subjectCode,
                branch: selectedQuestions[0].branch,
                regulation: selectedQuestions[0].regulation,
                year: selectedQuestions[0].year,
                semester: selectedQuestions[0].semester
            }
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