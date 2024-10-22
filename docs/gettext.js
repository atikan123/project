const express = require('express');
const multer = require('multer');
const fetch = require('node-fetch');
const cheerio = require('cheerio');
const { Document, Packer, Paragraph, HeadingLevel, TextRun } = require("docx");
const pdfParse = require('pdf-parse');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const { v4: uuidv4 } = require('uuid');
const Tesseract = require('tesseract.js');

const app = express();
const { AlignmentType } = require("docx");

const upload = multer().fields([
    { name: 'pdf', maxCount: 1 },
    { name: 'images', maxCount: 10 }
]);

const extractTextFromPDF = async (buffer) => {
    try {
        const data = await pdfParse(buffer);
        return data.text;
    } catch (err) {
        console.error('Error extracting PDF text:', err);
        return '';
    }
};

// ฟังก์ชันดึงข้อความจากภาพทั้งไทยและอังกฤษ
const extractTextFromImages = async (imageFiles) => {
    const extractedTexts = [];
    
    for (const imageFile of imageFiles) {
        try {
            const { data: { text } } = await Tesseract.recognize(imageFile.buffer, 'tha+eng');
            extractedTexts.push(text);
        } catch (err) {
            console.error(`Error extracting text from image:`, err);
        }
    }

    return extractedTexts.join('\n');
};

// ฟังก์ชันดึงภาพจาก URL
const extractImagesFromURL = async (url) => {
    const imageUrls = [];
    try {
        const response = await fetch(url);
        const html = await response.text();
        const $ = cheerio.load(html);

        // หาภาพจากแท็ก <img>
        $('img').each((i, img) => {
            const src = $(img).attr('src');
            if (src) {
                // ตรวจสอบว่าเป็น URL ที่สมบูรณ์หรือไม่
                const imageUrl = src.startsWith('http') ? src : new URL(src, url).href;
                imageUrls.push(imageUrl);
            }
        });
    } catch (err) {
        console.error('Error fetching website:', err);
    }
    return imageUrls;
};

app.post('/', upload, async (req, res) => {
    const url = req.body.url;
    const pdfFile = req.files['pdf'] ? req.files['pdf'][0] : null;
    const imageFiles = req.files['images'] || [];
    const children = [];
    const extractedImageFiles = []; // เก็บชื่อไฟล์ภาพ

    // การดึงข้อมูลจาก URL
    if (url) {
        try {
            const imageUrls = await extractImagesFromURL(url);
            for (const imageUrl of imageUrls) {
                const response = await fetch(imageUrl);
                const buffer = await response.buffer();
                const imagePath = path.join(__dirname, `${uuidv4()}.png`);
                fs.writeFileSync(imagePath, buffer);
                extractedImageFiles.push(imagePath); // เก็บชื่อไฟล์ที่ถูกบันทึก
            }

            // ดึง headings และ paragraphs
            const response = await fetch(url);
            const html = await response.text();
            const $ = cheerio.load(html);

            $('body *').each((i, elem) => {
                const tagName = $(elem).get(0).tagName;
                if (tagName.startsWith('h') && tagName.length === 2) {
                    const headingText = $(elem).text();
                    const headingLevel = parseInt(tagName.charAt(1), 10);
                    children.push(new Paragraph({ children: [new TextRun({ text: headingText, bold: true })], heading: headingLevel }));
                } else if (tagName === 'p') {
                    const paragraphText = $(elem).text();
                    children.push(new Paragraph(paragraphText));
                }
            });

        } catch (err) {
            console.log('Error fetching website:', err);
        }
    }

    // การดึงข้อมูลจาก PDF
    if (pdfFile) {
        try {
            const pdfText = await extractTextFromPDF(pdfFile.buffer);
            if (pdfText) {
                const pdfLines = pdfText.split('\n').map(line => new Paragraph(line));
                children.push(new Paragraph({ text: 'Extracted Text from PDF', heading: HeadingLevel.HEADING_1 }));
                children.push(...pdfLines); // เพิ่มบรรทัดจาก PDF โดยตรง
            }
        } catch (err) {
            console.error(`Error extracting text from PDF: ${err}`);
        }
    }
    
    // การดึงข้อความจากภาพ
    if (imageFiles.length > 0) {
        try {
            children.push(new Paragraph({ text: 'Extracted Text from Images', heading: HeadingLevel.HEADING_1 }));
            const imageTexts = await extractTextFromImages(imageFiles);
            if (imageTexts) {
                children.push(new Paragraph(imageTexts));
            }
        } catch (err) {
            console.error(`Error extracting text from images: ${err}`);
        }
    }

    // สร้างเอกสาร Word จากข้อมูลที่ดึงมา
    const doc = new Document({ sections: [{ properties: {}, children: children }] });
    const wordBuffer = await Packer.toBuffer(doc);
    const zipFilePath = path.join(__dirname, 'output.zip');
    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => {
        console.log(`ZIP file created: ${archive.pointer()} total bytes`);
        res.set({
            'Content-Type': 'application/zip',
            'Content-Disposition': 'attachment; filename=output.zip',
        });
        res.sendFile(zipFilePath);
        // ลบไฟล์ภาพที่เก็บไว้
        extractedImageFiles.forEach(filePath => fs.unlinkSync(filePath));
    });

    archive.on('error', (err) => {
        throw err;
    });

    archive.pipe(output);
    archive.append(wordBuffer, { name: 'document.docx' });

    // เพิ่มไฟล์ภาพลงใน ZIP
    extractedImageFiles.forEach(filePath => {
        archive.file(filePath, { name: path.basename(filePath) });
    });

    archive.finalize();
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
