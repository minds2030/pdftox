const axios = require('axios');
const ExcelJS = require('exceljs');

async function start() {
    const fileId = process.env.FILE_ID;
    const gasUrl = process.env.GAS_URL;
    const groqKey = process.env.GROQ_API_KEY;

    // 1. جلب النص من جوجل (OCR)
    console.log("Fetching text from Google OCR...");
    const response = await axios.get(`${gasUrl}?fileId=${fileId}`);
    const fullText = response.data;

    // 2. تقسيم النص لقطع (Chunks)
    const lines = fullText.split('\n').filter(l => l.trim());
    const chunkSize = 30; // 30 سطر لكل طلب عشان الـ TPM
    let finalRows = [];

    console.log(`Processing ${lines.length} lines in chunks...`);

    for (let i = 0; i < lines.length; i += chunkSize) {
        const chunk = lines.slice(i, i + chunkSize).join('\n');
        
        try {
            const res = await axios.post('https://api.groq.com/openai/v1/chat/completions', {
                model: "llama-3.3-70b-versatile",
                messages: [{role: "user", content: `Convert this OCR text to a 2D JSON array. No headers except first chunk. JSON ONLY:\n${chunk}`}]
            }, {
                headers: { 'Authorization': `Bearer ${groqKey}` }
            });

            const data = JSON.parse(res.data.choices[0].message.content.replace(/```json|```/g, ""));
            finalRows.push(...data);
            console.log(`Processed chunk ${Math.floor(i/chunkSize) + 1}`);
            
            // انتظر ثانية لتجنب الـ Rate Limit
            await new Promise(r => setTimeout(r, 1000)); 
        } catch (e) {
            console.error("Chunk failed, skipping...");
        }
    }

    // 3. إنشاء ملف إكسيل احترافي
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Data');
    sheet.addRows(finalRows);
    
    await workbook.xlsx.writeFile('Result.xlsx');
    console.log("Excel file created: Result.xlsx");
}

start();
