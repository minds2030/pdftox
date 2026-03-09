const axios = require('axios');
const ExcelJS = require('exceljs');

async function start() {
    const fileId = process.env.FILE_ID;
    const gasUrl = process.env.GAS_URL;
    const groqKey = process.env.GROQ_API_KEY;

    try {
        console.log("Fetching spatial data from Google OCR...");
        const response = await axios.get(`${gasUrl}?fileId=${fileId}`);
        const elements = response.data;

        if (!elements || (typeof elements === 'string' && elements.includes("Error"))) {
            throw new Error("Failed to get data: " + JSON.stringify(elements));
        }

        // تقليل حجم الـ Chunk لضمان عدم حدوث Error في الـ JSON
        const chunkSize = 50; 
        let finalRows = [];

        console.log(`Processing ${elements.length} elements in chunks...`);

        for (let i = 0; i < elements.length; i += chunkSize) {
            const chunk = elements.slice(i, i + chunkSize);
            const isFirst = (i === 0);
            
            console.log(`Processing Chunk ${Math.floor(i/chunkSize) + 1}...`);

            const res = await axios.post('https://api.groq.com/openai/v1/chat/completions', {
                model: "llama-3.3-70b-versatile",
                messages: [{
                    role: "system", 
                    content: "Strict JSON Mode: Transform the spatial data into a 2D JSON array. NO CONVERSATION. NO MARKDOWN. ONLY [[row1],[row2]]."
                }, {
                    role: "user", 
                    content: `Rules:
                    1. Use startIndex to identify column gaps and row breaks.
                    2. If a column is missing, insert "".
                    3. Do not change or fix any characters.
                    Data: ${JSON.stringify(chunk)}`
                }],
                temperature: 0,
                response_format: { "type": "json_object" } // تفعيل وضع الـ JSON الصارم لو الموديل بيدعمه
            }, {
                headers: { 'Authorization': `Bearer ${groqKey}`, 'Content-Type': 'application/json' }
            });

            let rawContent = res.data.choices[0].message.content;
            
            // تنظيف صارم لأي حرف بره المصفوفة
            const match = rawContent.match(/\[\s*\[[\s\S]*\]\s*\]/);
            if (match) {
                const data = JSON.parse(match[0]);
                finalRows.push(...data);
            } else {
                console.error("Critical: Model returned non-JSON content.");
            }
            
            await new Promise(r => setTimeout(r, 1000)); 
        }

        if (finalRows.length > 0) {
            console.log("Saving results to Result.xlsx...");
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Data');
            
            // إضافة البيانات مع التأكد من أن كل صف مصفوفة
            finalRows.forEach(row => {
                if (Array.isArray(row)) sheet.addRow(row);
            });

            await workbook.xlsx.writeFile('Result.xlsx');
            console.log("✅ DONE! Task Completed Successfully.");
        }

    } catch (e) {
        console.error("❌ ERROR:", e.message);
        process.exit(1);
    }
}

start();
