const axios = require('axios');
const ExcelJS = require('exceljs');

async function start() {
    const fileId = process.env.FILE_ID;
    const gasUrl = process.env.GAS_URL;
    const groqKey = process.env.GROQ_API_KEY;

    try {
        // 1. جلب بيانات الإحداثيات من جوجل (التي تم تحديثها لترسل JSON)
        console.log("Fetching spatial data (Indices) from Google OCR...");
        const response = await axios.get(`${gasUrl}?fileId=${fileId}`);
        const elements = response.data; // مصفوفة الـ JSON التي تحتوي على الكلمات والمواقع

        if (!elements || (typeof elements === 'string' && elements.includes("Error"))) {
            throw new Error("Failed to get data from Google OCR: " + JSON.stringify(elements));
        }

        // 2. تقسيم العناصر لقطع (Chunks)
        // بما أن العناصر JSON، سنقسمها بعدد الكلمات (مثلاً كل 100 عنصر)
        const chunkSize = 100; 
        let finalRows = [];

        console.log(`Processing ${elements.length} spatial elements in chunks...`);

        for (let i = 0; i < elements.length; i += chunkSize) {
            const chunk = elements.slice(i, i + chunkSize);
            const isFirstChunk = (i === 0);
            
            console.log(`Sending chunk ${Math.floor(i/chunkSize) + 1} with spatial awareness...`);

            // تعريف الـ Prompt داخل الطلب مباشرة لضمان الدقة
            const res = await axios.post('https://api.groq.com/openai/v1/chat/completions', {
                model: "llama-3.3-70b-versatile",
                messages: [{
                    role: "system", 
                    content: "You are a spatial data engineer. You will receive JSON elements with text and startIndex. Use the startIndex to detect row breaks and column gaps. If a column is missing in the index sequence, keep it empty \"\"."
                }, {
                    role: "user", 
                    content: `Convert these spatial elements into a 2D JSON array. ${isFirstChunk ? "Include headers in the first row." : "No headers."} Data:\n${JSON.stringify(chunk)}`
                }],
                temperature: 0
            }, {
                headers: { 'Authorization': `Bearer ${groqKey}`, 'Content-Type': 'application/json' }
            });

            let rawContent = res.data.choices[0].message.content;
            const jsonMatch = rawContent.match(/\[\s*\[[\s\S]*\]\s*\]/);
            
            if (jsonMatch) {
                const data = JSON.parse(jsonMatch[0]);
                finalRows.push(...data);
                console.log(`✅ Chunk ${Math.floor(i/chunkSize) + 1} added.`);
            } else {
                console.error(`⚠️ Chunk ${Math.floor(i/chunkSize) + 1} invalid format.`);
            }
            
            await new Promise(r => setTimeout(r, 1500)); 
        }

        // 3. إنشاء ملف إكسيل
        if (finalRows.length > 0) {
            console.log("Creating Excel file...");
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Data');
            sheet.addRows(finalRows);
            
            sheet.getRow(1).font = { bold: true };
            sheet.getRow(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD3D3D3' }
            };

            await workbook.xlsx.writeFile('Result.xlsx');
            console.log("✅ Success! Result.xlsx is ready.");
        }

    } catch (globalError) {
        console.error("❌ Global Script Error:", globalError.message);
        process.exit(1);
    }
}

start();
