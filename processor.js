const axios = require('axios');
const ExcelJS = require('exceljs');

async function start() {
    const fileId = process.env.FILE_ID;
    const gasUrl = process.env.GAS_URL;
    const groqKey = process.env.GROQ_API_KEY;

    try {
        // 1. جلب النص من جوجل (OCR)
        console.log("Fetching text from Google OCR...");
        const response = await axios.get(`${gasUrl}?fileId=${fileId}`);
        const fullText = response.data;

        if (!fullText || fullText.includes("Error")) {
            throw new Error("Failed to get text from Google OCR: " + fullText);
        }

        // 2. تقسيم النص لقطع (Chunks)
        const lines = fullText.split('\n').filter(l => l.trim());
        const chunkSize = 30; 
        let finalRows = [];

        console.log(`Processing ${lines.length} lines in chunks...`);

        for (let i = 0; i < lines.length; i += chunkSize) {
            const chunk = lines.slice(i, i + chunkSize).join('\n');
            const isFirstChunk = (i === 0);
            
            console.log(`Sending chunk ${Math.floor(i/chunkSize) + 1}...`);

            try {
                const res = await axios.post('https://api.groq.com/openai/v1/chat/completions', {
                    model: "llama-3.3-70b-versatile",
                    messages: [{
                        role: "system", 
                        content: "You are a data extractor. Convert OCR text into a 2D JSON array [[row1], [row2]]."
                    }, {
                        role: "user", 
                        content: `Convert this text to a 2D JSON array. ${isFirstChunk ? "Include headers in the first row." : "No headers."} Return ONLY the JSON. Text:\n${chunk}`
                    }],
                    temperature: 0
                }, {
                    headers: { 'Authorization': `Bearer ${groqKey}`, 'Content-Type': 'application/json' }
                });

                let rawContent = res.data.choices[0].message.content;
                
                // تنظيف الرد باستخدام Regex لاستخراج المصفوفة فقط
                const jsonMatch = rawContent.match(/\[\s*\[[\s\S]*\]\s*\]/);
                
                if (jsonMatch) {
                    const data = JSON.parse(jsonMatch[0]);
                    finalRows.push(...data);
                    console.log(`✅ Chunk ${Math.floor(i/chunkSize) + 1} added (${data.length} rows).`);
                } else {
                    console.error(`⚠️ Chunk ${Math.floor(i/chunkSize) + 1} returned invalid format.`);
                }
                
                // انتظر ثانية ونصف لتجنب الـ Rate Limit (أضمن للـ 800 صفحة)
                await new Promise(r => setTimeout(r, 1500)); 
            } catch (chunkError) {
                console.error(`❌ Error in chunk ${Math.floor(i/chunkSize) + 1}:`, chunkError.message);
            }
        }

        // 3. إنشاء ملف إكسيل
        if (finalRows.length > 0) {
            console.log("Creating Excel file...");
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Data');
            
            // إضافة البيانات
            sheet.addRows(finalRows);
            
            // تنسيق بسيط لأول صف (العناوين)
            sheet.getRow(1).font = { bold: true };
            sheet.getRow(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD3D3D3' }
            };

            await workbook.xlsx.writeFile('Result.xlsx');
            console.log("✅ Success! Excel file created: Result.xlsx");
        } else {
            console.log("❌ No data was processed. Check Groq logs.");
        }

    } catch (globalError) {
        console.error("❌ Global Script Error:", globalError.message);
        process.exit(1);
    }
}

start();
