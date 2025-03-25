const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");
const { execSync } = require("child_process");

// Get folder path from command line argument or use default
const folderPath = process.argv[2] || "D://hashim-bhai-work/to-be-paid/folder-6";

// Display the folder path being processed
console.log(`Processing documents in folder: ${folderPath}`);

function countWords(text) {
    return text.split(/\s+/).filter(word => word.length > 0).length;
}

async function processDocuments() {
    let totalWordCount = 0;
    const files = fs.readdirSync(folderPath)
        .map(file => ({
            name: file,
            number: parseInt(file.match(/^(\d+)-/)?.[1] || '0', 10) // Extracts leading number
        }))
        .sort((a, b) => a.number - b.number)
        .map(file => file.name); 

    const processingPromises = [];
    const results = []; // Stores only successful results

    for (const file of files) {
        const filePath = path.join(folderPath, file);
        const fileStats = fs.statSync(filePath);
        
        if (!fileStats.isFile()) continue;

        if (file.endsWith(".docx")) {
            const processPromise = new Promise((resolve) => {
                try {
                    const buffer = fs.readFileSync(filePath);
                    mammoth.extractRawText({ buffer })
                        .then(result => {
                            const wordCount = countWords(result.value);
                            const resultText = `Processed ${file}: ${wordCount} words`;
                            console.log(resultText);
                            results.push(resultText); // Save only success messages
                            resolve(wordCount);
                        })
                        .catch(err => {
                            console.error(`Error in ${file}: ${err.message}`); // Log only in console
                            resolve(0);
                        });
                } catch (error) {
                    console.error(`Error reading ${file}: ${error.message}`); // Log only in console
                    resolve(0);
                }
            });
            processingPromises.push(processPromise);
        } else if (file.endsWith(".doc")) {
            try {
                const wordCount = parseInt(
                    execSync(
                        `powershell -command "& { $word = New-Object -ComObject Word.Application; $doc = $word.Documents.Open('${filePath}'); $count = $doc.Words.Count; $doc.Close(); $word.Quit(); Write-Output $count; }"`
                    ).toString().trim(), 
                    10
                );
                const resultText = `Processed ${file}: ${wordCount} words`;
                console.log(resultText);
                results.push(resultText);
                totalWordCount += wordCount;
            } catch (error) {
                console.error(`Error processing ${file}: ${error.message}`); // Log only in console
            }
        }
    }

    const wordCounts = await Promise.all(processingPromises);
    totalWordCount += wordCounts.reduce((sum, count) => sum + count, 0);

    const totalText = `Total word count across all files: ${totalWordCount}`;
    console.log(totalText);
    results.push(totalText); // Add only successful total count

    // Save results without errors
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const resultsFilePath = path.join(folderPath, `word-count-results-${timestamp}.txt`);
    fs.writeFileSync(resultsFilePath, results.join('\n'));

    console.log(`\nResults have been saved to: ${resultsFilePath}`);
}


// Run the main function
processDocuments().catch(err => {
    console.error("An error occurred:", err);
});
