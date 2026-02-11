// Core Node.js modules
const fs = require("fs");               // File system access
const path = require("path");           // File path utilities
const { execSync } = require("child_process"); // Run system commands

// External library for reading .docx files
const mammoth = require("mammoth");

// Get folder path from command line argument
// Example: node script.js "E://MyFolder"
// If not provided, use default path
const folderPath = process.argv[2] || "your-folder-path";

// Log which folder is being processed
console.log(`Processing documents in folder: ${folderPath}`);

/**
 * Counts words in a string
 * - Splits text by whitespace
 * - Filters out empty values
 */
function countWords(text) {
    return text
        .split(/\s+/)
        .filter(word => word.length > 0)
        .length;
}

/**
 * Main async function to process all documents
 */
async function processDocuments() {
    let totalWordCount = 0;

    // 1. Get and Sort files (This part stays the same)
    const files = fs.readdirSync(folderPath)
        .map(file => ({
            name: file,
            number: parseInt(file.match(/^(\d+)-/)?.[1] || "0", 10)
        }))
        .sort((a, b) => a.number - b.number)
        .map(file => file.name);

    // 2. We will now process files one-by-one (Sequentially) 
    // to guarantee the order in the 'results' array.
    const results = [];

    for (const file of files) {
        const filePath = path.join(folderPath, file);
        const fileStats = fs.statSync(filePath);

        if (!fileStats.isFile()) continue;

        // PROCESS .DOCX
        if (file.endsWith(".docx")) {
            try {
                const buffer = fs.readFileSync(filePath);
                // We use 'await' here so the script waits for File 1 
                // to finish before moving to File 2
                const result = await mammoth.extractRawText({ buffer });
                const wordCount = countWords(result.value);
                
                const message = `Processed ${file}: ${wordCount} words`;
                console.log(message);
                results.push(message);
                totalWordCount += wordCount;
            } catch (err) {
                console.error(`Error in ${file}: ${err.message}`);
            }
        }

        // PROCESS .DOC (This part is already sequential)
        else if (file.endsWith(".doc")) {
            try {
                const wordCount = parseInt(
                    execSync(`powershell -command "& {
                        $word = New-Object -ComObject Word.Application;
                        $doc = $word.Documents.Open('${filePath}');
                        $count = $doc.Words.Count;
                        $doc.Close();
                        $word.Quit();
                        Write-Output $count;
                    }"`).toString().trim(), 10
                );

                const message = `Processed ${file}: ${wordCount} words`;
                console.log(message);
                results.push(message);
                totalWordCount += wordCount;
            } catch (error) {
                console.error(`Error processing ${file}: ${error.message}`);
            }
        }
    }

    // 3. Save Final Total
    const totalText = `Total word count across all files: ${totalWordCount}`;
    console.log(totalText);
    results.push(totalText);

    // 4. Write to file (This part stays the same)
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const resultsFilePath = path.join(folderPath, `word-count-results-${timestamp}.txt`);
    fs.writeFileSync(resultsFilePath, results.join("\n"));

    console.log(`\nResults saved to: ${resultsFilePath}`);
}

// Run the script
processDocuments().catch(err => {
    console.error("An error occurred:", err);
});
