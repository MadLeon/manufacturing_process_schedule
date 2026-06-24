// JavaScript file for Acrobat Pro to merge PDF files
// This script merges multiple PDF files into one

// Function to merge PDFs
function mergePDFs(sourceFiles, outputFile) {
    try {
        // Open the first PDF as the base document
        var baseDoc = app.open(sourceFiles[0]);
        
        if (baseDoc == null) {
            console.println("Error: Could not open base PDF: " + sourceFiles[0]);
            return false;
        }
        
        // Insert pages from remaining PDFs
        for (var i = 1; i < sourceFiles.length; i++) {
            try {
                console.println("Merging: " + sourceFiles[i]);
                
                // Insert pages from current PDF at the end
                var insertDoc = app.open(sourceFiles[i]);
                if (insertDoc != null) {
                    // Insert all pages from insertDoc into baseDoc
                    baseDoc.insertPages({
                        nIndex: baseDoc.numPages - 1,
                        cPath: sourceFiles[i]
                    });
                    
                    insertDoc.close(false);
                }
            } catch (e) {
                console.println("Warning: Could not merge " + sourceFiles[i] + ": " + e);
            }
        }
        
        // Save the merged PDF
        baseDoc.saveAs({
            cPath: outputFile
        });
        
        baseDoc.close(false);
        
        console.println("Successfully saved merged PDF: " + outputFile);
        return true;
        
    } catch (e) {
        console.println("Error in mergePDFs: " + e);
        return false;
    }
}

// Main script execution
try {
    // Get parameters from command line arguments
    // The parameters will be passed as: pdfMerge.js file1.pdf file2.pdf ... output.pdf
    
    console.println("Acrobat PDF Merge Script started");
    
    // Note: When calling this script from PowerShell, pass all PDF files and output file
    // Example: acrobat.exe /s /n pdfMerge.js arg1 arg2 ...
    
} catch (e) {
    console.println("Script error: " + e);
}
