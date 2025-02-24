document.getElementById('processDocx').addEventListener('click', async () => {
    const fileInput = document.getElementById('uploadDocx');
    const downloadLink = document.getElementById('downloadLink');

    if (!fileInput.files.length) {
        alert("Please upload a .docx file first.");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = async function(event) {
        try {
            const zip = new JSZip();
            await zip.loadAsync(event.target.result);

            if (!zip.file("word/comments.xml")) {
                alert("No comments found in the document.");
                return;
            }

            // Load comments.xml and parse to JSON
            const commentsXml = await zip.file("word/comments.xml").async("string");
            let commentsJson = xmljs.xml2js(commentsXml, { compact: true });

            // Remove timestamps from all comments
            if (commentsJson["w:comments"] && commentsJson["w:comments"]["w:comment"]) {
                let comments = commentsJson["w:comments"]["w:comment"];
                if (!Array.isArray(comments)) comments = [comments];

                comments.forEach(comment => {
                    if (comment._attributes && comment._attributes["w:date"]) {
                        delete comment._attributes["w:date"];
                    }
                });

                // Convert back to XML and update the file in ZIP
                const updatedXml = xmljs.js2xml(commentsJson, { compact: true, spaces: 4 });
                zip.file("word/comments.xml", updatedXml);
            }

            // Repackage into a new .docx file
            const cleanedBlob = await zip.generateAsync({ type: "blob" });
            const url = URL.createObjectURL(cleanedBlob);

            // Show download link
            downloadLink.href = url;
            downloadLink.download = "cleaned_" + file.name;
            downloadLink.style.display = "block";
            downloadLink.textContent = "Download Cleaned DOCX";
        } catch (error) {
            console.error("Error processing file:", error);
            alert("Failed to process the document. Ensure it's a valid .docx file.");
        }
    };

    reader.readAsArrayBuffer(file);
});
