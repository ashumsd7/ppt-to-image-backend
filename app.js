import express from "express";
import fs from "fs";
import path from "path";
import axios from "axios";
import { exec } from "child_process";
import { BlobServiceClient } from "@azure/storage-blob";
import { fileURLToPath } from "url";
import cors from "cors"; // Import the cors package
import { promises as fs2 } from "node:fs";
import { pdf } from "pdf-to-img";
// Initialize Express app
const app = express();
const port = 3000;

// Handle __dirname in ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Enable CORS for all origins
app.use(cors()); // Allow all origins by default

// Middleware to parse JSON bodies
app.use(express.json());

// Function to download PPTX file from URL
async function downloadPPT(url, outputPath) {
  const writer = fs.createWriteStream(outputPath);
  const response = await axios({
    url: url,
    method: "GET",
    responseType: "stream",
  });

  response.data.pipe(writer);

  return new Promise((resolve, reject) => {
    writer.on("finish", resolve);
    writer.on("error", reject);
  });
}

// Function to convert PPTX to PDF
const pptxToPdf = async (pptxPath, pptxFileName) => {
  try {
    // Step 1: Verify if the PPTX file exists
    if (!fs.existsSync(pptxPath)) {
      console.error("The specified PPTX file does not exist:", pptxPath);
      return;
    }

    console.log("Found the PPTX file:", pptxPath);

    // Step 2: Set the path for saving the PDF
    const desktopPath = path.join(process.env.USERPROFILE, "Desktop"); // Desktop path for Windows
    const pdfPath = path.join(desktopPath, `${pptxFileName}.pdf`); // Use the same name for PDF

    console.log("Converting PPTX to PDF...");

    // Convert PPTX to PDF using LibreOffice (fixed path format for Windows)
    const command = `soffice --headless --convert-to pdf "${pptxPath}" --outdir "tempPdfFolder"`;

    return new Promise((resolve, reject) => {
      exec(command, (err, stdout, stderr) => {
        if (err) {
          console.error("Error during conversion:", err.message);
          reject("Error during conversion");
        } else {
          console.log(
            `PPTX converted to PDF successfully! PDF saved as: ${pdfPath}`
          );
          resolve(pdfPath);
        }
      });
    });
  } catch (error) {
    console.error("Error:", error.message);
    throw error;
  }
};


// Express endpoint to handle PPTX URL conversion
app.post("/convert-pptx", async (req, res) => {
  const { pptxUrl } = req.body;

  console.log("URL Received ",pptxUrl);

  if (!pptxUrl) {
    return res.status(400).json({ error: "pptxUrl is required" });
  }

  try {
    // Get the base name of the PPTX file from the URL (without extension)
    const pptxFileName = path.basename(pptxUrl, ".pptx");
    const pptxPath = path.join(__dirname, `${pptxFileName}.pptx`);

    // Step 1: Download the PPTX file
    console.log("Downloading PPTX from URL...");
    await downloadPPT(pptxUrl, pptxPath);

    // Step 2: Convert PPTX to PDF
    console.log("Converting PPTX to PDF...");
    const pdfPath = await pptxToPdf(pptxPath, pptxFileName);
    console.log("Pdf saved locally");

    // Step 3: Upload the PDF to Blob Storage
    console.log(
      "now calling main to get the pdf and then extratc image and upload it"
    );
    const filePath = path.join("tempPdfFolder", `Codemonk.pdf`);
    const uploadedPdfURL = await uploadFileToBlobForPdf(filePath, `${pptxFileName}.pdf`);
    console.log("uploadedPdfURL",uploadedPdfURL);
    const {uploadedUrls, pdfUrl} = await main(uploadedPdfURL);

    console.log("uploadedUrls", uploadedUrls);

    // Return the PDF URL in the response
    return res.json({ uploadedUrls,pdfUrl });
  } catch (error) {
    console.error("Error during conversion and upload:", error);
    return res.status(500).json({ error: "Error during conversion or upload" });
  }
});

async function uploadFileToBlobForPdf(filePath, fileName) {
  try {
    const sasToken =
      "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2025-06-04T13:26:22Z&st=2024-11-05T05:26:22Z&spr=https,http&sig=pAcLQDyT%2BRNtUABOSobtIhb%2FuSA43rbiU0btYf%2FVttw%3D";
    const containerName = `cmpptgencontainerv1`;
    const storageAccountName = "codemonkpptgen";

    // Create a BlobServiceClient
    const blobServiceClient = new BlobServiceClient(
      `https://${storageAccountName}.blob.core.windows.net/?${sasToken}`
    );

    // Get a container client
    const containerClient = blobServiceClient.getContainerClient(containerName);

    // Create a block blob client for the file
    const blockBlobClient = containerClient.getBlockBlobClient(fileName);

    // Read the file as a buffer
    const fileBuffer = await fs2.readFile(filePath);

    // Set the Content-Type to 'image/png' for PNG files
    // const contentType = "appl/png";

    // Upload the Blob with the correct Content-Type
    await blockBlobClient.upload(fileBuffer, fileBuffer.length);

    // Generate the file URL
    const fileUrl = `https://${storageAccountName}.blob.core.windows.net/${containerName}/${fileName}`;
    console.log("File uploaded successfully! File URL:", fileUrl);
    return fileUrl;
  } catch (error) {
    console.error("Error uploading file:", error);
    throw error;
  }
}

// ---------------------- starts main

async function uploadFileToBlob(filePath, fileName) {
  try {
    const sasToken =
      "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2025-06-04T13:26:22Z&st=2024-11-05T05:26:22Z&spr=https,http&sig=pAcLQDyT%2BRNtUABOSobtIhb%2FuSA43rbiU0btYf%2FVttw%3D";
    const containerName = `cmpptgencontainerv1`;
    const storageAccountName = "codemonkpptgen";

    // Create a BlobServiceClient
    const blobServiceClient = new BlobServiceClient(
      `https://${storageAccountName}.blob.core.windows.net/?${sasToken}`
    );

    // Get a container client
    const containerClient = blobServiceClient.getContainerClient(containerName);

    // Create a block blob client for the file
    const blockBlobClient = containerClient.getBlockBlobClient(fileName);

    // Read the file as a buffer
    const fileBuffer = await fs2.readFile(filePath);

    // Set the Content-Type to 'image/png' for PNG files
    const contentType = "image/png";

    // Upload the Blob with the correct Content-Type
    await blockBlobClient.upload(fileBuffer, fileBuffer.length, {
      blobHTTPHeaders: { blobContentType: contentType },
    });

    // Generate the file URL
    const fileUrl = `https://${storageAccountName}.blob.core.windows.net/${containerName}/${fileName}`;
    console.log("File uploaded successfully! File URL:", fileUrl);
    return fileUrl;
  } catch (error) {
    console.error("Error uploading file:", error);
    throw error;
  }
}

async function main(pdfUrl) {
  try {
    // Get the current directory path using `import.meta.url` in ES Modules
    const __dirname = path.dirname(new URL(import.meta.url).pathname); // Correctly extract the directory path

    // Create the directory to save images
    const folderName = "temp_images";
    const outputDir = path.join(__dirname, folderName); // Correctly join __dirname and folderName

    // Ensure the directory exists, create if it doesn't
    await fs2.mkdir("outputDir", { recursive: true });

    console.log(`Created folder: ${outputDir}`);

    let counter = 1;
    const document = await pdf("./tempPdfFolder/codemonk.pdf", { scale: 3 });

    // Create an array to store the uploaded URLs
    const uploadedUrls = [];

    // Loop through each page and save as PNG images inside the 'temp_images' folder
    for await (const image of document) {
      const imagePath = path.join("outputDir", `page${counter}.png`);
      console.log("Saving image to:", imagePath);

      // Save the image temporarily
      await fs2.writeFile(imagePath, image);
      console.log(`Saved: ${imagePath}`);

      // Upload the image to Azure Blob Storage and get the URL
      const fileName = `page${Date.now()}.png`; // Using Date.now() to generate a unique file name
      const uploadedUrl = await uploadFileToBlob(imagePath, fileName);
      uploadedUrls.push(uploadedUrl);

      console.log("Uploaded URL:", uploadedUrl);
      counter++;
    }

    // After all images are uploaded, log the list of URLs
    console.log("All uploaded URLs:", uploadedUrls);
    return {uploadedUrls,pdfUrl};
  } catch (error) {
    console.error("Error:", error.message);
  }
}

// ------------------------end main

// Start the server
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
