import { promises as fs } from "node:fs";
import path from "path";
import { pdf } from "pdf-to-img";
import { BlobServiceClient } from "@azure/storage-blob"; // Azure Blob Storage SDK

// Function to upload a file to Azure Blob Storage
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
    const fileBuffer = await fs.readFile(filePath);

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

async function main() {
  try {
    // Get the current directory path using `import.meta.url` in ES Modules
    const __dirname = path.dirname(new URL(import.meta.url).pathname); // Correctly extract the directory path

    // Create the directory to save images
    const folderName = "temp_images";
    const outputDir = path.join(__dirname, folderName);  // Correctly join __dirname and folderName

    // Ensure the directory exists, create if it doesn't
    await fs.mkdir('outputDir', { recursive: true });

    console.log(`Created folder: ${outputDir}`);

    let counter = 1;
    const document = await pdf("./sample.pdf", { scale: 3 });

    // Create an array to store the uploaded URLs
    const uploadedUrls = [];

    // Loop through each page and save as PNG images inside the 'temp_images' folder
    for await (const image of document) {
      const imagePath = path.join('outputDir', `page${counter}.png`);
      console.log("Saving image to:", imagePath);

      // Save the image temporarily
      await fs.writeFile(imagePath, image);
      console.log(`Saved: ${imagePath}`);

      // Upload the image to Azure Blob Storage and get the URL
      const fileName = `page${Date.now()}.png`;  // Using Date.now() to generate a unique file name
      const uploadedUrl = await uploadFileToBlob(imagePath, fileName);
      uploadedUrls.push(uploadedUrl);

      console.log("Uploaded URL:", uploadedUrl);
      counter++;
    }

    // After all images are uploaded, log the list of URLs
    console.log("All uploaded URLs:", uploadedUrls);

  } catch (error) {
    console.error("Error:", error.message);
  }
}

// Execute the main function
main().catch(console.error);
