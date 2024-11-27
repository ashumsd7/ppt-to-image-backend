import axios from 'axios';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import pptPng from 'ppt-png';

// Handle __dirname in ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Function to download a PPT file
async function downloadPPT(url, outputPath) {
    const writer = fs.createWriteStream(outputPath);
    const response = await axios({
        url: url,
        method: 'GET',
        responseType: 'stream',
    });

    response.data.pipe(writer);

    return new Promise((resolve, reject) => {
        writer.on('finish', resolve);
        writer.on('error', reject);
    });
}

// Convert PPT to PNGs
async function convertPPTToImages(pptPath, outputDir) {
    try {
        const converter = new pptPng(pptPath, outputDir);
        const images = await converter.convert();
        console.log(`Conversion completed! Images saved in: ${outputDir}`);
        console.log(images); // List of image paths
    } catch (err) {
        console.error('Error during conversion:', err);
    }
}

// Main function to download and process the PPT
async function main() {
    const pptUrl = 'https://codemonkpptgen.blob.core.windows.net/cmpptgencontainerv1/Slide_text_379_1732692286147.pptx'; // Replace with your PPT URL
    const pptPath = path.join(__dirname, 'Slide_text_379_1732692286147.pptx');
    const outputDir = path.join(__dirname, 'images');

    // Ensure the output directory exists
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    console.log('Downloading PPT file...');
    await downloadPPT(pptUrl, pptPath);

    console.log('Converting PPT to images...');
    await convertPPTToImages(pptPath, outputDir);
}

main().catch(console.error);
