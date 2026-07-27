const { generateInterviewReport } = require("./src/services/ai.service");
const puppeteer = require("puppeteer-core");
const chromium = require("@sparticuz/chromium");
const fs = require("fs");

async function testLocalPdf() {
    const isLocal = process.platform === "win32" || process.platform === "darwin";
    console.log("Is Local:", isLocal);
    
    let launchOptions = {};
    if (!isLocal) {
        launchOptions = {
            args: chromium.args,
            defaultViewport: chromium.defaultViewport,
            executablePath: await chromium.executablePath(),
            headless: chromium.headless,
        };
    } else {
        launchOptions = {
            headless: true,
            args: [
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu"
            ]
        };
        if (process.platform === "win32") {
            const possiblePaths = [
                "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
                "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
                `${process.env.LOCALAPPDATA}\\Google\\Chrome\\Application\\chrome.exe`
            ];
            for (const path of possiblePaths) {
                if (fs.existsSync(path)) {
                    launchOptions.executablePath = path;
                    break;
                }
            }
        }
    }
    
    console.log("Launch Options:", launchOptions);
    
    try {
        const browser = await puppeteer.launch(launchOptions);
        const page = await browser.newPage();
        await page.setContent("<h1>Hello World from PDF</h1>", { waitUntil: "networkidle0" });
        const pdf = await page.pdf({ format: "A4" });
        await browser.close();
        console.log("Success! PDF generated. Length:", pdf.length);
        fs.writeFileSync("test.pdf", pdf);
        console.log("PDF saved to test.pdf");
    } catch (err) {
        console.error("Failed to generate PDF:", err);
    }
}

testLocalPdf();
