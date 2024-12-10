const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const path = require('path');
const axios = require('axios');
const fs = require('fs');
require('dotenv').config(); // Load environment variables

// Read API keys from environment variables
const CAPTCHA_API_KEY = process.env.CAPTCHA_API_KEY;

// Custom delay function to simulate human typing speed with random delay
async function delay(minTime, maxTime) {
    const randomDelay = Math.floor(Math.random() * (maxTime - minTime + 1)) + minTime;
    return new Promise(resolve => setTimeout(resolve, randomDelay));
}

// Function to generate a random 10-character unique ID
function generateUniqueID() {
    const characters = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ$@_';
    let id = '';
    for (let i = 0; i < 10; i++) {
        id += characters.charAt(Math.floor(Math.random() * characters.length));
    }
    console.log('Generated Unique ID:', id);  // This line displays the generated ID in the console
    return id;
}

// Function to solve CAPTCHA using 2Captcha API
async function solveCaptcha(page) {
    const siteKey = await page.$eval('.g-recaptcha', el => el.getAttribute('data-sitekey'));
    console.log('Site Key:', siteKey);

    const requestUrl = `http://2captcha.com/in.php?key=${CAPTCHA_API_KEY}&method=userrecaptcha&googlekey=${siteKey}&pageurl=${page.url()}`;
    const response = await axios.get(requestUrl);

    if (!response.data.includes('OK|')) {
        throw new Error('Failed to submit CAPTCHA request: ' + response.data);
    }

    const captchaId = response.data.split('|')[1];
    console.log('CAPTCHA ID:', captchaId);

    let captchaSolution;

    // Poll for the CAPTCHA solution
    while (true) {
        const resultUrl = `http://2captcha.com/res.php?key=${CAPTCHA_API_KEY}&action=get&id=${captchaId}`;
        const resultResponse = await axios.get(resultUrl);

        if (resultResponse.data === 'CAPCHA_NOT_READY') {
            console.log('CAPTCHA not ready yet, waiting...');
            await delay(5000, 7000); // Wait 5-7 seconds
            continue;
        }

        if (resultResponse.data.includes('OK|')) {
            captchaSolution = resultResponse.data.split('|')[1];
            console.log('CAPTCHA Solved:', captchaSolution);
            break;
        } else {
            throw new Error('Failed to retrieve CAPTCHA solution: ' + resultResponse.data);
        }
    }

    return captchaSolution;
}

// Function to handle the manual resume upload
async function uploadResumeManually(page) {
    const fileInput = await page.$('input[type="file"]');
    if (!fileInput) {
        throw new Error('File input field not found');
    }

    console.log('Waiting for the user to manually upload the resume...');
    // Wait for the file input to be populated (user has uploaded the file)
    await page.waitForFunction('document.querySelector("input[type=\'file\']").files.length > 0', { timeout: 0 });
    console.log('File uploaded manually by the user!');
}

// Function to process each row of the Excel file and fill the form
async function processRow(browser, row) {
    const page = await browser.newPage();
    const startTime = Date.now(); // Track start time of this row

    try {
        await page.goto('https://login-c-chi-52.vercel.app/', { waitUntil: 'domcontentloaded' });

        const [regdNo, name, email, branch, gender, account, number] = row;

        // Start filling the form with the data
        await page.type('#regd_no', regdNo); // Registration Number
        console.log('Filling Registration Number:', regdNo);
        await delay(1000, 2000); // Random delay between 1-2 seconds

        await page.type('#name', name); // Name
        console.log('Filling Name:', name);
        await delay(1000, 2000); // Random delay between 1-2 seconds

        await page.type('#email', email); // Email
        console.log('Filling Email:', email);
        await delay(1000, 2000); // Random delay between 1-2 seconds

        await page.select('#branch', branch); // Branch
        console.log('Selecting Branch:', branch);
        await delay(1000, 2000); // Random delay between 1-2 seconds

        // Gender selection (Male or Female)
        if (gender === 'Male') {
            await page.click('input[name="gender"][value="male"]');
            console.log('Selecting Male gender');
        } else if (gender === 'Female') {
            await page.click('input[name="gender"][value="female"]');
            console.log('Selecting Female gender');
        }

        await delay(1000, 2000); // Random delay between 1-2 seconds

        // Bank account selection and handling input
        if (account.toLowerCase() === "yes") {
            await page.click('input[name="bank_account"][value="yes"]');
            console.log("Selected Yes for Bank Account");

            await page.waitForSelector("#account_number", { visible: true });
            await page.type("#account_number", number);
            console.log("Filling IBAN Number:", number);
        } else {
            await page.click('input[name="bank_account"][value="no"]');
            console.log("Selected No for Bank Account");

            await page.waitForSelector("#wallet_number", { visible: true });
            await page.type("#wallet_number", number);
            console.log("Filling Binance Account Number:", number);
        }

        await delay(1000, 2000); // Random delay between 1-2 seconds

        // Wait for user to manually upload the resume
        await uploadResumeManually(page);

        // Solve CAPTCHA
        const captchaSolution = await solveCaptcha(page);
        await page.evaluate((captchaSolution) => {
            document.querySelector('textarea[name="g-recaptcha-response"]').value = captchaSolution;
        }, captchaSolution);
        console.log('CAPTCHA Solved and injected');

        // Submit the form
        await page.click('input[type="submit"]');
        console.log('Form Submitted');
        await page.waitForNavigation({ waitUntil: 'load' });
        console.log('Form Submitted Successfully!');

        // Generate and return a random ID
        const randomId = generateUniqueID();
        console.log('Generated Unique ID:', randomId); // Added log to explicitly print the ID

        const endTime = Date.now(); // Track end time of this row
        console.log(`Time taken for this row: ${(endTime - startTime) / 1000} seconds`);

        return randomId;
    } catch (err) {
        console.error('Error processing row:', row, err);
        return null;
    }
}

// Main function to read data from Excel and process each row
async function fillForm() {
    const browser = await puppeteer.launch({ headless: false });
    const overallStartTime = Date.now(); // Start timer for overall execution

    try {
        // Read the data from Excel file
        const workbook = xlsx.readFile(path.join(__dirname, 'loginfile.xlsx'));
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        // Ensure there's data beyond the headers
        if (data.length < 2) {
            console.error('No data found in Excel file.');
            return;
        }

        // Add "ID" header if it doesn't exist
        if (!data[0].includes('ID')) {
            data[0].push('ID');
        }

        // Process each row and collect random IDs
        for (let i = 1; i < data.length; i++) {
            console.log(`Processing row ${i}:`, data[i]);
            const randomId = await processRow(browser, data[i]);
            if (randomId !== null) {
                data[i][7] = randomId; // Add the unique ID to the 8th column (H)
            }
        }

        // Write the updated data with IDs to a new Excel file
        const updatedSheet = xlsx.utils.aoa_to_sheet(data);
        workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;
        const updatedFilePath = path.join(__dirname, 'loginfile_with_ids.xlsx');
        xlsx.writeFile(workbook, updatedFilePath);
        console.log('Updated data written to:', updatedFilePath);
    } catch (error) {
        console.error('Error in fillForm:', error);
    } finally {
        const overallEndTime = Date.now();
        console.log(`Total time for execution: ${(overallEndTime - overallStartTime) / 1000} seconds`);
        await browser.close();
    }
}

// Execute the script
fillForm().catch(err => console.error('Error in fillForm:', err));
