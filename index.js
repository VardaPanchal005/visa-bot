require("dotenv").config();

const axios = require("axios");
const cheerio = require("cheerio");
const XLSX = require("xlsx");
const nodemailer = require("nodemailer");

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL,
    pass: process.env.PASSWORD,
  },
});

async function getLatestODS() {
  const res = await axios.get(process.env.PAGE_URL, {
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
      "Accept": "text/html,application/xhtml+xml",
      "Accept-Language": "en-US,en;q=0.9",
      "Referer": "https://www.google.com",
    },
  });

  const $ = cheerio.load(res.data);

  const baseUrl = "https://www.ireland.ie";
  let link = null;

  $("a").each((_, el) => {
    const href = $(el).attr("href");

    if (href && href.includes("NDVO") && href.endsWith(".ods")) {
      link = href.startsWith("http")
        ? href
        : new URL(href, baseUrl).href; 
    }
  });

  return link;
}

async function fetchData(odsUrl) {
 const res = await axios.get(odsUrl, {
  responseType: "arraybuffer",
  headers: {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
    "Accept": "*/*",
    "Referer": "https://www.google.com"
  }
});

  const workbook = XLSX.read(res.data, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

async function checkIRL() {
  try {
    const myIRL = String(process.env.MY_IRL).trim();

    const link = await getLatestODS();
    if (!link) {
      console.log("No ODS link found");
      return;
    }

    console.log("Using file:", link);

    const data = await fetchData(link);

    let status = "NOT_FOUND";

    for (let row of data) {
  if (!Array.isArray(row)) continue;

  const appNumber = String(row[0] || "").trim();
  const decision = String(row[1] || "").trim();

  if (
    appNumber.toLowerCase().includes("application") ||
    appNumber.toLowerCase().includes("visa") ||
    appNumber === ""
  ) {
    continue;
  }

  if (appNumber === myIRL) {
    status = decision || "FOUND_NO_DECISION";
    break;
  }
}

    console.log("Status:", status);

    if (status === "Refused") {
      await sendEmail("❌ Visa Refused", link, status);
    } 
    else if (status === "Approved") {
      await sendEmail("✅ Visa Approved", link, status);
    } 
    else if (status === "FOUND_NO_DECISION") {
      console.log("Found but no decision column filled");
    } 
    else {
      await sendEmail("⏳ No Decision Yet", link, status);
    }

  } catch (err) {
    console.error(err);
  }
}

async function sendEmail(subject, link, status) {
  await transporter.sendMail({
    from: process.env.EMAIL,
    to: process.env.TO_EMAIL,
    subject,
    text: `Status: ${status}\n\nLink: ${link}`,
  });

  console.log("Email sent:", subject);
}

checkIRL();
