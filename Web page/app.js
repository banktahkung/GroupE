const express = require("express");
const bodyParser = require("body-parser");
const dotenv = require("dotenv");
const sha256 = require("js-sha256").sha256;
const { createClient } = require("@supabase/supabase-js");
const session = require("express-session");
const nodemailer = require("nodemailer");

// ! DONT FORGOT TO DELETE THE CODE WITH STATEMENT `console.log() IN PRODUCTION MODE`
// ! NEVER SHARE THE `.env` FILE IN PUBLIC
// % Port Use
const PORT = process.env.PORT || 4000;

// - App
const app = express();

// - Get env content
const env = dotenv.config().parsed;

// - Establish connection between database and server
// const supabase = createClient(env.SUPABASE_URL, env.SUPABASE_KEY);

console.log(env);

// - Header of the table
const OutputTableHeader = [
  "DOCUMENT_NAME_Template",
  "DOCUMENT_TYPE",
  "DOCUMENT_NO",
  "DOCUMENT_DATE",
  "TAX_TYPE",
  "VAT_RATE",
  "AMOUNT",
  "BASIS_AMOUNT",
  "BASIS_AMOUNT_Template",
  "CALCULATED_AMOUNT",
  "TAX_BASIS_AMOUNT",
  "VAT",
  "VAT_Template",
  "GRAND_TOTAL",
  "GRAND_TOTAL_Template",
  "ALLOWANCE_CHARGE",
  "AC_AMOUNT",
  "REASON",
  "REASON_CODE",
  "ALLOWANCE_CODE",
  "SELLER_NAME",
  "SELLER_TAX_ID_TYPE",
  "SELLER_TAX_ID",
  "SELLER_BRANCH_CODE",
  "SELLER_ADDRESS",
  "SELLER_BUILDING_NO",
  "SELLER_BUILDING_NAME",
  "SELLER_STREET_NAME",
  "SELLER_DISTRICT",
  "SELLER_CITY",
  "SELLER_PROVINCE",
  "SELLER_POSTAL_CODE",
  "SELLER_EMAIL",
  "BUYER_NAME",
  "BUYER_TAX_ID_TYPE",
  "BUYER_TAX_ID",
  "BUYER_BRANCH_CODE",
  "BUYER_EMAIL",
  "BUYER_BUILDING_NO",
  "BUYER_BUILDING_NAME",
  "BUYER_STREET_NAME",
  "BUYER_DISTRICT",
  "BUYER_CITY",
  "BUYER_PROVINCE",
  "BUYER_ADDRESS",
  "BUYER_POSTAL_CODE",
  "BUYER_COUNTRY_CODE",
  "BUYER_PHONE_NO",
  "REFERENCED",
  "REFERENCED_DATE",
  "REFERENCED_TYPE",
  "ORIGINAL_AMOUNT",
  "DIFFERENCE_AMOUNT",
  "PURPOSE",
  "PURPOSE_CODE",
  "PROJECT_NAME",
  "PAYMENT_TERM",
  "DUE_DATE",
  "REFERENCE_PO(QT.NO.)",
  "CURRENCY_Template",
  "TEXT_AMOUNT",
  "WITHOLDING_TAX_DEDUCTION_SENTENCE",
  "BANK_DETAILS",
  "LINE_NO",
  "ITEM_CODE",
  "ITEM_NAME",
  "ITEM_DESCRIPTION",
  "ITEM_DESCRIPTION-1",
  "QUANTITY",
  "UNIT_CODE",
  "UNIT_PRICE",
  "UNIT_PRICE_Template",
  "ITEM_AMOUNT",
  "ITEM_AMOUNT_Template",
  "ITEM_VAT",
  "ITEM_TOTAL",
  "ITEM_ALLOWANCE_CHARGE",
  "ITEM_AC_AMOUNT",
  "ITEM_REASON",
  "ITEM_REASON_CODE",
  "ITEM_ALLOWANCE_CODE",
  "ITEM_TAX_TYPE",
  "ITEM_VAT_RATE",
  "REMARKS",
  "CUST_ID",
  "OUTPUT_FILENAME",
  "PDF_PATH",
  "XML_PATH",
  "EMAIL_SUBJECT",
  "SEND_TO_RD",
  "ATTACHMENT",
  "DEPARTMENT",
  "PRINT_PAPER",
  "PAYMENT_BY",
  "PAYMENT_Date",
  "DELIVERY_BY",
  "DELIVERY_DATE",
];

// - Set of profile using for the table conversion
let SetOfProfile = [];

// ! Profile Setting Rule !
// ! Column and Row identification : {Col}{Row}
// ! Seperate each data with the ; (semicolon)
// ! Ex : C1;B2; => First Column : C1 | Second Column : B2
// % Default profile state the belonged value column from original file, if "CV" mean using the ConditionalValue, null'' mean null value
let DefaultProfile = [
  "CV", // A column
  "CV", // B column (ConditionalValue used)
  "P7", // C column
  "O8", // D column
  "CV", // E column
  "CV", // F column
  "Q79", // G to I column (repeated for all three)
  "Q79",
  "Q79",
  "Q81", // J column
  "Q79", // K column
  "Q81", // L to M column (repeated for both)
  "Q81",
  "Q83", // N to O column (repeated for both)
  "Q83",
  null, // P to T column (null for all)
  null,
  null,
  null,
  null,
  "CV", // U column
  null, // V column
  "CV", // W column
  "CV", // X column (uncertain)
  "CV", // Y column
  "CV", // Z column (derived from C5, first part)
  "CV", // AA column
  "CV", // AB column
  "CV", // AC column
  "CV", // AD column
  "CV", // AE column
  "CV", // AF column
  "CV", // AG column
  "B7", // AH column
  null, // AI column
  "D8", // AJ column
  "CV", // AK column
  "V7", // AL column
  null, // AM to AR column (null for all)
  null,
  null,
  null,
  null,
  null,
  "B9", // AS column
  "CV", // AT column
  "V9", // AU column
  "B11", // AV column
  null, // AW to BC column (null for all)
  null,
  null,
  null,
  null,
  null,
  null,
  "CV", // BD column
  "P10", // BE column
  "R10", // BF column
  "P11", // BG column
  "R9", // BH column
  "E14", // BI column
  "CV", // BJ column
  null,
  null, // BK column
  "A19", // BL column
  null, // BM column
  "E19", // BN column
  null, // BO to BP column (null for all)
  "M19", // BQ column
  null, // BR column
  "O19", // BS to BT column (repeated for both)
  "O19",
  "Q19", // BU to BV column (repeated for both)
  "Q19",
  "Q81", // BW column
  "Q83", // BX column
  null, // BY to CC column (null for all)
  null,
  null,
  null,
  null,
  "CV", // CD column
  "CV", // CE column
  null, // CF to CH column (null for all)
  null,
  null,
  "CV", // CI column
  "CV", // CJ column
  null, // CK column
  "CV", // CL column (ConditionalValue used)
  "V12", // CM column
  "CV", // CN column
  "U13", // CO column
  null, // CP to CS column (null for all)
  null,
  null,
  null,
];

// % Conditional value is used for those column that doesn't belong to any value instead it checks whether the given column
// % met the condition or not
// % Identify by {"(Target Column) column" : {(CONDITION/DEFAULT/OPERATION) : {condition from the key}, COUNT : {number}, IGNORE : {(CHECK COLUMN) : (CHECK VALUE)}}}
// % CONDITION : {("(Checking Column) (OUTPUT table/INPUT table)" : {(condition) : (value), (condition) : (value) ,... })}
// % DEFAULT : (default value)
// % OPERATION : ["Data Column", "length", "(start/end)", "(value)"] - for substring
// % OPERATION : ["Data Column", (Target Column/ Default value), "add" , "(Target Column/ Default value)"] - for joining string
// % COUNT : dedicated the how many time this value will be print out. (optional)
// % IGNORE : ignore the value of specific cell (optional)
let ConditionalValue = {
  "A column": { DEFAULT: "RECEIPT" },
  "B column": {
    CONDITION: {
      "A OUTPUT": {
        RECEIPT: "T01",
        "TAX INVOICE / INVOICE / DEBIT NOTE": "T02",
      },
    },
  },
  "CL column": {
    CONDITION: { "V12 INPUT": { null: "N", "not null": "Y" } },
  },
  "AK column": {
    CONDITION: {
      "G8 INPUT": {
        "(Head Office)": "000000",
        ELSE: "000001",
      },
    },
  },
  "E column": {
    DEFAULT: "VAT",
  },
  "F column": {
    DEFAULT: 7,
  },
  "W column": {
    OPERATION: ["L3", "length", "start", "13"],
  },
  "U column": {
    DEFAULT: "Toyota Tsusho Systems (Thailand) Co.,Ltd.",
  },
  "X column": {
    DEFAULT: "00000",
  },
  "Y column": { OPERATION: ["C4", "add", "C5"] },
  "Z column": {
    DEFAULT: "87/2",
  },
  "AA column": {
    DEFAULT: "CRC Tower, All Seasons Place",
  },
  "AB column": {
    DEFAULT: "Wireless Road",
  },
  "AC column": {
    DEFAULT: "Lumpini",
  },
  "AD column": {
    DEFAULT: "Pathumwan",
  },
  "AE column": {
    DEFAULT: "Bangkok",
  },
  "AF column": {
    DEFAULT: 10330,
  },
  "AG column": {
    DEFAULT: "eTaxInvoice.noreply.th@ttsystems.com",
  },

  "AT column": {
    OPERATION: ["B9", "length", "start", "-5"],
  },
  "BD column": {
    OPERATION: ["E15", "length", "start", "10"],
  },
  "BI column": {
    DEFAULT: "Bath",
  },
  "BJ column": {
    OPERATION: ["E71", "length", "from", "1", "-1"],
  },
  "CD column": {
    DEFAULT: "VAT",
  },
  "CE column": {
    DEFAULT: 7,
  },
  "CI column": {
    DEFAULT: "PATH TO SOMETHING",
  },
  "CJ column": {
    DEFAULT: "PATH TO SOMETHING",
  },
  "CN column": {
    DEFAULT: "IT Solution",
  },
};

// * Session
app.use(
  session({
    secret: "3xcelF1lw",
    resave: false,
    saveUninitialized: true,
  })
);
app.use(bodyParser.json());

let transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    type: `OAuth2`,
    user: env.GMAIL,
    pass: env.GMAIL_PASS,
    clientId: env.CLIENT_ID,
    clientSecret: env.CLIENT_SECRET,
    refreshToken: env.REFRESH_TOKEN,
  },
});

// > Setting View engine
app.set("view engine", "ejs");
app.use(express.static("public"));

// - GET Method
// > Default Routing
app.get("/", async (req, res) => {
  req.session.token = await Hashing();

  res.redirect("/login");
});

// > Login Page
app.get("/login", async (req, res) => {
  // > Clear the user information
  req.session.user = null;

  // > Generate token if the token is not generated
  if (!req.session.token) {
    req.session.token = await Hashing();
  }

  res.render("login");
});

// > Main menu Page
app.get("/excelconvert", (req, res) => {
  // > Redirect to the `login` page, if the user is not login yet
  /*if (!req.session.user) {
    return res.redirect("/login");
  }*/

  // > Render the main menu page
  res.render("excelConvert");
});

// > Header of the result table
app.get("/header", (req, res) => {
  // > Send the json file to fron end
  res.json({ Header: OutputTableHeader });
});

// > Profile Queries
app.get("/profile", (req, res) => {
  res.json({
    "Default Profile": {
      defaultProfile: DefaultProfile,
      conditionalValue: ConditionalValue,
    },
  });
});

// - POST Method
// > Login Page
app.post("/login", (req, res) => {
  const { gmail, password } = req.body;

  // * Check whether the password is correct or not
  if (
    req.session.token.includes(password) ||
    password.includes(req.session.token)
  ) {
    req.session.user = gmail;
    res.sendStatus(200);
  } else {
    return res.sendStatus(300);
  }
});

// > Token Checking Page
app.post("/token", async (req, res) => {
  console.log(req.body);

  const email = req.body.gmail;
  token = req.session.token;

  await sendMail(email, `${token}`, "- Token -");

  res.sendStatus(200);
});

// > Profile Saving Process
app.post("/saveProfile", (req, res) => {
  const { profile, User } = req.body;
});

// > Json proceeding process
app.post("/SendJson", (req, res) => {
  const { json } = req.body;

  if (!json) {
    console.error("No JSON data received");
    return res.status(400).send("No JSON data received");
  }

  console.log(json); // This should log the received JSON data
  res.sendStatus(200);
});

// + Running the program
// - Run the app with given port
app.listen(PORT, () => {
  // > Server URL
  const serverURL = `http://localhost:${PORT}`;

  console.log(`Server is running on port ${PORT}`);
  console.log(`Server running at: ${serverURL}`);
});

// > Generate a random string of a given length
function generateRandomString(length) {
  const characters =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let result = "";
  const charactersLength = characters.length;
  for (let i = 0; i < length; i++) {
    result += characters.charAt(Math.floor(Math.random() * charactersLength));
  }
  return result;
}

// > Get a random length between 15 and 64
function randomLength(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// > Hashing function
function Hashing() {
  const length = randomLength(15, 64);
  const randomStr = generateRandomString(length);
  const hash = sha256(randomStr);

  return hash;
}

// > Send mail
function sendMail(destination, message, subject) {
  const htmlMessage = `
    <div style="font-family: Arial, sans-serif; line-height: 1.5; height: 500px; padding: 20px;">
      <h2 style="color: #333;">${subject}</h2>
      <div style="display: flex; align-items: center;">
        <p style="margin: 0;">Your token is:</p>
        <p style="color: #ed401b; margin-left: 10px; font-weight: bold;">${message}</p>
      </div>
      <p style="color: #555; font-size: 1.4em;">
        If you did not request this code, please ignore this email.
      </p>
      <footer style="margin-top: 20px;">
        <p style="font-size: 1.4em; color: #777;">
          Regards,<br />
          Thanachart Wangphembundit, Head of Programming
        </p>
      </footer>      
    </div>
    <script>


    </script>
  `;

  const MailOption = {
    from: env.GMAIL,
    to: destination,
    subject: subject,
    html: htmlMessage,
  };

  // Sending mail using transporter
  transporter.sendMail(MailOption, (error, info) => {
    // > Error if error :)
    if (error) {
      console.error(`Error sending email : ${error}`);
    } else {
      console.log(`Successfully sending email : ${info}`);
    }
  });
}
