const express = require("express");
const bodyParser = require("body-parser");
const session = require("express-session");

<<<<<<< Updated upstream:Web page/app.js
// ! DONT FORGOT TO DELETE THE CODE WITH STATEMENT `console.log() IN PRODUCTION MODE`
// ! NEVER SHARE THE `.env` FILE IN PUBLIC
// % Port Use
const PORT = 4000;
=======
const PORT = process.env.PORT || 4500;
>>>>>>> Stashed changes:Deployment/app.js

const app = express();

<<<<<<< Updated upstream:Web page/app.js
// - Get env content
const env = dotenv.config().parsed;

// - Establish connection between database and server
const supabase = createClient(env.SUPABASE_URL, env.SUPABASE_KEY);

// - Header of the table
=======
>>>>>>> Stashed changes:Deployment/app.js
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

app.use(
  session({
    secret: "3xcelF1lw",
    resave: false,
    saveUninitialized: true,
  })
);
app.use(bodyParser.json());

app.set("view engine", "ejs");
app.use(express.static("public"));

app.get("/", async (req, res) => {
  res.redirect("/excelConvert");
});

app.get("/excelconvert", (req, res) => {
  res.render("excelConvert");
});

app.get("/header", (req, res) => {
  res.json({ Header: OutputTableHeader });
});

app.get("/profile", (req, res) => {
  res.json({
    "Default Profile": {
      defaultProfile: DefaultProfile,
      conditionalValue: ConditionalValue,
    },
  });
});

app.listen(PORT, () => {
  console.log(`Server is now  online`);
});
