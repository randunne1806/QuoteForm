const XLSX = require('xlsx');
const { MongoClient, ServerApiVersion } = require('mongodb');
const uri = "mongodb+srv://sean:gOBYvLh1jis9P3tL@cluster0.liqxi.mongodb.net/?retryWrites=true&w=majority";
const url = "mongodb+srv://seanrandunne:POxQegQ51Z2VKwpH@cluster0.on1akk8.mongodb.net/?retryWrites=true&w=majority";

const client = new MongoClient(url, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
});
const filePath = './306624 - SAM339.xlsm';
const sheetName = 'Order';
const sheetName2 = 'Quote';

// Example usage for your specific table range
const startCell = { row: 3, col: 'C' };
const endCell = { row: 10, col: 'L' };

// Example usage
const customKeyMap2 = {
    CustomerCode: 'B2',
    CustName: 'A3',
    QuoteNum: 'B4',
    PartNum: 'B5',
    Revision: 'B6',
    NumUpPerPanel: 'B7',
    NumUpPerArray: 'B8',
    XOutsAllowed: 'B9',
    Tooling: 'B10',
    Amortize: 'B11',
    ElectricalTest: 'B12',
    Amortized: 'B13',
    Stencils: 'B15',
};

const customKeyMap = {
  JobNumber: 'B3',
  PreviousRev: 'B4',
  PartNumber: 'B5',
  Revision: 'B6',
  GeneralSpec: 'B7',
  QualitySpec: 'B9',
  ITAR: 'B10',
  AspectRatio: 'B11',
  CrossSection: 'B12',
  Priority: 'E3',
  DueDate: 'E4',
  OrderQty: 'E5',
  OrderValue: 'E6',
  NumUpPanel: 'E7',
  LayerCount: 'E9',
  Thickness: 'E10',
  SurfaceFinish: 'E11',
  MaterialType: 'E12',
  IntegratorNum: 'E14',
  CustCode: 'H3',
  Company: 'H4',
  ContactName: 'H5',
  ContactPhone: 'H6',
  Email: 'H7',
  ImpdTest: 'H9',
  ElectricalTest: 'H10',
  HiPot: 'H11',
  XOuts: 'H12',
}; // Replace with your custom key-to-cell-address mapping

async function uploadToMongo(document){
    await client.connect();
    const db = client.db('amitron-labs-lake');
    const collection = db.collection("Quote-Form");
    await db.collection('Quote-Form').insertOne(document);
    console.log(document);
    console.log("Form data successfully uploaded to MongoDB");
}

// Function to extract specific cells from an Excel file and save to a dictionary with custom keys
function extractCells(filePath, sheetName, keyMap) {
  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets[sheetName];

  const cellValues = {};

  Object.entries(keyMap).forEach(([customKey, cellAddress]) => {
    const cellValue = worksheet[cellAddress]?.v; // Use optional chaining to handle undefined cells
    cellValues[customKey] = cellValue;
  });
  //console.log(cellValues);
  //uploadToMongo(cellValues);
  return cellValues;
}

function extractTable(filePath, sheetName, startCell, endCell) {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];

    const cellValues = {};

    // Convert column letter to index (e.g., 'A' => 0, 'B' => 1, etc.)
    const columnToIndex = col => col.charCodeAt(0) - 'A'.charCodeAt(0);

    // Extract data from the specified range
    for (let row = startCell.row; row <= endCell.row; row++) {
        const keyCellAddress = XLSX.utils.encode_cell({ r: row, c: columnToIndex(startCell.col) });
        const key = worksheet[keyCellAddress]?.v;

        // Skip the row if a valid key is not found
        if (!key) {
            continue;
        }

        const values = [];

        // Start from the first column after the key
        for (let col = columnToIndex(startCell.col) + 1; col <= columnToIndex(endCell.col); col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cellAddress2 = XLSX.utils.encode_cell({ r: 2, c: col });
            const cellValue = worksheet[cellAddress]?.v;
            const cellValue2 = worksheet[cellAddress2]?.v;
            //console.log(cellValue2);
            // Only add non-undefined values to the array
            if (cellValue !== undefined) {
                const entry = {};
                entry[cellValue2] = cellValue;
                //console.log(cellValue + " - row: " + row + " col: " + col)
                values.push(entry);
            }
        }

        cellValues[key] = values;
    }

    return cellValues;
}

function extractSingleCell(filePath, sheetName, cellAddress) {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];

    const cellValue = worksheet[cellAddress]?.v; // Use optional chaining to handle undefined cells

    return cellValue;
}

function createDictionary(keys, values) {
    if (keys.length !== values.length) {
        //console.log(values.length);
        throw new Error('Arrays must be of the same size');
    }

    return keys.reduce((result, key, index) => {
        result[key] = values[index];
        return result;
    }, {});
}

function extractRowValues(filePath, sheetName, section, rowNumber) {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];

    const columnKeys = [section, 'PerPiece', 'PerPanel', 'Minimum', 'FlatAdder', '%Adder', 'Tooling'];
    const startCol = 'B';
    const endCol = 'L';

    const rowData = {};
    let rows = [];

    // Convert column letter to index (e.g., 'A' => 0, 'B' => 1, etc.)
    const columnToIndex = col => col.charCodeAt(0) - 'A'.charCodeAt(0);

    // Extract data for each predefined column key
    for (let i = 0; i < 11; i++) {
        const key = columnKeys[i];
        const col = columnToIndex(startCol) + i;
        const cellAddress = XLSX.utils.encode_cell({ r: rowNumber, c: col });
        const cellValue = worksheet[cellAddress]?.v;
        //console.log(col);
        // Only add non-undefined values to the dictionary
        if (col != 2 && col != 3 && col != 4 && col != 5) {
            rows.push(cellValue);
            //console.log(col + " " + cellValue)
            rowData[key] = cellValue;
        }
        //console.log(rows);
    }
    //console.log(rows);
    return createDictionary(columnKeys, rows);
}

const extractedValueH12 = extractSingleCell(filePath, sheetName2, 'H12');

const extractedOrderValues = extractCells(filePath, sheetName, customKeyMap);

const extractedValues = extractCells(filePath, sheetName2, customKeyMap2);

const extractedTable = extractTable(filePath, sheetName2, startCell, endCell);

const finalDict = {
    OrderWksht: extractedOrderValues,
    CustomerInfo: extractedValues,
    PriceBracket: extractedTable,
    Notes: extractedValueH12,
    LayerCount: extractRowValues(filePath, sheetName2, 'Basic Properties', 19),
    Construction: extractRowValues(filePath, sheetName2, 'Basic Properties', 20),
    MaterialType: extractRowValues(filePath, sheetName2, 'Basic Properties', 21),
    Thickness: extractRowValues(filePath, sheetName2, 'Basic Properties', 22),
    SurfaceFinish: extractRowValues(filePath, sheetName2, 'Basic Properties', 23),
    GoldTab: extractRowValues(filePath, sheetName2, 'Basic Properties', 24),
    TwoOzLayers: extractRowValues(filePath, sheetName2, 'Copper Weight', 27),
    ThreeOzLayers: extractRowValues(filePath, sheetName2, 'Copper Weight', 28),
    FourOzLayers: extractRowValues(filePath, sheetName2, 'Copper Weight', 29),
    FiveOzLayers: extractRowValues(filePath, sheetName2, 'Copper Weight', 30),
    SixOzLayers: extractRowValues(filePath, sheetName2, 'Copper Weight', 31),
    UnevenCopper: extractRowValues(filePath, sheetName2, 'Copper Weight', 32),
    ExtraCopper: extractRowValues(filePath, sheetName2, 'Copper Weight', 33),
    SmallestHole: extractRowValues(filePath, sheetName2, 'Mechanical & Fabrication', 36),
    EdgeBevel: extractRowValues(filePath, sheetName2, 'Mechanical & Fabrication', 37),
    Slots: extractRowValues(filePath, sheetName2, 'Mechanical & Fabrication', 38),
    CounterSink: extractRowValues(filePath, sheetName2, 'Mechanical & Fabrication', 39),
    EdgePlating: extractRowValues(filePath, sheetName2, 'Mechanical & Fabrication', 40),
    Scoring: extractRowValues(filePath, sheetName2, 'Mechanical & Fabrication', 41),
    BreakApart: extractRowValues(filePath, sheetName2, 'Mechanical & Fabrication', 42),
    SolderMask: extractRowValues(filePath, sheetName2, 'Inks & Coatings', 45),
    Legend: extractRowValues(filePath, sheetName2, 'Inks & Coatings', 46),
    CarbonInk: extractRowValues(filePath, sheetName2, 'Inks & Coatings', 47),
    PeelableMask: extractRowValues(filePath, sheetName2, 'Inks & Coatings', 48),
    ViaFillPlug: extractRowValues(filePath, sheetName2, 'Inks & Coatings', 49),
    PlasmaDesmear: extractRowValues(filePath, sheetName2, 'Special Processes', 52),
    AluminumFinish: extractRowValues(filePath, sheetName2, 'Special Processes', 53),
    ElectricalTest: extractRowValues(filePath, sheetName2, 'Quality & Testing', 57),
    QualitySpec: extractRowValues(filePath, sheetName2, 'Quality & Testing', 58),
    TDRTesting: extractRowValues(filePath, sheetName2, 'Quality & Testing', 59),
    FirstArticle: extractRowValues(filePath, sheetName2, 'Quality & Testing', 60),
    HiPotTesting: extractRowValues(filePath, sheetName2, 'Quality & Testing', 61),
    ITAR: extractRowValues(filePath, sheetName2, 'Quality & Testing', 62)
}

uploadToMongo(finalDict);
