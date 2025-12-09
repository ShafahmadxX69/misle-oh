import ExcelJS from 'exceljs';
import { FormData, OrderItem } from '../types';

// Helper to read an uploaded file into an ExcelJS Workbook
const readWorkbook = async (file: File): Promise<ExcelJS.Workbook> => {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
};

// --- CUFT DICTIONARY (Ported from VBA) ---
const getCUFT = (model: string, size: string): number | null => {
  const cuftDict: Record<string, Record<string, number>> = {
    "FQ832": { "21": 2.08, "22": 2.35, "26": 3.52, "29": 4.85 },
    "FQ825": { "21": 2.08, "22": 2.35, "26": 3.52, "29": 4.85 },
    "FR885": { "21": 2.08, "22": 2.35, "26": 3.52, "29": 4.85 },
    "F1627": { "21": 1.95, "24": 2.55, "26": 3.36, "29": 4.59, "29.5": 4.65, "30": 5.45, "19": 1.818, "20": 1.818, "27": 4.391, "28": 4.391 },
    "F1628": { "21": 1.95, "24": 2.55, "26": 3.36, "29": 4.59, "29.5": 4.65, "30": 5.45, "19": 1.818, "20": 1.818, "27": 4.391, "28": 4.391 },
    "FR893": { "21": 2.15, "22": 2.37, "25": 3.78, "28": 5.5, "32": 5.79 },
    "PP8": { "S": 1.95, "M": 3.88, "L": 5.89 },
    "PP10": { "S": 1.94, "M": 3.9, "L": 5.85 },
    "PP12": { "S": 1.88, "M": 3.21, "L": 5.1 },
    "FJ616": { "16": 1.55, "19.5": 1.92, "20": 2.01, "21": 2.29, "24": 3.3, "29": 4.37, "31": 3.98, "32": 6.16 },
    "FJ616-1": { "16": 1.55, "19.5": 1.92, "20": 2.01, "21": 2.29, "24": 3.3, "29": 4.37, "31": 3.98, "32": 6.16 },
    "FK648": { "20": 1.93, "24": 3.11, "29": 4.16 },
    "FK636": { "16": 1.5, "20": 2.01, "21": 2.11, "24": 3.71, "29": 4.38 },
    "FK636-1": { "16": 1.5, "20": 2.01, "21": 2.11, "24": 3.71, "29": 4.38 },
    "FL688": { "16": 1.42, "17": 1.59, "19": 2.47, "20": 2.06, "21": 2.24, "24": 3.29, "29": 4.7, "31": 4.18, "32": 6.38, "29.5": 4.63 },
    "FL688-1": { "16": 1.42, "17": 1.59, "19": 2.47, "20": 2.06, "21": 2.24, "24": 3.29, "29": 4.7, "31": 4.18, "32": 6.38, "29.5": 4.63 },
    "FL688-6": { "16": 1.42, "17": 1.59, "19": 2.47, "20": 2.06, "21": 2.24, "24": 3.29, "29": 4.7, "31": 4.18, "32": 6.38, "29.5": 4.63 },
    "FH496": { "21": 2.09, "22": 2.2, "27": 3.91, "29": 4.83, "30": 4.83, "31": 5.68, "32": 5.75 },
    "FR898": { "20": 2.47, "21": 2.12, "24": 3.53, "29": 4.7 },
    "F1909": { "21.5": 2.17, "22": 2.38, "24": 3.3, "25": 3.79, "28": 5.5, "30": 5.78, "32": 5.8, "29": 4.95 },
    "FG417": { "20": 2.08, "27": 3.98, "29": 4.71, "32": 5.75 },
    "FQ819-1": { "19": 1.48, "21": 2.38, "26": 3.62, "29": 5.035 },
    "FJ587-1": { "23": 2.22, "32": 4.8 },
    "FBP01/": { "S#": 2.09, "M#": 4.2, "L#": 5.5, "24#": 2.78 },
    "PFBP01": { "S#": 2.09, "M#": 4.2, "L#": 5.5, "24#": 2.78 },
    "FL678": { "19.5": 1.64, "20": 1.74, "21": 1.73, "25": 3.14, "28": 4.2 },
    "FQ822": { "19.5": 1.74, "20": 1.78, "21": 2.14, "25": 3.43, "28": 4.55, "31": 6.2 },
    "FP763-1": { "20": 2.16, "29": 4.41 }
  };

  const cleanModel = model.toUpperCase();
  const cleanSize = size.toUpperCase();

  // Find matching key
  const dictKey = Object.keys(cuftDict).find(k => cleanModel.includes(k.toUpperCase()));
  if (dictKey && cuftDict[dictKey][cleanSize]) {
    return cuftDict[dictKey][cleanSize];
  }
  return null;
};

// --- PROCESS LOGIC ---

export const processAutomation = async (
  currentData: FormData,
  siFile: File | null,
  indexFile: File | null
): Promise<FormData> => {
  let newData = { ...currentData };
  let siWorkbook: ExcelJS.Workbook | null = null;
  let indexWorkbook: ExcelJS.Workbook | null = null;

  // 1. IMPORT SI DATA (Corresponds to ImportSIDataToSheet1)
  if (siFile) {
    try {
      siWorkbook = await readWorkbook(siFile);
      const wsSI = siWorkbook.getWorksheet("SI ") || siWorkbook.worksheets[0]; // Fallback to first sheet if "SI " not found

      if (wsSI) {
        let brandVal = "";
        let destVal = "";

        // Extract Header Info
        wsSI.eachRow((row, rowNumber) => {
          const colC = row.getCell(3).value?.toString().trim() || "";
          const colI = row.getCell(9).value?.toString().trim() || "";
          
          if (colC === "CONTAINER") {
            newData.containerQty = row.getCell(5).value?.toString() || ""; // E
          }
          if (colI === "INVOICE") {
            newData.invFlowNo = row.getCell(10).value?.toString() || ""; // J
          }
          if (colC === "SHIPPING MARK") {
            brandVal = row.getCell(5).value?.toString() || ""; // E
          }
          if (colI === "SHIP TO") {
            destVal = row.getCell(10).value?.toString() || ""; // J
          }
        });

        if (brandVal || destVal) {
          newData.customer = `${brandVal} TO ${destVal}`;
        }

        // Extract Items
        const newItems: OrderItem[] = [];
        let headerRow = 0;
        
        // Find Header Row
        wsSI.eachRow((row, rowNumber) => {
          if (row.getCell(5).value === "Item") {
             headerRow = rowNumber;
          }
        });

        if (headerRow > 0) {
          // Map Headers
          const headerValues: Record<string, number> = {};
          const row = wsSI.getRow(headerRow);
          row.eachCell((cell, colNumber) => {
             headerValues[cell.value?.toString() || ""] = colNumber;
          });

          // Read Data
          let currentRow = headerRow + 1;
          while (wsSI.getCell(currentRow, headerValues["Item"]).value) {
            const itemVal = wsSI.getCell(currentRow, headerValues["Item"]).value?.toString() || "";
            const qtyCtn = wsSI.getCell(currentRow, headerValues["QTY CTN"]).value?.toString() || "";
            const color = wsSI.getCell(currentRow, headerValues["COLOR"]).value?.toString() || "";
            const po = wsSI.getCell(currentRow, headerValues["PO"]).value?.toString() || "";
            const orderNo = wsSI.getCell(currentRow, headerValues["ORDER NO"]).value?.toString() || "";
            
            let sku = "";
            if (headerValues["SKU"]) sku = wsSI.getCell(currentRow, headerValues["SKU"]).value?.toString() || "";
            if (!sku && headerValues["PRODUCT_VARIANT"]) sku = wsSI.getCell(currentRow, headerValues["PRODUCT_VARIANT"]).value?.toString() || "";

            newItems.push({
              id: Math.random().toString(36).substr(2, 9),
              materialNo: "", // Filled later
              nameAndSpec: itemVal,
              pcsPerCtn: "1 PCS",
              totalCtnQty: qtyCtn,
              description: color, // Temporary, will be replaced by ColorCode macro
              customerPo: po,
              uliPo: orderNo,
              brand: brandVal,
              sku: sku.trim()
            });

            currentRow++;
          }
          newData.items = newItems;
        }
      }
    } catch (e) {
      console.error("Error processing SI File", e);
      alert("Error processing SI File. Check format.");
    }
  }

  // 2. PROCESS INDEX LOGIC (ColorCode & MaterialNo)
  if (indexFile && newData.items.length > 0) {
    try {
      indexWorkbook = await readWorkbook(indexFile);
      const wsIndex = indexWorkbook.getWorksheet("Sheet1") || indexWorkbook.worksheets[0];

      if (wsIndex) {
        // Cache index data to avoid slow repeated lookups
        const indexRows: any[] = [];
        wsIndex.eachRow((row, rowNumber) => {
            if (rowNumber < 2) return; // Skip header
            indexRows.push({
                materialNo: row.getCell(2).value?.toString().trim() || "", // B
                model: row.getCell(3).value?.toString().trim().replace(/['"]/g, "") || "", // C
                colorRaw: row.getCell(6).value?.toString().trim() || "", // F (e.g. "BLACK 1067...")
                so: row.getCell(8).value?.toString().trim() || "", // H
                colorMandarin: row.getCell(15).value?.toString().trim() || "", // O
                colorP: row.getCell(16).value?.toString().trim().toLowerCase() || "" // P
            });
        });

        // Loop through current items to apply logic
        newData.items = newData.items.map(item => {
            const soTemplate = item.uliPo.trim();
            let modelTemplate = item.nameAndSpec.trim().replace(/['"]/g, "");
            
            // Logic from VBA: If strict structure, parse model
            // Parsing model name logic is tricky, simplifying to basic trim for now, 
            // but mimicking: split space, take last part if multiple parts exist
            if (modelTemplate.includes(" ")) {
                const parts = modelTemplate.split(" ");
                modelTemplate = parts[parts.length - 1];
            }

            const colorCodeTemplate = item.description.trim().toLowerCase(); // Currently holds the raw color from SI
            const skuCode = item.sku || "";

            // Find Match
            const match = indexRows.find(idx => {
                let idxModel = idx.model;
                if (idxModel.includes(" ")) {
                     const parts = idxModel.split(" ");
                     idxModel = parts[parts.length - 1];
                }
                
                // VBA ColorCode Logic: Check if index colorP is inside the template color description
                // The VBA: If Replace(colorPIndex, " ", "") = Replace(colorCodeTemplate, " ", "")
                // But in `ImportSI`, description is set to `wsSI...("COLOR")`.
                
                // Let's try flexible matching for color
                const colorMatch = idx.colorP && colorCodeTemplate.replace(/\s/g, "").includes(idx.colorP.replace(/\s/g, ""));
                
                return idx.so === soTemplate && idxModel === modelTemplate && colorMatch;
            });

            if (match) {
                // Apply ColorCode Logic
                const newDesc = skuCode ? `${match.colorMandarin} ${skuCode}` : match.colorMandarin;
                
                // Apply MaterialNo Logic (simplified: usually matches same row)
                const newMaterial = match.materialNo;

                return {
                    ...item,
                    description: newDesc,
                    materialNo: newMaterial
                };
            }
            
            // Second pass for MaterialNo specifically if exact color code match fails but SO/Model matches?
            // VBA does separate loops, but generally they look for the same row.
            // If no match found, keep original
            return item;
        });
      }
    } catch (e) {
      console.error("Error processing Index File", e);
      alert("Error processing Index File.");
    }
  }

  // 3. AGGREGATE SO (VBA Sub SO)
  const uniqueSOs = new Set<string>();
  newData.items.forEach(item => {
      if (item.uliPo) uniqueSOs.add(item.uliPo);
  });
  // Note: App maps poNo to "Order No / PO NO", usually ULI PO goes to H column, but the aggregated one goes to Header.
  // VBA: ws.Range("C4").Value = soString (C4 is PO NO in Excel Template)
  // App: poNo maps to C4.
  if (uniqueSOs.size > 0) {
      newData.poNo = Array.from(uniqueSOs).join("/");
  }

  // 4. AGGREGATE PO (VBA Sub PO) -> Customer PO
  // VBA: ws.Range("E5").Value = poString (E5 is Remark in Excel Template)
  // App: remark maps to E5 block.
  // Logic: Aggregate Customer PO (Col G in items), ignore "Dok No"
  const uniquePOs = new Set<string>();
  newData.items.forEach(item => {
      const val = item.customerPo?.trim();
      if (val && !val.toLowerCase().includes("dok no")) {
          uniquePOs.add(val);
      }
  });
  if (uniquePOs.size > 0) {
      newData.remark = Array.from(uniquePOs).join("/");
  }

  return newData;
};

// --- CUFT CALCULATION ---
export const calculateCUFTReport = (data: FormData): string => {
    let totalCUFT = 0;
    
    // Calculate per item
    data.items.forEach(item => {
        let modelSize = item.nameAndSpec.replace(/['"]/g, "").trim();
        let sizePart = "";

        // Extract Size
        if (modelSize.includes("/")) {
            const parts = modelSize.split("/");
            sizePart = parts[parts.length - 1];
        } else if (modelSize.includes("-")) {
            const parts = modelSize.split("-");
            sizePart = parts[parts.length - 1];
        }

        sizePart = sizePart.trim().toUpperCase();

        // Identify Model from Dict Keys
        // This is a simplified check, ideally needs the full dict keys from the VBA
        // Re-using the getCUFT helper
        const cuftVal = getCUFT(modelSize, sizePart);
        
        if (cuftVal) {
             const qty = parseFloat(item.totalCtnQty) || 0;
             totalCUFT += (cuftVal * qty);
        }
    });

    return `Total Calculated CUFT: ${totalCUFT.toFixed(2)}`;
};
