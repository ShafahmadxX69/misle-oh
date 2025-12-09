import ExcelJS from 'exceljs';
import { FormData } from '../types';
import saveAs from 'file-saver';

// Helper to create thin borders
const borderStyle: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

// Helper for center alignment
const alignCenter: Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'center', wrapText: true };
const alignLeft: Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'left', wrapText: true };

// Helper for yellow background
const yellowFill: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFF00' } // Yellow
};

const whiteFill: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFFFF' } // White
};

export const generateExcel = async (data: FormData) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Stuffing List');

  // Define Columns (Approximate widths based on image)
  worksheet.columns = [
    { key: 'A', width: 2 },  // Padding column or margin
    { key: 'B', width: 25 }, // Labels / Material No
    { key: 'C', width: 25 }, // Name Spec
    { key: 'D', width: 10 }, // Qty
    { key: 'E', width: 10 }, // Total
    { key: 'F', width: 40 }, // Description
    { key: 'G', width: 15 }, // Customer PO
    { key: 'H', width: 15 }, // ULI PO
    { key: 'I', width: 15 }, // Brand
  ];

  // --- Header Section ---
  
  // Row 1 & 2: Logo and Title
  // Logo (Merged B1:D2)
  worksheet.mergeCells('B1:D2');
  const logoCell = worksheet.getCell('B1');
  logoCell.value = 'PT. UNIVERSAL LUGGAGE INDONESIA'; // Placeholder for logo text
  logoCell.style = { 
    font: { bold: true, size: 14, color: { argb: 'FF000000' } }, // Black color
    alignment: alignCenter,
    border: borderStyle
  };

  // Title (Merged E1:I2)
  worksheet.mergeCells('E1:I2');
  const titleCell = worksheet.getCell('E1');
  titleCell.value = 'FORM STUFFING LIST\n装  箱  單';
  titleCell.style = { 
    font: { bold: true, size: 16 }, 
    alignment: alignCenter,
    border: borderStyle
  };

  // --- Info Section (Rows 3-11) ---

  const addInfoRow = (rowNum: number, label: string, value: string, mergeValue: boolean = true) => {
    const labelCell = worksheet.getCell(`B${rowNum}`);
    labelCell.value = label;
    labelCell.border = borderStyle;
    labelCell.alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };

    if (mergeValue) {
        worksheet.mergeCells(`C${rowNum}:D${rowNum}`);
    }
    
    const valueCell = worksheet.getCell(`C${rowNum}`);
    valueCell.value = value;
    valueCell.border = borderStyle;
    valueCell.fill = value ? yellowFill : whiteFill;
    valueCell.alignment = alignCenter;
    valueCell.font = { bold: true };
  };

  addInfoRow(3, '發票流水號 (INV FLOW No)', data.invFlowNo);
  addInfoRow(4, '訂單號碼 (PO.NO)', data.poNo);
  addInfoRow(5, '客戶 (Customer)', data.customer);
  addInfoRow(6, '出貨日期 (Shipping date)', data.shippingDate);
  addInfoRow(7, '船名 (Vessel name)', data.vesselName);
  addInfoRow(8, '櫃數 (Container Qty)', data.containerQty);
  addInfoRow(9, '货 柜 号(Container NO)', data.containerNo);
  addInfoRow(10, '出货单编号 (Delivery note No)', data.deliveryNoteNo);
  addInfoRow(11, '嘜頭 (Mark)', data.mark);

  // REMARK Block
  worksheet.mergeCells('E3:I3');
  const remarkLabel = worksheet.getCell('E3');
  remarkLabel.value = '備註說明\nREMARK';
  remarkLabel.alignment = alignCenter;
  remarkLabel.border = borderStyle;

  worksheet.mergeCells('E4:I4'); 
  worksheet.getCell('E4').border = borderStyle;

  worksheet.mergeCells('E5:I11');
  const remarkValue = worksheet.getCell('E5');
  remarkValue.value = data.remark;
  remarkValue.fill = yellowFill;
  remarkValue.alignment = alignCenter;
  remarkValue.font = { bold: true, size: 24 };
  remarkValue.border = borderStyle;

  // --- Table Headers (Rows 12-13) ---
  
  // Row 12
  worksheet.getRow(12).height = 20;
  worksheet.getCell('B12').value = '料  号';
  worksheet.getCell('C12').value = '品名/规格';
  worksheet.getCell('D12').value = '每箱数量';
  worksheet.getCell('E12').value = '箱数合计';
  worksheet.getCell('F12').value = '每箱包含的要点及颜色';
  worksheet.getCell('G12').value = '客戶PO';
  worksheet.getCell('H12').value = '工廠';
  worksheet.getCell('I12').value = '品牌';

  // Row 13
  worksheet.getRow(13).height = 20;
  worksheet.getCell('B13').value = 'Material No';
  worksheet.getCell('C13').value = '(Name and spec)';
  worksheet.getCell('D13').value = 'PCS/CTN';
  worksheet.getCell('E13').value = 'Total Ctn Qty';
  worksheet.getCell('F13').value = 'main point &color for each ctn\n(务必要写清楚 be detailed)';
  worksheet.getCell('G13').value = 'Customer PO';
  worksheet.getCell('H13').value = 'ULI PO';
  worksheet.getCell('I13').value = 'Brand';

  // Merging Headers
  worksheet.mergeCells('B12:B13');
  worksheet.mergeCells('C12:C13');
  worksheet.mergeCells('F12:F13');
  worksheet.mergeCells('G12:G13');
  worksheet.mergeCells('H12:H13');
  worksheet.mergeCells('I12:I13');

  // Styling Headers
  ['B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12', 'I12', 'D13', 'E13'].forEach(cell => {
    const c = worksheet.getCell(cell);
    c.alignment = alignCenter;
    c.border = borderStyle;
    c.font = { bold: true, size: 9 };
  });

  // --- Table Body (Rows 14+) ---

  // Row 14: Special Container Info
  worksheet.getCell('B14').value = data.containerSizeInfo;
  worksheet.getCell('B14').fill = yellowFill;
  worksheet.getCell('B14').alignment = alignCenter;
  worksheet.getCell('B14').border = borderStyle;
  
  ['C14', 'D14', 'E14', 'F14', 'G14', 'H14', 'I14'].forEach(cell => {
      worksheet.getCell(cell).border = borderStyle;
  });

  // Data Rows (15 - 21)
  let currentRow = 15;
  data.items.forEach((item) => {
    const r = currentRow;
    
    // Set Values
    worksheet.getCell(`B${r}`).value = item.materialNo;
    worksheet.getCell(`C${r}`).value = item.nameAndSpec;
    worksheet.getCell(`D${r}`).value = item.pcsPerCtn;
    worksheet.getCell(`E${r}`).value = item.totalCtnQty; // Ensure number format if possible
    worksheet.getCell(`F${r}`).value = item.description;
    worksheet.getCell(`G${r}`).value = item.customerPo;
    worksheet.getCell(`H${r}`).value = item.uliPo;
    worksheet.getCell(`I${r}`).value = item.brand;

    // Styles
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${r}`);
      cell.border = borderStyle;
      cell.alignment = alignCenter;
      cell.fill = yellowFill; // Data rows are yellow
    });

    currentRow++;
  });

  // Fill up to row 22 if data is short, to match template look
  while (currentRow <= 22) {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${currentRow}`);
      cell.border = borderStyle;
      cell.alignment = alignCenter;
    });
    currentRow++;
  }

  // --- Footer ---
  const footerRowStart = 23;
  
  // Total Row
  worksheet.getCell(`B${footerRowStart}`).value = '合   計 TOTAL';
  worksheet.getCell(`B${footerRowStart}`).alignment = alignCenter;
  worksheet.getCell(`B${footerRowStart}`).border = borderStyle;
  worksheet.mergeCells(`B${footerRowStart}:D${footerRowStart}`);
  
  // Calculate Total
  const totalQty = data.items.reduce((sum, item) => sum + (parseInt(item.totalCtnQty) || 0), 0);
  worksheet.getCell(`E${footerRowStart}`).value = totalQty;
  worksheet.getCell(`E${footerRowStart}`).fill = yellowFill;
  worksheet.getCell(`E${footerRowStart}`).border = borderStyle;
  worksheet.getCell(`E${footerRowStart}`).alignment = alignCenter;
  
  // Borders for rest of total row
  ['F', 'G', 'H', 'I'].forEach(col => {
      worksheet.getCell(`${col}${footerRowStart}`).border = borderStyle;
  });

  // Signatures Row 24
  const sigRow = 24;
  worksheet.mergeCells(`B${sigRow}:D${sigRow}`);
  worksheet.getCell(`B${sigRow}`).value = '生管主管 (Production control)：';
  
  worksheet.mergeCells(`E${sigRow}:F${sigRow}`);
  worksheet.getCell(`E${sigRow}`).value = '生管填表 (production fill in)：';
  
  worksheet.mergeCells(`G${sigRow}:I${sigRow}`);
  worksheet.getCell(`G${sigRow}`).value = '業務確認(Business Unit)：';

  // Signatures Row 25
  const sigRow2 = 25;
  worksheet.mergeCells(`B${sigRow2}:D${sigRow2}`);
  worksheet.getCell(`B${sigRow2}`).value = '資材主管 (Warehouse manage)：';

  worksheet.mergeCells(`E${sigRow2}:F${sigRow2}`);
  worksheet.getCell(`E${sigRow2}`).value = '成品倉 (finished goods warehouse)：';

  worksheet.mergeCells(`G${sigRow2}:I${sigRow2}`);
  worksheet.getCell(`G${sigRow2}`).value = '貨櫃檢驗確認 (Container examine)：';

  // Document Info Row 26
  const docRow = 26;
  worksheet.getCell(`C${docRow}`).value = 'Usia Penyimpanan : 1 tahun (保存年限：一年)';
  
  worksheet.mergeCells(`G${docRow}:H${docRow}`);
  worksheet.getCell(`G${docRow}`).value = 'Dok No : Form - PPIC - 03';

  // Generate Buffer
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, 'Stuffing_List.xlsx');
};