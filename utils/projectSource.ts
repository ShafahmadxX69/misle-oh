import JSZip from 'jszip';
import saveAs from 'file-saver';

// --- File Content Templates ---

const PACKAGE_JSON = `{
  "name": "luggage-form-generator",
  "private": true,
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "tsc && vite build",
    "lint": "eslint . --ext ts,tsx --report-unused-disable-directives --max-warnings 0",
    "preview": "vite preview"
  },
  "dependencies": {
    "exceljs": "^4.4.0",
    "file-saver": "^2.0.5",
    "lucide-react": "^0.344.0",
    "react": "^18.2.0",
    "react-dom": "^18.2.0"
  },
  "devDependencies": {
    "@types/file-saver": "^2.0.7",
    "@types/react": "^18.2.64",
    "@types/react-dom": "^18.2.21",
    "@vitejs/plugin-react": "^4.2.1",
    "autoprefixer": "^10.4.18",
    "postcss": "^8.4.35",
    "tailwindcss": "^3.4.1",
    "typescript": "^5.2.2",
    "vite": "^5.1.4"
  }
}`;

const TSCONFIG_JSON = `{
  "compilerOptions": {
    "target": "ES2020",
    "useDefineForClassFields": true,
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "module": "ESNext",
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "jsx": "react-jsx",
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true
  },
  "include": ["src"],
  "references": [{ "path": "./tsconfig.node.json" }]
}`;

const TSCONFIG_NODE_JSON = `{
  "compilerOptions": {
    "composite": true,
    "skipLibCheck": true,
    "module": "ESNext",
    "moduleResolution": "bundler",
    "allowSyntheticDefaultImports": true
  },
  "include": ["vite.config.ts"]
}`;

const VITE_CONFIG_TS = `import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
})`;

const TAILWIND_CONFIG_JS = `/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {},
  },
  plugins: [],
}`;

const POSTCSS_CONFIG_JS = `export default {
  plugins: {
    tailwindcss: {},
    autoprefixer: {},
  },
}`;

const INDEX_HTML = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Universal Luggage Form Generator</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;700&display=swap" rel="stylesheet">
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>`;

const SRC_INDEX_CSS = `@tailwind base;
@tailwind components;
@tailwind utilities;

@layer base {
  body {
    font-family: 'Inter', 'Noto Sans TC', sans-serif;
    @apply text-black;
  }
  input, textarea {
    @apply outline-none bg-transparent text-black;
  }
}

.excel-table td, .excel-table th {
    border: 1px solid #000;
}
.excel-table {
    border: 2px solid #000;
}

@media print {
    body {
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }
    .no-print {
        display: none;
    }
}`;

const SRC_MAIN_TSX = `import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.tsx'
import './index.css'

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)`;

// Note: These files are hardcoded versions of what is currently in the app.
// Ideally, this would be read from fs, but in this env we must reconstruct.

const SRC_TYPES_TS = `export interface OrderItem {
  id: string;
  materialNo: string;
  nameAndSpec: string;
  pcsPerCtn: string;
  totalCtnQty: string;
  description: string;
  customerPo: string;
  uliPo: string;
  brand: string;
}

export interface FormData {
  invFlowNo: string;
  poNo: string;
  customer: string;
  shippingDate: string;
  vesselName: string;
  containerQty: string;
  containerNo: string;
  deliveryNoteNo: string;
  mark: string;
  remark: string;
  items: OrderItem[];
  containerSizeInfo: string;
}

export const INITIAL_DATA: FormData = {
  invFlowNo: 'INV-E2500619',
  poNo: 'YOE-25090040/YOE-25090041',
  customer: 'TIMBUK2 TO USA',
  shippingDate: '',
  vesselName: '',
  containerQty: '5*40HQ+1*20GP',
  containerNo: '',
  deliveryNoteNo: '',
  mark: '',
  remark: '410185',
  containerSizeInfo: '1.1*40HQ(2300")',
  items: [
    {
      id: '1',
      materialNo: 'CFR873021U47US11',
      nameAndSpec: 'FR873/21"',
      pcsPerCtn: '1 PCS',
      totalCtnQty: '156',
      description: 'U47#暗橄榄 Dark Olive/Moss 1067-70-1268',
      customerPo: '410185',
      uliPo: 'YOE-25090040',
      brand: 'TIMBUK2'
    },
    {
      id: '2',
      materialNo: 'CFR873021139US11',
      nameAndSpec: 'FR873/21"',
      pcsPerCtn: '1 PCS',
      totalCtnQty: '782',
      description: '139#黑色 Black 1067-70-1310',
      customerPo: '410185',
      uliPo: 'YOE-25090040',
      brand: 'TIMBUK2'
    },
    {
      id: '3',
      materialNo: 'CFR873021U45US11',
      nameAndSpec: 'FR873/21"',
      pcsPerCtn: '1 PCS',
      totalCtnQty: '122',
      description: 'U45#芒果黄 Mango/Marigold 1067-70-1312',
      customerPo: '410185',
      uliPo: 'YOE-25090040',
      brand: 'TIMBUK2'
    },
    {
      id: '4',
      materialNo: '',
      nameAndSpec: '',
      pcsPerCtn: '',
      totalCtnQty: '',
      description: '',
      customerPo: '',
      uliPo: '',
      brand: ''
    },
    {
      id: '5',
      materialNo: '',
      nameAndSpec: '',
      pcsPerCtn: '',
      totalCtnQty: '',
      description: '',
      customerPo: '',
      uliPo: '',
      brand: ''
    },
    {
      id: '6',
      materialNo: '',
      nameAndSpec: '',
      pcsPerCtn: '',
      totalCtnQty: '',
      description: '',
      customerPo: '',
      uliPo: '',
      brand: ''
    }
  ]
};`;

const SRC_COMPONENTS_FORMTABLE_TSX = `import React from 'react';
import { FormData, OrderItem } from '../types';

interface FormTableProps {
  data: FormData;
  onChange: (field: keyof FormData, value: any) => void;
  onItemChange: (index: number, field: keyof OrderItem, value: string) => void;
  onAddRow: () => void;
}

const InputCell: React.FC<{
  value: string;
  onChange: (val: string) => void;
  className?: string;
  placeholder?: string;
}> = ({ value, onChange, className = '', placeholder }) => (
  <textarea
    className={\`w-full h-full bg-yellow-200 p-1 resize-none text-center flex items-center justify-center border-none focus:ring-2 focus:ring-blue-500 text-sm font-medium text-black \${className}\`}
    value={value}
    onChange={(e) => onChange(e.target.value)}
    placeholder={placeholder}
    rows={1}
    style={{ minHeight: '1.5rem' }}
  />
);

const FormTable: React.FC<FormTableProps> = ({ data, onChange, onItemChange, onAddRow }) => {
  const calculateTotal = () => {
    return data.items.reduce((sum, item) => sum + (parseInt(item.totalCtnQty) || 0), 0);
  };

  return (
    <div className="overflow-x-auto p-4 bg-gray-100 min-h-screen">
      <div className="min-w-[1000px] max-w-[1200px] mx-auto bg-white shadow-xl p-8">
        
        {/* The Main Excel-like Table */}
        <table className="w-full border-collapse border border-black excel-table text-black">
          <colgroup>
            <col className="w-[120px]" /> {/* B - Labels */}
            <col className="w-[120px]" /> {/* C - Inputs */}
            <col className="w-[80px]" />  {/* D - Inputs */}
            <col className="w-[80px]" />  {/* E - Inputs */}
            <col className="w-[200px]" /> {/* F - Desc */}
            <col className="w-[100px]" /> {/* G - PO */}
            <col className="w-[100px]" /> {/* H - PO */}
            <col className="w-[100px]" /> {/* I - Brand */}
          </colgroup>

          <tbody>
            {/* Row 1 & 2: Header Block */}
            <tr className="h-16">
              <td colSpan={3} className="border border-black p-2 text-center align-middle">
                {/* Logo Area */}
                <div className="flex flex-col items-center justify-center h-full">
                  <div className="text-black font-bold text-xl">PT. UNIVERSAL LUGGAGE</div>
                  <div className="text-black font-bold text-lg">INDONESIA</div>
                  {/* Placeholder for actual logo image */}
                  <div className="text-xs text-gray-400 mt-1">[LOGO B1:D2]</div>
                </div>
              </td>
              <td colSpan={5} className="border border-black p-2 text-center align-middle bg-white">
                <h1 className="text-2xl font-bold">FORM STUFFING LIST</h1>
                <h2 className="text-xl font-bold tracking-widest">裝 箱 單</h2>
              </td>
            </tr>

            {/* Row 3 - 11: Header Information & Remarks */}
            
            {/* Row 3 - Start of Info & Start of Remark Header */}
            <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                發票流水號 (INV FLOW No)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.invFlowNo} onChange={(v) => onChange('invFlowNo', v)} />
              </td>
              {/* Remark Header */}
              <td colSpan={5} className="border border-black text-center text-xs font-bold align-middle bg-white">
                備註說明 <br/> REMARK
              </td>
            </tr>

            {/* Row 4 */}
            <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                訂單號碼 (PO.NO)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.poNo} onChange={(v) => onChange('poNo', v)} />
              </td>
              {/* Remark Body Starts - Merged down to Row 11 */}
              <td rowSpan={8} colSpan={5} className="border border-black p-0 align-middle">
                <textarea 
                   className="w-full h-full bg-yellow-200 text-center font-bold text-4xl p-4 resize-none focus:outline-none text-black"
                   value={data.remark}
                   onChange={(e) => onChange('remark', e.target.value)}
                />
              </td>
            </tr>

             {/* Row 5 */}
             <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                客戶 (Customer)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.customer} onChange={(v) => onChange('customer', v)} />
              </td>
            </tr>

            {/* Row 6 */}
             <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                出貨日期 (Shipping date)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.shippingDate} onChange={(v) => onChange('shippingDate', v)} />
              </td>
            </tr>

            {/* Row 7 */}
             <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                船名 (Vessel name)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.vesselName} onChange={(v) => onChange('vesselName', v)} />
              </td>
            </tr>

            {/* Row 8 */}
             <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                櫃數 (Container Qty)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.containerQty} onChange={(v) => onChange('containerQty', v)} />
              </td>
            </tr>

            {/* Row 9 */}
             <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                貨櫃號 (Container NO)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.containerNo} onChange={(v) => onChange('containerNo', v)} />
              </td>
            </tr>

             {/* Row 10 */}
             <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                出貨單編號 (Delivery note No)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.deliveryNoteNo} onChange={(v) => onChange('deliveryNoteNo', v)} />
              </td>
            </tr>

            {/* Row 11 */}
             <tr className="h-8">
              <td className="border border-black px-2 text-right text-xs font-bold bg-white">
                嘜頭 (Mark)
              </td>
              <td colSpan={2} className="border border-black p-0">
                <InputCell value={data.mark} onChange={(v) => onChange('mark', v)} />
              </td>
            </tr>

            {/* Table Headers (Row 12 & 13) */}
            <tr className="h-8 text-center text-xs font-bold">
              <td className="border border-black">料  號</td>
              <td className="border border-black">品名/規格</td>
              <td className="border border-black">每箱數量</td>
              <td className="border border-black">箱數合計</td>
              <td className="border border-black">每箱包含的要點及顏色</td>
              <td className="border border-black">客戶PO</td>
              <td className="border border-black">工廠</td>
              <td className="border border-black">品牌</td>
            </tr>
             <tr className="h-8 text-center text-xs font-bold">
              <td className="border border-black">Material No</td>
              <td className="border border-black">(Name and spec)</td>
              <td className="border border-black">PCS/CTN</td>
              <td className="border border-black">Total Ctn Qty</td>
              <td className="border border-black">main point &color for each ctn <br/> (務必要寫清楚 be detailed)</td>
              <td className="border border-black">Customer PO</td>
              <td className="border border-black">ULI PO</td>
              <td className="border border-black">Brand</td>
            </tr>

            {/* Data Rows */}
            {/* Row 14: Special Container Info Row */}
            <tr className="h-10">
               <td className="border border-black p-0">
                  <InputCell value={data.containerSizeInfo} onChange={(v) => onChange('containerSizeInfo', v)} />
               </td>
               <td className="border border-black bg-gray-50"></td>
               <td className="border border-black bg-gray-50"></td>
               <td className="border border-black bg-gray-50"></td>
               <td className="border border-black bg-gray-50"></td>
               <td className="border border-black bg-gray-50"></td>
               <td className="border border-black bg-gray-50"></td>
               <td className="border border-black bg-gray-50"></td>
            </tr>

            {/* Dynamic Items */}
            {data.items.map((item, index) => (
              <tr key={item.id} className="h-10">
                <td className="border border-black p-0">
                  <InputCell value={item.materialNo} onChange={(v) => onItemChange(index, 'materialNo', v)} />
                </td>
                <td className="border border-black p-0">
                   <InputCell value={item.nameAndSpec} onChange={(v) => onItemChange(index, 'nameAndSpec', v)} />
                </td>
                <td className="border border-black p-0">
                   <InputCell value={item.pcsPerCtn} onChange={(v) => onItemChange(index, 'pcsPerCtn', v)} />
                </td>
                <td className="border border-black p-0">
                   <InputCell value={item.totalCtnQty} onChange={(v) => onItemChange(index, 'totalCtnQty', v)} />
                </td>
                <td className="border border-black p-0">
                   <InputCell value={item.description} onChange={(v) => onItemChange(index, 'description', v)} className="text-left px-2" />
                </td>
                <td className="border border-black p-0">
                   <InputCell value={item.customerPo} onChange={(v) => onItemChange(index, 'customerPo', v)} />
                </td>
                <td className="border border-black p-0">
                   <InputCell value={item.uliPo} onChange={(v) => onItemChange(index, 'uliPo', v)} />
                </td>
                <td className="border border-black p-0">
                   <InputCell value={item.brand} onChange={(v) => onItemChange(index, 'brand', v)} />
                </td>
              </tr>
            ))}
            
            {/* Action Row (Not in export, just for UI) */}
            <tr className="no-print">
                <td colSpan={8} className="p-2 text-center border border-dashed border-gray-300">
                    <button 
                        onClick={onAddRow}
                        className="px-4 py-1 text-sm bg-blue-50 text-blue-600 rounded hover:bg-blue-100 transition-colors"
                    >
                        + Add Row
                    </button>
                </td>
            </tr>

            {/* Footer Total */}
            <tr className="h-10">
                <td colSpan={3} className="border border-black text-center font-bold align-middle bg-white">
                    合   計 TOTAL
                </td>
                <td className="border border-black bg-yellow-200 text-center font-bold align-middle">
                    {calculateTotal()}
                </td>
                <td className="border border-black"></td>
                <td className="border border-black"></td>
                <td className="border border-black"></td>
                <td className="border border-black"></td>
            </tr>

            {/* Signatures Row 1 */}
            <tr className="h-16">
                <td colSpan={3} className="border border-black p-2 align-bottom text-xs">
                    生管主管 (Production control)：
                </td>
                <td colSpan={2} className="border border-black p-2 align-bottom text-xs">
                    生管填表 (production fill in)：
                </td>
                <td colSpan={3} className="border border-black p-2 align-bottom text-xs">
                    業務確認(Business Unit)：
                </td>
            </tr>

            {/* Signatures Row 2 */}
            <tr className="h-16">
                <td colSpan={3} className="border border-black p-2 align-bottom text-xs">
                    資材主管 (Warehouse manage)：
                </td>
                <td colSpan={2} className="border border-black p-2 align-bottom text-xs">
                    成品倉 (finished goods warehouse)：
                </td>
                <td colSpan={3} className="border border-black p-2 align-bottom text-xs">
                    貨櫃檢驗確認 (Container examine)：
                </td>
            </tr>

            {/* Doc Info */}
            <tr className="h-8 border-none">
                 <td colSpan={2} className="text-xs pt-2"></td>
                 <td colSpan={4} className="text-xs pt-2 text-center">
                    Usia Penyimpanan : 1 tahun (保存年限：一年)
                 </td>
                 <td colSpan={2} className="text-xs pt-2 text-right font-bold">
                    Dok No : Form - PPIC - 03
                 </td>
            </tr>

          </tbody>
        </table>
      </div>
    </div>
  );
};

export default FormTable;`;

const SRC_SERVICES_EXCELSERVICE_TS = `import ExcelJS from 'exceljs';
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
  titleCell.value = 'FORM STUFFING LIST\\n裝  箱  單';
  titleCell.style = { 
    font: { bold: true, size: 16 }, 
    alignment: alignCenter,
    border: borderStyle
  };

  // --- Info Section (Rows 3-11) ---

  const addInfoRow = (rowNum: number, label: string, value: string, mergeValue: boolean = true) => {
    const labelCell = worksheet.getCell(\`B\${rowNum}\`);
    labelCell.value = label;
    labelCell.border = borderStyle;
    labelCell.alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };

    if (mergeValue) {
        worksheet.mergeCells(\`C\${rowNum}:D\${rowNum}\`);
    }
    
    const valueCell = worksheet.getCell(\`C\${rowNum}\`);
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
  addInfoRow(9, '貨櫃號 (Container NO)', data.containerNo);
  addInfoRow(10, '出貨單編號 (Delivery note No)', data.deliveryNoteNo);
  addInfoRow(11, '嘜頭 (Mark)', data.mark);

  // REMARK Block
  worksheet.mergeCells('E3:I3');
  const remarkLabel = worksheet.getCell('E3');
  remarkLabel.value = '備註說明\\nREMARK';
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
  worksheet.getCell('B12').value = '料  號';
  worksheet.getCell('C12').value = '品名/規格';
  worksheet.getCell('D12').value = '每箱數量';
  worksheet.getCell('E12').value = '箱數合計';
  worksheet.getCell('F12').value = '每箱包含的要點及顏色';
  worksheet.getCell('G12').value = '客戶PO';
  worksheet.getCell('H12').value = '工廠';
  worksheet.getCell('I12').value = '品牌';

  // Row 13
  worksheet.getRow(13).height = 20;
  worksheet.getCell('B13').value = 'Material No';
  worksheet.getCell('C13').value = '(Name and spec)';
  worksheet.getCell('D13').value = 'PCS/CTN';
  worksheet.getCell('E13').value = 'Total Ctn Qty';
  worksheet.getCell('F13').value = 'main point &color for each ctn\\n(務必要寫清楚 be detailed)';
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
    worksheet.getCell(\`B\${r}\`).value = item.materialNo;
    worksheet.getCell(\`C\${r}\`).value = item.nameAndSpec;
    worksheet.getCell(\`D\${r}\`).value = item.pcsPerCtn;
    worksheet.getCell(\`E\${r}\`).value = item.totalCtnQty; // Ensure number format if possible
    worksheet.getCell(\`F\${r}\`).value = item.description;
    worksheet.getCell(\`G\${r}\`).value = item.customerPo;
    worksheet.getCell(\`H\${r}\`).value = item.uliPo;
    worksheet.getCell(\`I\${r}\`).value = item.brand;

    // Styles
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(\`\${col}\${r}\`);
      cell.border = borderStyle;
      cell.alignment = alignCenter;
      cell.fill = yellowFill; // Data rows are yellow
    });

    currentRow++;
  });

  // Fill up to row 22 if data is short, to match template look
  while (currentRow <= 22) {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(\`\${col}\${currentRow}\`);
      cell.border = borderStyle;
      cell.alignment = alignCenter;
    });
    currentRow++;
  }

  // --- Footer ---
  const footerRowStart = 23;
  
  // Total Row
  worksheet.getCell(\`B\${footerRowStart}\`).value = '合   計 TOTAL';
  worksheet.getCell(\`B\${footerRowStart}\`).alignment = alignCenter;
  worksheet.getCell(\`B\${footerRowStart}\`).border = borderStyle;
  worksheet.mergeCells(\`B\${footerRowStart}:D\${footerRowStart}\`);
  
  // Calculate Total
  const totalQty = data.items.reduce((sum, item) => sum + (parseInt(item.totalCtnQty) || 0), 0);
  worksheet.getCell(\`E\${footerRowStart}\`).value = totalQty;
  worksheet.getCell(\`E\${footerRowStart}\`).fill = yellowFill;
  worksheet.getCell(\`E\${footerRowStart}\`).border = borderStyle;
  worksheet.getCell(\`E\${footerRowStart}\`).alignment = alignCenter;
  
  // Borders for rest of total row
  ['F', 'G', 'H', 'I'].forEach(col => {
      worksheet.getCell(\`\${col}\${footerRowStart}\`).border = borderStyle;
  });

  // Signatures Row 24
  const sigRow = 24;
  worksheet.mergeCells(\`B\${sigRow}:D\${sigRow}\`);
  worksheet.getCell(\`B\${sigRow}\`).value = '生管主管 (Production control)：';
  
  worksheet.mergeCells(\`E\${sigRow}:F\${sigRow}\`);
  worksheet.getCell(\`E\${sigRow}\`).value = '生管填表 (production fill in)：';
  
  worksheet.mergeCells(\`G\${sigRow}:I\${sigRow}\`);
  worksheet.getCell(\`G\${sigRow}\`).value = '業務確認(Business Unit)：';

  // Signatures Row 25
  const sigRow2 = 25;
  worksheet.mergeCells(\`B\${sigRow2}:D\${sigRow2}\`);
  worksheet.getCell(\`B\${sigRow2}\`).value = '資材主管 (Warehouse manage)：';

  worksheet.mergeCells(\`E\${sigRow2}:F\${sigRow2}\`);
  worksheet.getCell(\`E\${sigRow2}\`).value = '成品倉 (finished goods warehouse)：';

  worksheet.mergeCells(\`G\${sigRow2}:I\${sigRow2}\`);
  worksheet.getCell(\`G\${sigRow2}\`).value = '貨櫃檢驗確認 (Container examine)：';

  // Document Info Row 26
  const docRow = 26;
  worksheet.getCell(\`C\${docRow}\`).value = 'Usia Penyimpanan : 1 tahun (保存年限：一年)';
  
  worksheet.mergeCells(\`G\${docRow}:H\${docRow}\`);
  worksheet.getCell(\`G\${docRow}\`).value = 'Dok No : Form - PPIC - 03';

  // Generate Buffer
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, 'Stuffing_List.xlsx');
};`;

const SRC_APP_TSX = `import React, { useState } from 'react';
import FormTable from './components/FormTable';
import { FormData, INITIAL_DATA, OrderItem } from './types';
import { generateExcel } from './services/excelService';
import { Download, FileSpreadsheet, RefreshCw } from 'lucide-react';

const App: React.FC = () => {
  const [formData, setFormData] = useState<FormData>(INITIAL_DATA);
  const [isExporting, setIsExporting] = useState(false);

  const handleFieldChange = (field: keyof FormData, value: any) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleItemChange = (index: number, field: keyof OrderItem, value: string) => {
    const newItems = [...formData.items];
    newItems[index] = { ...newItems[index], [field]: value };
    setFormData(prev => ({ ...prev, items: newItems }));
  };

  const handleAddRow = () => {
    const newItem: OrderItem = {
      id: Math.random().toString(36).substr(2, 9),
      materialNo: '',
      nameAndSpec: '',
      pcsPerCtn: '',
      totalCtnQty: '',
      description: '',
      customerPo: '',
      uliPo: '',
      brand: ''
    };
    setFormData(prev => ({ ...prev, items: [...prev.items, newItem] }));
  };

  const handleExport = async () => {
    setIsExporting(true);
    try {
      await generateExcel(formData);
    } catch (error) {
      console.error('Export failed', error);
      alert('Failed to generate Excel file');
    } finally {
      setIsExporting(false);
    }
  };

  const handleReset = () => {
    if (window.confirm("Are you sure you want to reset the form?")) {
        setFormData(INITIAL_DATA);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Header / Toolbar */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-50 px-6 py-4 flex items-center justify-between shadow-sm no-print">
        <div className="flex items-center gap-2">
            <div className="bg-blue-600 p-2 rounded-lg text-white">
                <FileSpreadsheet size={24} />
            </div>
            <div>
                <h1 className="text-xl font-bold text-gray-800">Luggage Form Generator</h1>
                <p className="text-xs text-gray-500">Edit yellow cells directly matching the template</p>
            </div>
        </div>
        
        <div className="flex items-center gap-3">
          <button 
            onClick={handleReset}
            className="flex items-center gap-2 px-4 py-2 text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors font-medium text-sm"
          >
            <RefreshCw size={16} />
            Reset
          </button>
          
          <button 
            onClick={handleExport}
            disabled={isExporting}
            className="flex items-center gap-2 px-6 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors font-medium shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <Download size={18} />
            {isExporting ? 'Generating...' : 'Download .xlsx'}
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1">
        <FormTable 
            data={formData} 
            onChange={handleFieldChange} 
            onItemChange={handleItemChange}
            onAddRow={handleAddRow}
        />
      </main>

      {/* Footer Instructions */}
      <footer className="bg-white border-t p-4 text-center text-sm text-gray-500 no-print">
        <p>This tool mimics the "FORM STUFFING LIST" layout. Data is local to your browser session.</p>
        <p className="mt-1 text-xs">Generated files are compatible with Microsoft Excel and Google Sheets.</p>
      </footer>
    </div>
  );
};

export default App;`;

// --- ZIP Generation Logic ---

export const downloadProjectZip = async () => {
  const zip = new JSZip();

  // Root files
  zip.file('package.json', PACKAGE_JSON);
  zip.file('tsconfig.json', TSCONFIG_JSON);
  zip.file('tsconfig.node.json', TSCONFIG_NODE_JSON);
  zip.file('vite.config.ts', VITE_CONFIG_TS);
  zip.file('tailwind.config.js', TAILWIND_CONFIG_JS);
  zip.file('postcss.config.js', POSTCSS_CONFIG_JS);
  zip.file('index.html', INDEX_HTML);

  // Src files
  const src = zip.folder('src');
  if (src) {
      src.file('main.tsx', SRC_MAIN_TSX);
      src.file('App.tsx', SRC_APP_TSX);
      src.file('index.css', SRC_INDEX_CSS);
      src.file('types.ts', SRC_TYPES_TS);
      
      const services = src.folder('services');
      if (services) {
          services.file('excelService.ts', SRC_SERVICES_EXCELSERVICE_TS);
      }

      const components = src.folder('components');
      if (components) {
          components.file('FormTable.tsx', SRC_COMPONENTS_FORMTABLE_TSX);
      }
  }

  const blob = await zip.generateAsync({ type: 'blob' });
  saveAs(blob, 'luggage-form-generator-project.zip');
};
