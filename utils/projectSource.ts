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
    "react-dom": "^18.2.0",
    "jszip": "^3.10.1"
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
  sku?: string;
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
import { downloadProjectZip } from './utils/projectSource';
import { processAutomation, calculateCUFTReport } from './services/macroService';
import { Download, FileSpreadsheet, RefreshCw, FileCode, Play, Upload, Calculator } from 'lucide-react';

const App: React.FC = () => {
  const [formData, setFormData] = useState<FormData>(INITIAL_DATA);
  const [isExporting, setIsExporting] = useState(false);
  const [isZipping, setIsZipping] = useState(false);
  const [siFile, setSiFile] = useState<File | null>(null);
  const [indexFile, setIndexFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);

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

  const handleDownloadSource = async () => {
      if (!window.confirm("Download the full source code project (ZIP)?")) return;
      setIsZipping(true);
      try {
          await downloadProjectZip();
      } catch (error) {
          console.error('Zip failed', error);
          alert('Failed to generate project ZIP');
      } finally {
          setIsZipping(false);
      }
  };

  const handleReset = () => {
    if (window.confirm("Are you sure you want to reset the form?")) {
        setFormData(INITIAL_DATA);
        setSiFile(null);
        setIndexFile(null);
    }
  };

  const handleProcessMacros = async () => {
    if (!siFile && !indexFile) {
        alert("Please upload at least 'SI' file or 'Index.xlsx' to run processing.");
        return;
    }
    
    setIsProcessing(true);
    try {
        const newData = await processAutomation(formData, siFile, indexFile);
        setFormData(newData);
        alert("Automation (Macros) processed successfully!");
    } catch (e) {
        console.error("Macro Processing Failed", e);
        alert("Processing Failed.");
    } finally {
        setIsProcessing(false);
    }
  };

  const handleCalculateCUFT = () => {
    const report = calculateCUFTReport(formData);
    alert(report);
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Header / Toolbar */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-50 px-6 py-4 flex flex-col md:flex-row md:items-center justify-between shadow-sm no-print gap-4">
        <div className="flex items-center gap-2">
            <div className="bg-blue-600 p-2 rounded-lg text-white">
                <FileSpreadsheet size={24} />
            </div>
            <div>
                <h1 className="text-xl font-bold text-gray-800">Luggage Form Generator</h1>
                <p className="text-xs text-gray-500">Edit yellow cells or upload SI/Index for automation</p>
            </div>
        </div>

        {/* File Upload Section */}
        <div className="flex items-center gap-2 bg-gray-50 p-2 rounded-md border border-gray-200">
             <div className="relative group">
                <input 
                    type="file" 
                    id="siFile" 
                    className="hidden" 
                    accept=".xlsx, .xlsm"
                    onChange={(e) => setSiFile(e.target.files?.[0] || null)}
                />
                <label 
                    htmlFor="siFile"
                    className={\`flex items-center gap-1 text-xs px-2 py-1 rounded cursor-pointer border \${siFile ? 'bg-green-100 border-green-300 text-green-800' : 'bg-white border-gray-300 text-gray-600'}\`}
                >
                    <Upload size={12} />
                    {siFile ? siFile.name.substring(0, 10) + '...' : 'Upload SI'}
                </label>
             </div>
             
             <div className="relative group">
                <input 
                    type="file" 
                    id="indexFile" 
                    className="hidden" 
                    accept=".xlsx, .xlsm"
                    onChange={(e) => setIndexFile(e.target.files?.[0] || null)}
                />
                <label 
                    htmlFor="indexFile"
                    className={\`flex items-center gap-1 text-xs px-2 py-1 rounded cursor-pointer border \${indexFile ? 'bg-green-100 border-green-300 text-green-800' : 'bg-white border-gray-300 text-gray-600'}\`}
                >
                    <Upload size={12} />
                    {indexFile ? indexFile.name.substring(0, 10) + '...' : 'Upload Index'}
                </label>
             </div>

             <button 
                onClick={handleProcessMacros}
                disabled={isProcessing}
                className="flex items-center gap-1 px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-800 rounded text-xs font-semibold transition-colors"
                title="Imports SI Data, Matches Colors/Materials from Index, and aggregates SO/PO"
             >
                <Play size={12} />
                {isProcessing ? 'Processing...' : 'Run Macros'}
             </button>
        </div>
        
        {/* Actions */}
        <div className="flex items-center gap-2">
          <button
            onClick={handleCalculateCUFT}
            className="p-2 text-gray-600 hover:text-blue-600 hover:bg-gray-100 rounded-full transition-colors"
            title="Calculate CUFT"
          >
            <Calculator size={18} />
          </button>

          <button 
            onClick={handleDownloadSource}
            disabled={isZipping}
            className="hidden md:flex items-center gap-2 px-3 py-2 text-gray-600 bg-white border border-gray-300 hover:bg-gray-50 rounded-lg transition-colors font-medium text-sm"
            title="Download full project source code"
          >
             <FileCode size={16} />
          </button>

          <button 
            onClick={handleReset}
            className="flex items-center gap-2 px-3 py-2 text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors font-medium text-sm"
          >
            <RefreshCw size={16} />
          </button>
          
          <button 
            onClick={handleExport}
            disabled={isExporting}
            className="flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors font-medium shadow-md disabled:opacity-50 disabled:cursor-not-allowed text-sm"
          >
            <Download size={18} />
            {isExporting ? '...' : 'Export'}
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
        <p>1. Upload <b>SI File</b> to populate items. 2. Upload <b>Index.xlsx</b> to match materials/colors. 3. Click <b>Run Macros</b>.</p>
        <p className="mt-1 text-xs">Generated files are compatible with Microsoft Excel and Google Sheets.</p>
      </footer>
    </div>
  );
};

export default App;`;

const SRC_SERVICES_MACROSERVICE_TS = `import ExcelJS from 'exceljs';
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
          newData.customer = \`\${brandVal} TO \${destVal}\`;
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
                // But in \`ImportSI\`, description is set to \`wsSI...("COLOR")\`.
                
                // Let's try flexible matching for color
                const colorMatch = idx.colorP && colorCodeTemplate.replace(/\\s/g, "").includes(idx.colorP.replace(/\\s/g, ""));
                
                return idx.so === soTemplate && idxModel === modelTemplate && colorMatch;
            });

            if (match) {
                // Apply ColorCode Logic
                const newDesc = skuCode ? \`\${match.colorMandarin} \${skuCode}\` : match.colorMandarin;
                
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

    return \`Total Calculated CUFT: \${totalCUFT.toFixed(2)}\`;
};`;

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
          services.file('macroService.ts', SRC_SERVICES_MACROSERVICE_TS);
      }

      const components = src.folder('components');
      if (components) {
          components.file('FormTable.tsx', SRC_COMPONENTS_FORMTABLE_TSX);
      }
  }

  const blob = await zip.generateAsync({ type: 'blob' });
  saveAs(blob, 'luggage-form-generator-project.zip');
};