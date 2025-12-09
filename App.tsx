import React, { useState } from 'react';
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
                    className={`flex items-center gap-1 text-xs px-2 py-1 rounded cursor-pointer border ${siFile ? 'bg-green-100 border-green-300 text-green-800' : 'bg-white border-gray-300 text-gray-600'}`}
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
                    className={`flex items-center gap-1 text-xs px-2 py-1 rounded cursor-pointer border ${indexFile ? 'bg-green-100 border-green-300 text-green-800' : 'bg-white border-gray-300 text-gray-600'}`}
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

export default App;