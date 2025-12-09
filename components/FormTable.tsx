import React from 'react';
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
    className={`w-full h-full bg-yellow-200 p-1 resize-none text-center flex items-center justify-center border-none focus:ring-2 focus:ring-blue-500 text-sm font-medium text-black ${className}`}
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

export default FormTable;