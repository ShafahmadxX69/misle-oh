export interface OrderItem {
  id: string;
  materialNo: string;
  nameAndSpec: string;
  pcsPerCtn: string; // Using string to allow mixed input if needed, though number is typical
  totalCtnQty: string;
  description: string;
  customerPo: string;
  uliPo: string;
  brand: string;
  sku?: string; // Internal field for macro processing (Column P in VBA)
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
  containerSizeInfo: string; // For cell B14 (e.g., 1.1*40HQ(2300"))
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
};