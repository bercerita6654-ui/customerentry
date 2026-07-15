export interface Customer {
  id: number; // Represents the 1-based row index in Google Sheets
  name: string;
  phone: string;
  address: string;
  priceUsed: string;
  keterangan: string;
  pembayaran: string;
}

export interface SpreadsheetConfig {
  spreadsheetId: string;
  sheetName: string;
}
