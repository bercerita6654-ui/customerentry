import { Customer } from '../types';

const DEFAULT_SPREADSHEET_ID = '1b0vLGhuCqPekEEBP5Ps4ueGETyHedYLskduGmen_Qow';
const DEFAULT_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQNiYXFAms8RoacUxrg-dQd9k50D9g4dDx4ZSTPBtUyNuXK7OzvSFj6jLFw9juw8aN8LJwRmgOs19mF/pub?output=csv';

// Custom CSV parser that correctly handles comma-separated fields with quoted values (like addresses with commas)
export function parseCSV(csvText: string): string[][] {
  const lines: string[][] = [];
  let row: string[] = [];
  let inQuotes = false;
  let currentValue = '';

  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        currentValue += '"';
        i++; // skip next double quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      row.push(currentValue.trim());
      currentValue = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      if (char === '\r' && nextChar === '\n') {
        i++;
      }
      row.push(currentValue.trim());
      if (row.length > 0 && (row.length > 1 || row[0] !== '')) {
        lines.push(row);
      }
      row = [];
      currentValue = '';
    } else {
      currentValue += char;
    }
  }

  if (currentValue !== '' || row.length > 0) {
    row.push(currentValue.trim());
    lines.push(row);
  }

  return lines;
}

// Helper to extract detailed error message from Google Sheets API responses
async function parseGoogleApiError(response: Response, defaultMessage: string): Promise<string> {
  try {
    const text = await response.clone().text();
    try {
      const parsed = JSON.parse(text);
      if (parsed?.error?.message) {
        return parsed.error.message;
      }
    } catch {
      // Ignore JSON parse error
    }
    return text || response.statusText || defaultMessage;
  } catch (e) {
    return response.statusText || defaultMessage;
  }
}

export interface SheetMetadata {
  sheetId: number;
  title: string;
}

// Fetch all sheets' metadata from a spreadsheet to map sheetName to sheetId
export async function fetchSpreadsheetMetadata(
  spreadsheetId: string,
  accessToken: string
): Promise<SheetMetadata[]> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties(sheetId,title))`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) {
    const errorMessage = await parseGoogleApiError(response, response.statusText);
    console.error('Spreadsheet metadata error details:', errorMessage);
    throw new Error(`Gagal memuat metadata spreadsheet: ${errorMessage}`);
  }
  const data = await response.json();
  if (!data.sheets || data.sheets.length === 0) {
    throw new Error('Tidak ada sheet yang ditemukan di spreadsheet ini.');
  }
  return data.sheets.map((s: any) => ({
    sheetId: s.properties.sheetId,
    title: s.properties.title,
  }));
}

// Fetch customers from the public CSV url (No Auth required)
export async function fetchCustomersFromPublicCSV(csvUrl: string = DEFAULT_CSV_URL): Promise<Customer[]> {
  const res = await fetch(csvUrl);
  if (!res.ok) {
    throw new Error(`Gagal memuat data dari CSV publik: ${res.statusText}`);
  }
  const text = await res.text();
  const rows = parseCSV(text);
  if (rows.length < 2) return [];

  // Parse headers to match index. Sample headers:
  // NAMA CUSTOMER ,NO TELPN,ALAMAT,HARGA DIPAKAI,KETERANGAN ,PEMBAYARAN
  const parsedRows: Customer[] = [];
  
  // Header row is index 0. Actual data starts at index 1.
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.length < 3 && r.join('').trim() === '') continue; // skip empty rows
    parsedRows.push({
      id: i + 1, // Row number in sheet is 1-based, index 1 of array maps to Row 2
      name: r[0] || '',
      phone: r[1] || '',
      address: r[2] || '',
      priceUsed: r[3] || '',
      keterangan: r[4] || '',
      pembayaran: r[5] || '',
    });
  }
  return parsedRows;
}

// Fetch customers using Google Sheets API (Auth required, live, real-time data)
export async function fetchCustomersFromSheetsAPI(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<Customer[]> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A1:F1000`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) {
    const errorMessage = await parseGoogleApiError(response, response.statusText);
    console.error('Sheets API read error details:', errorMessage);
    throw new Error(`Gagal membaca spreadsheet: ${errorMessage}`);
  }
  const data = await response.json();
  const rows = data.values || [];
  if (rows.length < 2) return [];

  const parsedRows: Customer[] = [];
  // First row is header
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.length === 0) continue;
    parsedRows.push({
      id: i + 1, // Row number is 1-based, index i of values maps to Row i+1
      name: r[0] || '',
      phone: r[1] || '',
      address: r[2] || '',
      priceUsed: r[3] || '',
      keterangan: r[4] || '',
      pembayaran: r[5] || '',
    });
  }
  return parsedRows;
}

// Add a customer row
export async function addCustomer(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string,
  customer: Omit<Customer, 'id'>
): Promise<void> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}:append?valueInputOption=USER_ENTERED`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      range: sheetName,
      majorDimension: 'ROWS',
      values: [[
        customer.name,
        customer.phone,
        customer.address,
        customer.priceUsed,
        customer.keterangan,
        customer.pembayaran
      ]],
    }),
  });

  if (!response.ok) {
    const errorMessage = await parseGoogleApiError(response, response.statusText);
    console.error('Sheets API append error details:', errorMessage);
    throw new Error(`Gagal menambahkan data: ${errorMessage}`);
  }
}

// Update an existing customer row
export async function updateCustomer(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string,
  customer: Customer
): Promise<void> {
  // customer.id corresponds to the 1-based row number
  const range = `${sheetName}!A${customer.id}:F${customer.id}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;
  const response = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      range: range,
      majorDimension: 'ROWS',
      values: [[
        customer.name,
        customer.phone,
        customer.address,
        customer.priceUsed,
        customer.keterangan,
        customer.pembayaran
      ]],
    }),
  });

  if (!response.ok) {
    const errorMessage = await parseGoogleApiError(response, response.statusText);
    console.error('Sheets API update error details:', errorMessage);
    throw new Error(`Gagal memperbarui data baris ${customer.id}: ${errorMessage}`);
  }
}

// Delete an existing customer row via batchUpdate
export async function deleteCustomerRow(
  spreadsheetId: string,
  sheetId: number,
  accessToken: string,
  rowId: number // 1-based row number
): Promise<void> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: rowId - 1, // 0-based inclusive (Row 2 maps to startIndex 1)
              endIndex: rowId, // 0-based exclusive (Row 2 maps to endIndex 2)
            },
          },
        },
      ],
    }),
  });

  if (!response.ok) {
    const errorMessage = await parseGoogleApiError(response, response.statusText);
    console.error('Sheets API batchUpdate delete error details:', errorMessage);
    throw new Error(`Gagal menghapus data baris ${rowId}: ${errorMessage}`);
  }
}

// Fetch customers using Google Apps Script Web App
export async function fetchCustomersFromAppsScript(appsScriptUrl: string): Promise<Customer[]> {
  // Append timestamp to prevent aggressive browser caching of GET requests
  const separator = appsScriptUrl.includes('?') ? '&' : '?';
  const url = `${appsScriptUrl}${separator}_t=${Date.now()}`;
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Gagal memuat data dari Apps Script: ${response.statusText}`);
  }
  const data = await response.json();
  if (!Array.isArray(data)) {
    throw new Error('Data yang dikembalikan dari Apps Script tidak valid (harus berupa array).');
  }
  return data;
}

// Add a customer via Google Apps Script Web App
export async function addCustomerViaAppsScript(
  appsScriptUrl: string,
  customer: Omit<Customer, 'id'>
): Promise<void> {
  const response = await fetch(appsScriptUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8', // Text/plain avoids standard CORS pre-flight blocks for Web Apps
    },
    body: JSON.stringify({
      action: 'add',
      ...customer
    }),
  });

  if (!response.ok) {
    throw new Error(`Gagal menambahkan data via Apps Script: ${response.statusText}`);
  }
  
  const text = await response.text();
  try {
    const result = JSON.parse(text);
    if (result && result.success === false) {
      throw new Error(result.message || 'Gagal menambahkan data via Apps Script.');
    }
  } catch (e) {
    // Some Apps Script web apps return redirect/html, but text/plain can be treated as success if response was ok
    console.warn('Response from Apps Script was not JSON, but status is OK:', text);
  }
}

// Update a customer via Google Apps Script Web App
export async function updateCustomerViaAppsScript(
  appsScriptUrl: string,
  customer: Customer
): Promise<void> {
  const response = await fetch(appsScriptUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8',
    },
    body: JSON.stringify({
      action: 'update',
      ...customer
    }),
  });

  if (!response.ok) {
    throw new Error(`Gagal memperbarui data via Apps Script: ${response.statusText}`);
  }
  
  const text = await response.text();
  try {
    const result = JSON.parse(text);
    if (result && result.success === false) {
      throw new Error(result.message || 'Gagal memperbarui data via Apps Script.');
    }
  } catch (e) {
    console.warn('Response from Apps Script was not JSON, but status is OK:', text);
  }
}

// Delete a customer via Google Apps Script Web App
export async function deleteCustomerViaAppsScript(
  appsScriptUrl: string,
  rowId: number
): Promise<void> {
  const response = await fetch(appsScriptUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8',
    },
    body: JSON.stringify({
      action: 'delete',
      id: rowId
    }),
  });

  if (!response.ok) {
    throw new Error(`Gagal menghapus data via Apps Script: ${response.statusText}`);
  }
  
  const text = await response.text();
  try {
    const result = JSON.parse(text);
    if (result && result.success === false) {
      throw new Error(result.message || 'Gagal menghapus data via Apps Script.');
    }
  } catch (e) {
    console.warn('Response from Apps Script was not JSON, but status is OK:', text);
  }
}
