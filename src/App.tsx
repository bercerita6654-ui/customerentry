import React, { useState, useEffect, useMemo } from 'react';
import { 
  Plus, 
  Search, 
  Settings, 
  LogOut, 
  MapPin, 
  Phone, 
  DollarSign, 
  CheckCircle, 
  Clock, 
  ArrowUpRight, 
  Database, 
  AlertCircle, 
  Filter, 
  Download, 
  HelpCircle, 
  RefreshCw, 
  FileSpreadsheet, 
  Info, 
  X, 
  User, 
  Check, 
  Trash2,
  Edit2,
  Lock,
  ArrowRight,
  Copy,
  Code,
  Printer,
  ArrowUp,
  ArrowDown,
  Bold,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Type,
  ChevronUp,
  ChevronDown,
  RotateCcw,
  FileText
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { auth, initAuth, googleSignIn, logout, getAccessToken } from './firebase';
import { 
  fetchCustomersFromPublicCSV, 
  fetchCustomersFromSheetsAPI, 
  addCustomer, 
  updateCustomer, 
  deleteCustomerRow,
  fetchSpreadsheetMetadata,
  SheetMetadata,
  fetchCustomersFromAppsScript,
  addCustomerViaAppsScript,
  updateCustomerViaAppsScript,
  deleteCustomerViaAppsScript
} from './services/sheetsService';
import { Customer } from './types';

export interface PrintLine {
  text: string;
  isBold: boolean;
  size: 'sm' | 'md' | 'lg' | 'xl';
  align: 'left' | 'center' | 'right';
}

// Default config values
const DEFAULT_SPREADSHEET_ID = '1b0vLGhuCqPekEEBP5Ps4ueGETyHedYLskduGmen_Qow';
const DEFAULT_SHEET_NAME = 'Sheet1';
const DEFAULT_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQNiYXFAms8RoacUxrg-dQd9k50D9g4dDx4ZSTPBtUyNuXK7OzvSFj6jLFw9juw8aN8LJwRmgOs19mF/pub?output=csv';

export default function App() {
  // Auth states
  const [currentUser, setCurrentUser] = useState<any>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [needsAuth, setNeedsAuth] = useState(false);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  // App settings states (loaded from localStorage or defaulted)
  const [spreadsheetId, setSpreadsheetId] = useState(() => {
    return localStorage.getItem('rekap_spreadsheet_id') || DEFAULT_SPREADSHEET_ID;
  });
  const [sheetName, setSheetName] = useState(() => {
    return localStorage.getItem('rekap_sheet_name') || DEFAULT_SHEET_NAME;
  });
  const [sheetId, setSheetId] = useState<number>(() => {
    const saved = localStorage.getItem('rekap_sheet_id');
    return saved ? parseInt(saved, 10) : 0;
  });
  const [availableSheets, setAvailableSheets] = useState<SheetMetadata[]>([]);

  // Customer data states
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  // Form states
  const [formName, setFormName] = useState('');
  const [formPhone, setFormPhone] = useState('');
  const [formAddress, setFormAddress] = useState('');
  const [formPriceUsed, setFormPriceUsed] = useState('');
  const [formKeterangan, setFormKeterangan] = useState('');
  const [formPembayaran, setFormPembayaran] = useState('');
  const [editingCustomer, setEditingCustomer] = useState<Customer | null>(null);

  // UI state managers
  const [searchQuery, setSearchQuery] = useState('');
  const [filterKeterangan, setFilterKeterangan] = useState('all');
  const [filterPembayaran, setFilterPembayaran] = useState('all');
  const [showConfigModal, setShowConfigModal] = useState(false);
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [tempSpreadsheetId, setTempSpreadsheetId] = useState(spreadsheetId);
  const [tempSheetName, setTempSheetName] = useState(sheetName);

  // Thermal print states
  const [printingCustomer, setPrintingCustomer] = useState<Customer | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [printText, setPrintText] = useState('');
  const [printPaperWidth, setPrintPaperWidth] = useState<'58mm' | '80mm'>('58mm');
  const [optDibantuSiapkan, setOptDibantuSiapkan] = useState(true);
  const [optKeterangan, setOptKeterangan] = useState(true);
  const [optPembayaran, setOptPembayaran] = useState(true);
  const [printLines, setPrintLines] = useState<PrintLine[]>([]);
  const [printFontFamily, setPrintFontFamily] = useState<'mono' | 'sans' | 'serif'>('sans');
  const [printBaseSize, setPrintBaseSize] = useState<'xs' | 'sm' | 'base' | 'lg'>('sm');
  const [activeLineIndex, setActiveLineIndex] = useState<number | null>(0);

  // Google Apps Script States
  const [connectionMode, setConnectionMode] = useState<'sheets-api' | 'apps-script'>(() => {
    return (localStorage.getItem('rekap_connection_mode') as 'sheets-api' | 'apps-script') || 'sheets-api';
  });
  const [appsScriptUrl, setAppsScriptUrl] = useState(() => {
    return localStorage.getItem('rekap_apps_script_url') || '';
  });
  const [tempConnectionMode, setTempConnectionMode] = useState<'sheets-api' | 'apps-script'>(connectionMode);
  const [tempAppsScriptUrl, setTempAppsScriptUrl] = useState(appsScriptUrl);
  const [copiedScript, setCopiedScript] = useState(false);

  // Auto-complete presets for easier typing
  const pricePresets = [
    'harga 1',
    'harga 2',
    'harga 2 (kalau dibawah 12 harga 1)'
  ];

  const statusPresets = [
    'DI KIRIM',
    'INSTAN COURIER (GOJEK, GOCAR,dll)',
    'DIAMBIL'
  ];

  const paymentPresets = [
    'TRANSFER',
    'CASH',
    'KREDIT'
  ];

  // 1. Initialize Auth on mount
  useEffect(() => {
    const unsubscribe = initAuth(
      (user, token) => {
        setCurrentUser(user);
        setAccessToken(token);
        setNeedsAuth(false);
      },
      () => {
        setCurrentUser(null);
        setAccessToken(null);
        setNeedsAuth(true);
      }
    );
    return () => unsubscribe();
  }, []);

  // Helper to handle Google Auth credential expiration or errors
  const checkAndHandleAuthError = async (err: any) => {
    const errorMsg = String(err?.message || err || '').toLowerCase();
    if (
      errorMsg.includes('invalid authentication credentials') ||
      errorMsg.includes('expected oauth 2 access token') ||
      errorMsg.includes('unauthorized') ||
      errorMsg.includes('401') ||
      errorMsg.includes('invalid credentials') ||
      errorMsg.includes('auth') ||
      errorMsg.includes('token')
    ) {
      console.warn('Authentication error detected, logging out and clearing stale token:', errorMsg);
      try {
        await logout();
      } catch (logoutErr) {
        console.error('Error during automatic logout:', logoutErr);
      }
      setAccessToken(null);
      setCurrentUser(null);
      setNeedsAuth(true);
      setError('Sesi Google Anda telah berakhir atau tidak valid. Silakan klik tombol "Masuk dengan Google" kembali untuk memperbarui akses.');
      return true;
    }
    return false;
  };

  // 2. Fetch customers data when auth token or config changes
  const loadData = async (forceAuthToken: string | null = accessToken) => {
    setLoading(true);
    setError(null);
    try {
      if (connectionMode === 'apps-script') {
        if (!appsScriptUrl.trim()) {
          setCustomers([]);
          setError('Silakan konfigurasikan Google Apps Script Web App URL di pengaturan (klik ikon gir di kanan atas).');
          return;
        }
        const data = await fetchCustomersFromAppsScript(appsScriptUrl);
        setCustomers(data);
      } else if (forceAuthToken) {
        // Authenticated mode: Load directly via Google Sheets API (live!)
        try {
          // Fetch sheets metadata to find available sheet names & sheetId for deletion
          const sheets = await fetchSpreadsheetMetadata(spreadsheetId, forceAuthToken);
          setAvailableSheets(sheets);
          
          // Verify if our selected sheet exists, otherwise set to first available sheet
          const currentSheetExists = sheets.some(s => s.title.toLowerCase() === sheetName.toLowerCase());
          let targetSheetName = sheetName;
          if (sheets.length > 0 && !currentSheetExists) {
            targetSheetName = sheets[0].title;
            setSheetName(targetSheetName);
            setSheetId(sheets[0].sheetId);
            localStorage.setItem('rekap_sheet_name', targetSheetName);
            localStorage.setItem('rekap_sheet_id', sheets[0].sheetId.toString());
          } else {
            const foundSheet = sheets.find(s => s.title.toLowerCase() === sheetName.toLowerCase());
            if (foundSheet) {
              setSheetId(foundSheet.sheetId);
              localStorage.setItem('rekap_sheet_id', foundSheet.sheetId.toString());
            }
          }

          const data = await fetchCustomersFromSheetsAPI(spreadsheetId, targetSheetName, forceAuthToken);
          setCustomers(data);
        } catch (apiErr: any) {
          console.error('Failed reading live Sheets API, falling back to public CSV', apiErr);
          const isAuthError = await checkAndHandleAuthError(apiErr);
          if (!isAuthError) {
            // Fallback to public CSV if Sheet ID matches the default one
            if (spreadsheetId === DEFAULT_SPREADSHEET_ID) {
              const data = await fetchCustomersFromPublicCSV(DEFAULT_CSV_URL);
              setCustomers(data);
              setError('Menggunakan data cache publik. Login kembali atau pastikan hak akses jika ingin mengedit.');
            } else {
              throw apiErr;
            }
          }
        }
      } else {
        // Public / Fallback mode: Load from published CSV (no credentials needed)
        // Only load if utilizing the default spreadsheet
        if (spreadsheetId === DEFAULT_SPREADSHEET_ID) {
          const data = await fetchCustomersFromPublicCSV(DEFAULT_CSV_URL);
          setCustomers(data);
        } else {
          setCustomers([]);
          setError('Silakan masuk dengan akun Google untuk membaca spreadsheet kustom Anda.');
        }
      }
    } catch (err: any) {
      console.error('Error fetching data:', err);
      const isAuthError = await checkAndHandleAuthError(err);
      if (!isAuthError) {
        setError(err.message || 'Gagal mengambil data dari Google Sheets. Pastikan Spreadsheet ID benar dan telah dibagikan.');
      }
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [spreadsheetId, sheetName, accessToken, connectionMode, appsScriptUrl]);

  // Function to build default print lines from customer & option states
  const generateDefaultPrintLines = (customer: Customer, showSiapkan: boolean, showKet: boolean, showPem: boolean): PrintLine[] => {
    const lines: PrintLine[] = [];
    
    // 1. Name & Phone
    const nameUpper = customer.name.toUpperCase();
    const phoneStr = customer.phone ? ` (${customer.phone})` : '';
    lines.push({
      text: `${nameUpper}${phoneStr}`,
      isBold: true,
      size: 'md',
      align: 'left'
    });
    
    // 2. Address
    if (customer.address) {
      lines.push({
        text: customer.address,
        isBold: false,
        size: 'sm',
        align: 'left'
      });
    }
    
    // 3. Dibantu Siapkan
    if (showSiapkan) {
      lines.push({
        text: '- DIBANTU SIAPKAN',
        isBold: false,
        size: 'sm',
        align: 'left'
      });
    }
    
    // 4. Keterangan
    if (showKet) {
      const ketText = customer.keterangan ? `- ${customer.keterangan.toUpperCase()}` : '- DIKIRIM';
      lines.push({
        text: ketText,
        isBold: false,
        size: 'sm',
        align: 'left'
      });
    }
    
    // 5. Pembayaran
    if (showPem) {
      const pemText = customer.pembayaran ? `- ${customer.pembayaran.toUpperCase()}` : '- TRANSFER';
      lines.push({
        text: pemText,
        isBold: false,
        size: 'sm',
        align: 'left'
      });
    }
    
    return lines;
  };

  // Effect to set up lines when customer or simple options change
  useEffect(() => {
    if (printingCustomer) {
      const defaultLines = generateDefaultPrintLines(
        printingCustomer, 
        optDibantuSiapkan, 
        optKeterangan, 
        optPembayaran
      );
      setPrintLines(defaultLines);
    }
  }, [printingCustomer, optDibantuSiapkan, optKeterangan, optPembayaran]);

  // Effect to dynamically update Thermal Print text template (for copy/paste or text editor)
  useEffect(() => {
    if (printLines.length > 0) {
      const plainText = printLines.map(line => line.text).join('\n');
      setPrintText(plainText);
    }
  }, [printLines]);

  const handleAddLine = (index?: number) => {
    const insertAt = index !== undefined ? index + 1 : printLines.length;
    const newLine: PrintLine = { text: '', isBold: false, size: 'sm', align: 'left' };
    const newLines = [...printLines];
    newLines.splice(insertAt, 0, newLine);
    setPrintLines(newLines);
    setActiveLineIndex(insertAt);
    
    setTimeout(() => {
      const inputEl = document.getElementById(`line-input-${insertAt}`);
      if (inputEl) {
        (inputEl as HTMLInputElement).focus();
      }
    }, 50);
  };

  const handleDeleteLine = (index: number) => {
    if (printLines.length <= 1) {
      const newLines = [...printLines];
      newLines[0] = { text: '', isBold: false, size: 'sm', align: 'left' };
      setPrintLines(newLines);
      setActiveLineIndex(0);
      return;
    }
    const newLines = printLines.filter((_, i) => i !== index);
    setPrintLines(newLines);
    
    const nextFocusIndex = Math.max(0, index - 1);
    setActiveLineIndex(nextFocusIndex);
    setTimeout(() => {
      const inputEl = document.getElementById(`line-input-${nextFocusIndex}`);
      if (inputEl) {
        (inputEl as HTMLInputElement).focus();
      }
    }, 50);
  };

  const handleLineKeyDown = (e: React.KeyboardEvent<HTMLInputElement>, idx: number) => {
    const inputEl = e.currentTarget;
    const val = inputEl.value;
    const start = inputEl.selectionStart ?? 0;
    const end = inputEl.selectionEnd ?? 0;

    if (e.key === 'ArrowUp') {
      if (idx > 0) {
        e.preventDefault();
        setActiveLineIndex(idx - 1);
        setTimeout(() => {
          const prevInput = document.getElementById(`line-input-${idx - 1}`) as HTMLInputElement | null;
          if (prevInput) {
            prevInput.focus();
            const newPos = Math.min(start, prevInput.value.length);
            prevInput.setSelectionRange(newPos, newPos);
          }
        }, 10);
      }
    } else if (e.key === 'ArrowDown') {
      if (idx < printLines.length - 1) {
        e.preventDefault();
        setActiveLineIndex(idx + 1);
        setTimeout(() => {
          const nextInput = document.getElementById(`line-input-${idx + 1}`) as HTMLInputElement | null;
          if (nextInput) {
            nextInput.focus();
            const newPos = Math.min(start, nextInput.value.length);
            nextInput.setSelectionRange(newPos, newPos);
          }
        }, 10);
      }
    } else if (e.key === 'Enter') {
      e.preventDefault();
      // Split current line at cursor
      const leftText = val.substring(0, start);
      const rightText = val.substring(end);

      const newLines = [...printLines];
      newLines[idx] = { ...newLines[idx], text: leftText };
      
      const newLine: PrintLine = {
        text: rightText,
        isBold: newLines[idx].isBold,
        size: newLines[idx].size,
        align: newLines[idx].align
      };
      newLines.splice(idx + 1, 0, newLine);
      setPrintLines(newLines);
      setActiveLineIndex(idx + 1);

      setTimeout(() => {
        const nextInput = document.getElementById(`line-input-${idx + 1}`) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.setSelectionRange(0, 0);
        }
      }, 10);
    } else if (e.key === 'Backspace') {
      // If cursor is at the very beginning of the line, merge with the line above
      if (start === 0 && end === 0) {
        if (idx > 0) {
          e.preventDefault();
          const prevLineText = printLines[idx - 1].text;
          const currentLineText = val;
          
          const newLines = printLines.filter((_, i) => i !== idx);
          newLines[idx - 1] = {
            ...newLines[idx - 1],
            text: prevLineText + currentLineText
          };
          
          setPrintLines(newLines);
          setActiveLineIndex(idx - 1);

          setTimeout(() => {
            const prevInput = document.getElementById(`line-input-${idx - 1}`) as HTMLInputElement | null;
            if (prevInput) {
              prevInput.focus();
              prevInput.setSelectionRange(prevLineText.length, prevLineText.length);
            }
          }, 10);
        }
      }
    }
  };

  const handleStartPrint = (customer: Customer) => {
    setPrintingCustomer(customer);
    setOptDibantuSiapkan(true);
    setOptKeterangan(true);
    setOptPembayaran(true);
    setPrintFontFamily('sans');
    setPrintBaseSize('sm');
    
    const initialLines = generateDefaultPrintLines(customer, true, true, true);
    setPrintLines(initialLines);
    setActiveLineIndex(0);
    setShowPrintModal(true);
  };

  // Auth Action Handlers
  const handleLogin = async () => {
    setIsLoggingIn(true);
    setError(null);
    try {
      const result = await googleSignIn();
      if (result) {
        setAccessToken(result.accessToken);
        setCurrentUser(result.user);
        setNeedsAuth(false);
        showToast('Berhasil masuk dengan Google!');
        // Trigger data load with the fresh token
        await loadData(result.accessToken);
      }
    } catch (err: any) {
      console.error('Login error:', err);
      setError('Gagal masuk dengan Google: ' + (err.message || err));
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    try {
      await logout();
      setCurrentUser(null);
      setAccessToken(null);
      setNeedsAuth(true);
      setAvailableSheets([]);
      showToast('Berhasil keluar.');
    } catch (err: any) {
      setError('Gagal keluar: ' + err.message);
    }
  };

  // Toast Helper
  const showToast = (message: string) => {
    setSuccessMessage(message);
    setTimeout(() => {
      setSuccessMessage(null);
    }, 4000);
  };

  // Phone number sanitizer to replace +62/62 with 0
  const sanitizePhoneNumber = (val: string): string => {
    let trimmed = val.trim();
    if (trimmed.startsWith('+62')) {
      trimmed = '0' + trimmed.slice(3);
    } else if (trimmed.startsWith('62')) {
      trimmed = '0' + trimmed.slice(2);
    }
    return trimmed;
  };

  // Form submit (Add or Edit)
  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formName.trim()) {
      setError('Nama customer wajib diisi.');
      return;
    }

    setSaving(true);
    setError(null);

    const payload = {
      name: formName.trim(),
      phone: sanitizePhoneNumber(formPhone),
      address: formAddress.trim(),
      priceUsed: formPriceUsed.trim(),
      keterangan: formKeterangan.trim(),
      pembayaran: formPembayaran.trim()
    };

    try {
      if (connectionMode === 'apps-script') {
        if (!appsScriptUrl.trim()) {
          setError('Google Apps Script Web App URL tidak diatur. Silakan konfigurasikan di pengaturan.');
          setSaving(false);
          return;
        }

        if (editingCustomer) {
          // Edit Mode Apps Script
          const updatedCustomer: Customer = {
            ...editingCustomer,
            ...payload
          };

          const confirmed = window.confirm(`Apakah Anda yakin ingin memperbarui data untuk customer "${formName}" via Google Apps Script?`);
          if (!confirmed) {
            setSaving(false);
            return;
          }

          await updateCustomerViaAppsScript(appsScriptUrl, updatedCustomer);
          showToast(`Berhasil memperbarui data customer "${formName}" (via Apps Script)`);
          setEditingCustomer(null);
        } else {
          // Add Mode Apps Script
          await addCustomerViaAppsScript(appsScriptUrl, payload);
          showToast(`Berhasil menambahkan customer "${formName}" (via Apps Script)`);
        }

        resetForm();
        await loadData();
      } else {
        // Direct Sheets API Mode
        if (!accessToken) {
          setError('Anda harus masuk dengan Google untuk menginput data.');
          setSaving(false);
          return;
        }

        if (editingCustomer) {
          // Edit Mode Direct API
          const updatedCustomer: Customer = {
            ...editingCustomer,
            ...payload
          };

          // Mutation confirmation warning for editing data (mandated by workspace-integration skill)
          const confirmed = window.confirm(`Apakah Anda yakin ingin memperbarui data untuk customer "${formName}" di Google Sheet?`);
          if (!confirmed) {
            setSaving(false);
            return;
          }

          await updateCustomer(spreadsheetId, sheetName, accessToken, updatedCustomer);
          showToast(`Berhasil memperbarui data customer "${formName}"`);
          setEditingCustomer(null);
        } else {
          // Add Mode Direct API
          await addCustomer(spreadsheetId, sheetName, accessToken, payload);
          showToast(`Berhasil menambahkan customer "${formName}" ke Google Sheet`);
        }

        // Reset form
        resetForm();
        // Reload spreadsheet data
        await loadData();
      }
    } catch (err: any) {
      console.error('Submit error:', err);
      const isAuthError = await checkAndHandleAuthError(err);
      if (!isAuthError) {
        setError('Gagal menyimpan data: ' + (err.message || err));
      }
    } finally {
      setSaving(false);
    }
  };

  // Start editing a customer (populates the form)
  const handleStartEdit = (customer: Customer) => {
    setEditingCustomer(customer);
    setFormName(customer.name);
    setFormPhone(customer.phone);
    setFormAddress(customer.address);
    setFormPriceUsed(customer.priceUsed);
    setFormKeterangan(customer.keterangan);
    setFormPembayaran(customer.pembayaran);
    
    // Scroll smoothly to form on mobile devices
    window.scrollTo({ top: 180, behavior: 'smooth' });
  };

  // Delete customer row
  const handleDelete = async (customer: Customer) => {
    if (connectionMode === 'apps-script') {
      if (!appsScriptUrl.trim()) {
        setError('Google Apps Script Web App URL tidak diatur. Silakan atur di menu pengaturan.');
        return;
      }

      const confirmed = window.confirm(
        `Apakah Anda yakin ingin menghapus data customer "${customer.name}" dari baris ${customer.id}? Tindakan ini akan menghapus baris tersebut secara permanen dari Google Sheet via Apps Script.`
      );
      if (!confirmed) return;

      setLoading(true);
      setError(null);
      try {
        await deleteCustomerViaAppsScript(appsScriptUrl, customer.id);
        showToast(`Berhasil menghapus customer "${customer.name}" (via Apps Script)`);
        if (editingCustomer?.id === customer.id) {
          resetForm();
        }
        await loadData();
      } catch (err: any) {
        console.error('Delete error:', err);
        setError('Gagal menghapus data via Apps Script: ' + (err.message || err));
      } finally {
        setLoading(false);
      }
    } else {
      // Direct API Mode
      if (!accessToken) {
        setError('Anda harus masuk dengan Google untuk menghapus data.');
        return;
      }

      // Deletion confirmation dialog (mandated by workspace-integration skill)
      const confirmed = window.confirm(
        `Apakah Anda yakin ingin menghapus data customer "${customer.name}" dari baris ${customer.id}? Tindakan ini akan menghapus baris tersebut secara permanen dari Google Sheet.`
      );
      if (!confirmed) return;

      setLoading(true);
      setError(null);
      try {
        await deleteCustomerRow(spreadsheetId, sheetId, accessToken, customer.id);
        showToast(`Berhasil menghapus customer "${customer.name}"`);
        if (editingCustomer?.id === customer.id) {
          resetForm();
        }
        await loadData();
      } catch (err: any) {
        console.error('Delete error:', err);
        const isAuthError = await checkAndHandleAuthError(err);
        if (!isAuthError) {
          setError('Gagal menghapus data: ' + (err.message || err));
        }
      } finally {
        setLoading(false);
      }
    }
  };

  const resetForm = () => {
    setEditingCustomer(null);
    setFormName('');
    setFormPhone('');
    setFormAddress('');
    setFormPriceUsed('');
    setFormKeterangan('');
    setFormPembayaran('');
  };

  // Save Config Settings
  const handleSaveConfig = () => {
    // Save Apps Script setting
    setConnectionMode(tempConnectionMode);
    setAppsScriptUrl(tempAppsScriptUrl);
    localStorage.setItem('rekap_connection_mode', tempConnectionMode);
    localStorage.setItem('rekap_apps_script_url', tempAppsScriptUrl);

    if (tempConnectionMode === 'sheets-api') {
      if (!tempSpreadsheetId.trim()) {
        setError('Spreadsheet ID tidak boleh kosong.');
        return;
      }
      const cleanId = tempSpreadsheetId.trim();
      const cleanName = tempSheetName.trim() || DEFAULT_SHEET_NAME;

      setSpreadsheetId(cleanId);
      setSheetName(cleanName);
      localStorage.setItem('rekap_spreadsheet_id', cleanId);
      localStorage.setItem('rekap_sheet_name', cleanName);
      
      // Reset cached sheet id until loaded
      localStorage.removeItem('rekap_sheet_id');
    }

    setShowConfigModal(false);
    showToast('Konfigurasi berhasil diperbarui!');
  };

  const handleResetConfig = () => {
    setTempSpreadsheetId(DEFAULT_SPREADSHEET_ID);
    setTempSheetName(DEFAULT_SHEET_NAME);
    setSpreadsheetId(DEFAULT_SPREADSHEET_ID);
    setSheetName(DEFAULT_SHEET_NAME);
    setSheetId(0);
    
    setConnectionMode('sheets-api');
    setAppsScriptUrl('');
    setTempConnectionMode('sheets-api');
    setTempAppsScriptUrl('');

    localStorage.removeItem('rekap_spreadsheet_id');
    localStorage.removeItem('rekap_sheet_name');
    localStorage.removeItem('rekap_sheet_id');
    localStorage.removeItem('rekap_connection_mode');
    localStorage.removeItem('rekap_apps_script_url');

    setShowConfigModal(false);
    showToast('Konfigurasi dikembalikan ke default.');
  };

  // Export local copy of state to CSV
  const handleExportLocalCSV = () => {
    try {
      const headers = ['NAMA CUSTOMER', 'NO TELPN', 'ALAMAT', 'HARGA DIPAKAI', 'KETERANGAN', 'PEMBAYARAN'];
      const csvRows = [headers.join(',')];
      
      customers.forEach(c => {
        const row = [
          `"${c.name.replace(/"/g, '""')}"`,
          `"${c.phone.replace(/"/g, '""')}"`,
          `"${c.address.replace(/"/g, '""')}"`,
          `"${c.priceUsed.replace(/"/g, '""')}"`,
          `"${c.keterangan.replace(/"/g, '""')}"`,
          `"${c.pembayaran.replace(/"/g, '""')}"`
        ];
        csvRows.push(row.join(','));
      });

      const blob = new Blob([csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.setAttribute('href', url);
      link.setAttribute('download', `rekap_customer_${new Date().toISOString().split('T')[0]}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast('Berhasil mengunduh salinan CSV!');
    } catch (err) {
      console.error('Export CSV error:', err);
    }
  };

  // Compute analytics / stats
  const stats = useMemo(() => {
    const total = customers.length;
    let dikirim = 0;
    let pending = 0;
    let transfer = 0;
    let cash = 0;

    customers.forEach(c => {
      const ket = c.keterangan.toLowerCase();
      const pem = c.pembayaran.toLowerCase();

      if (ket.includes('kirim') || ket.includes('selesai')) {
        dikirim++;
      } else {
        pending++;
      }

      if (pem.includes('transfer')) {
        transfer++;
      } else if (pem.includes('cash') || pem.includes('tunai') || pem.includes('bayar')) {
        cash++;
      }
    });

    return { total, dikirim, pending, transfer, cash };
  }, [customers]);

  // Compute filtered customer list
  const filteredCustomers = useMemo(() => {
    return customers.filter(c => {
      const matchesSearch = 
        c.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        c.phone.toLowerCase().includes(searchQuery.toLowerCase()) ||
        c.address.toLowerCase().includes(searchQuery.toLowerCase()) ||
        c.keterangan.toLowerCase().includes(searchQuery.toLowerCase()) ||
        c.pembayaran.toLowerCase().includes(searchQuery.toLowerCase());

      const matchesKeterangan = 
        filterKeterangan === 'all' || 
        c.keterangan.toLowerCase().includes(filterKeterangan.toLowerCase());

      const matchesPembayaran = 
        filterPembayaran === 'all' || 
        c.pembayaran.toLowerCase().includes(filterPembayaran.toLowerCase());

      return matchesSearch && matchesKeterangan && matchesPembayaran;
    });
  }, [customers, searchQuery, filterKeterangan, filterPembayaran]);

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col antialiased">
      {/* Dynamic Success Toast */}
      <AnimatePresence>
        {successMessage && (
          <motion.div 
            initial={{ opacity: 0, y: -50, scale: 0.9 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -20, scale: 0.95 }}
            className="fixed top-5 left-1/2 -translate-x-1/2 z-50 bg-slate-900 text-white border border-slate-800 font-medium px-6 py-3.5 rounded-2xl shadow-2xl flex items-center space-x-3"
            id="toast-notification"
          >
            <div className="w-2 h-2 rounded-full bg-blue-400 animate-pulse"></div>
            <span className="text-sm font-semibold">{successMessage}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Global Navigation Header - Sleek Design */}
      <header className="sticky top-0 z-40 bg-white/90 backdrop-blur-md border-b border-slate-100" id="header-nav">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            {/* Logo & Meta */}
            <div className="flex items-center space-x-3.5">
              <div className="w-9 h-9 bg-blue-600 rounded-xl flex items-center justify-center font-bold text-white shadow-lg shadow-blue-100">
                S
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="font-bold text-slate-900 text-lg tracking-tight">SyncRecap Pro</span>
                  <span className={`text-[10px] font-bold px-2.5 py-0.5 rounded-full border uppercase tracking-wider ${
                    connectionMode === 'apps-script'
                      ? 'bg-emerald-50 text-emerald-600 border-emerald-100'
                      : 'bg-blue-50 text-blue-600 border-blue-100'
                  }`}>
                    {connectionMode === 'apps-script' ? 'Apps Script Sync' : 'Google Sheets API'}
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-0.5 flex items-center gap-1.5">
                  <Database className={`w-3.5 h-3.5 ${connectionMode === 'apps-script' ? 'text-emerald-600' : 'text-blue-600'}`} />
                  {connectionMode === 'apps-script' ? (
                    <>
                      <span>Apps Script URL:</span>
                      <span className="font-mono text-slate-600 bg-slate-100 px-1.5 py-0.5 rounded text-[10px] truncate max-w-[140px] sm:max-w-xs inline-block">
                        {appsScriptUrl ? appsScriptUrl : 'Belum diatur'}
                      </span>
                    </>
                  ) : (
                    <>
                      <span>ID Sheet:</span>
                      <span className="font-mono text-slate-600 bg-slate-100 px-1.5 py-0.5 rounded text-[10px] truncate max-w-[140px] sm:max-w-xs inline-block">
                        {spreadsheetId}
                      </span>
                    </>
                  )}
                </p>
              </div>
            </div>

            {/* Quick Actions & Auth */}
            <div className="flex flex-wrap items-center gap-2 sm:gap-3">
              <button
                onClick={() => setShowHelpModal(true)}
                className="inline-flex items-center space-x-1.5 px-4 py-2 border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-sm font-semibold rounded-xl transition-all cursor-pointer"
                title="Petunjuk Integrasi"
                id="btn-help"
              >
                <HelpCircle className="w-4 h-4 text-slate-500" />
                <span className="hidden sm:inline">Petunjuk</span>
              </button>

              <button
                onClick={() => {
                  setTempSpreadsheetId(spreadsheetId);
                  setTempSheetName(sheetName);
                  setTempConnectionMode(connectionMode);
                  setTempAppsScriptUrl(appsScriptUrl);
                  setShowConfigModal(true);
                }}
                className="inline-flex items-center space-x-1.5 px-4 py-2 border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-sm font-semibold rounded-xl transition-all cursor-pointer"
                id="btn-config"
              >
                <Settings className="w-4 h-4 text-slate-500" />
                <span className="hidden sm:inline">Ubah Sheet</span>
              </button>

              {/* User Authentication state */}
              {currentUser ? (
                <div className="flex items-center space-x-2 bg-slate-100 border border-slate-200/60 pl-3 pr-1 py-1 rounded-xl">
                  <div className="flex items-center space-x-1.5">
                    {currentUser.photoURL ? (
                      <img src={currentUser.photoURL} alt={currentUser.displayName} className="w-6 h-6 rounded-full border border-slate-300" referrerPolicy="no-referrer" />
                    ) : (
                      <div className="w-6 h-6 bg-slate-900 text-white rounded-full flex items-center justify-center text-[10px] font-bold">
                        {currentUser.displayName ? currentUser.displayName[0] : 'U'}
                      </div>
                    )}
                    <span className="text-xs font-bold text-slate-700 max-w-[100px] truncate hidden md:inline">
                      {currentUser.displayName || 'User'}
                    </span>
                  </div>
                  <button
                    onClick={handleLogout}
                    className="p-1.5 hover:bg-red-50 text-slate-500 hover:text-red-600 rounded-lg transition-colors cursor-pointer"
                    title="Keluar"
                    id="btn-logout"
                  >
                    <LogOut className="w-4 h-4" />
                  </button>
                </div>
              ) : (
                <button
                  onClick={handleLogin}
                  disabled={isLoggingIn}
                  className="inline-flex items-center space-x-2 px-4 py-2 bg-slate-900 hover:bg-slate-800 disabled:bg-slate-700 text-white text-sm font-bold rounded-xl transition-all shadow-md shadow-slate-100 cursor-pointer"
                  id="btn-google-login"
                >
                  {isLoggingIn ? (
                    <RefreshCw className="w-4 h-4 animate-spin" />
                  ) : (
                    <svg className="w-4 h-4 fill-current" viewBox="0 0 24 24">
                      <path d="M12.24 10.285V13.4h6.887c-.275 1.565-1.88 4.604-6.887 4.604-4.33 0-7.866-3.577-7.866-8s3.536-8 7.866-8c2.46 0 4.105 1.025 5.047 1.926l2.427-2.334C17.955 2.192 15.34 1 12.24 1 6.033 1 12.24s5.033 11.24 11.24 11.24c6.478 0 10.793-4.537 10.793-10.986 0-.74-.08-1.3-.176-1.854H12.24z"/>
                    </svg>
                  )}
                  <span>{isLoggingIn ? 'Menghubungkan...' : 'Login Google'}</span>
                </button>
              )}
            </div>
          </div>
        </div>
      </header>

      {/* Sub-header section exact to design theme layout */}
      <div className="max-w-7xl w-full mx-auto px-6 pt-8 pb-4">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-4">
          <div>
            <h1 className="text-3xl font-bold text-slate-900 tracking-tight">Customer Data Entry</h1>
            <p className="text-slate-500 text-sm mt-1">Real-time synchronization with Google Sheets backend</p>
          </div>
          <div className="flex gap-3">
            <button 
              onClick={handleExportLocalCSV}
              disabled={customers.length === 0}
              className="px-4 py-2 border border-slate-200 hover:bg-slate-50 bg-white rounded-xl text-sm font-semibold text-slate-700 transition-all cursor-pointer"
            >
              Export CSV
            </button>
            <a 
              href={`https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`}
              target="_blank"
              referrerPolicy="no-referrer"
              className="px-4 py-2 bg-blue-600 hover:bg-blue-700 rounded-xl text-sm font-semibold text-white shadow-lg shadow-blue-100 transition-all flex items-center space-x-1.5 cursor-pointer"
            >
              <span>Open Google Sheet</span>
              <ArrowUpRight className="w-4 h-4" />
            </a>
          </div>
        </div>
      </div>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-6 py-4 space-y-6" id="main-content">
        
        {/* Error / Warning Alert Panel */}
        {error && (
          <div className="bg-red-50 border border-red-100 p-4 rounded-3xl shadow-sm flex items-start space-x-3" id="error-alert">
            <AlertCircle className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5" />
            <div className="flex-1">
              <h3 className="text-sm font-bold text-red-800">Perhatian</h3>
              <p className="text-xs text-red-700 mt-0.5">{error}</p>
            </div>
            <button onClick={() => setError(null)} className="text-red-400 hover:text-red-600 p-1 rounded cursor-pointer">
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Read Only banner indicator if not logged in */}
        {!accessToken && connectionMode !== 'apps-script' && (
          <div className="bg-amber-50/70 border border-amber-100 p-5 rounded-3xl shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-start md:items-center space-x-3.5">
              <div className="bg-amber-100 text-amber-800 p-2 rounded-2xl flex-shrink-0">
                <Lock className="w-5 h-5" />
              </div>
              <div>
                <h4 className="text-sm font-bold text-amber-800">Mode Lihat Saja (Read-Only)</h4>
                <p className="text-xs text-amber-700 mt-0.5 leading-relaxed">
                  Anda sedang melihat cache CSV publik. Untuk menginput, memperbarui, atau menghapus data customer, harap masuk dengan Akun Google Anda.
                </p>
              </div>
            </div>
            <button 
              onClick={handleLogin}
              className="inline-flex items-center space-x-1.5 px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold rounded-xl shadow transition-all cursor-pointer self-start md:self-auto"
            >
              <span>Sambungkan Google</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* Bento Dashboard Metrics Cards - Sleek rounded-3xl and style */}
        <section className="grid grid-cols-1 md:grid-cols-3 gap-4" id="dashboard-metrics">
          <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col justify-between">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Total Customer</span>
            <div className="flex items-baseline space-x-2 mt-3">
              <span className="text-3xl font-bold text-slate-900 tracking-tight">{loading ? '...' : stats.total}</span>
              <span className="text-xs text-slate-400">baris</span>
            </div>
            <div className="w-full bg-slate-100 h-1.5 rounded-full mt-4 overflow-hidden">
              <div className="bg-blue-600 h-1.5 rounded-full" style={{ width: '100%' }}></div>
            </div>
          </div>
        </section>

        {/* Workspace Dual Layout */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          
          {/* LEFT COLUMN: Data Entry Form */}
          <section className="lg:col-span-5 space-y-6" id="form-section">
            <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-6">
              
              {/* Form header title exact to design HTML */}
              <h3 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-6">
                {editingCustomer ? 'Perbarui Record' : 'New Record'}
              </h3>

              {editingCustomer && (
                <div className="mb-5 p-3.5 bg-amber-50/50 border border-amber-100 rounded-xl text-xs text-amber-800 font-semibold flex justify-between items-center">
                  <span>Mengedit baris data ke-{editingCustomer.id}</span>
                  <button type="button" onClick={resetForm} className="text-slate-400 hover:text-slate-600">
                    <X className="w-4 h-4" />
                  </button>
                </div>
              )}

              {/* Form body */}
              <form onSubmit={handleSubmit} className="space-y-4">
                {/* Customer Name */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-600">
                    Nama Customer <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    value={formName}
                    onChange={(e) => setFormName(e.target.value)}
                    placeholder="e.g. John Doe"
                    disabled={!accessToken && connectionMode !== 'apps-script'}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                  />
                </div>

                {/* Telephone */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-600">
                    No Telepon
                  </label>
                  <input
                    type="text"
                    value={formPhone}
                    onChange={(e) => {
                      let val = e.target.value;
                      if (val.startsWith('+62')) {
                        val = '0' + val.slice(3);
                      } else if (val.startsWith('62') && val.length > 2) {
                        val = '0' + val.slice(2);
                      }
                      setFormPhone(val);
                    }}
                    placeholder="e.g. 0812..."
                    disabled={!accessToken && connectionMode !== 'apps-script'}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                  />
                </div>

                {/* Address */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-600">
                    Alamat Lengkap
                  </label>
                  <textarea
                    rows={2}
                    value={formAddress}
                    onChange={(e) => setFormAddress(e.target.value)}
                    placeholder="e.g. Jl. Mahendradatta Selatan No. 2, Bali"
                    disabled={!accessToken && connectionMode !== 'apps-script'}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm outline-none focus:bg-white focus:border-blue-400 resize-none transition-all"
                  />
                </div>

                {/* Harga Dipakai (Price Type) */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-600">
                    Harga Dipakai
                  </label>
                  <input
                    type="text"
                    value={formPriceUsed}
                    onChange={(e) => setFormPriceUsed(e.target.value)}
                    placeholder="e.g. harga 2"
                    disabled={!accessToken && connectionMode !== 'apps-script'}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                  />
                  {/* Preset helpers */}
                  <div className="flex flex-wrap gap-1 mt-1.5">
                    {pricePresets.map((p) => (
                      <button
                        key={p}
                        type="button"
                        disabled={!accessToken && connectionMode !== 'apps-script'}
                        onClick={() => setFormPriceUsed(p)}
                        className="text-[10px] bg-slate-100 hover:bg-slate-200 disabled:opacity-50 text-slate-600 px-2 py-1 rounded-md transition-colors cursor-pointer font-semibold"
                      >
                        {p}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Keterangan */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-600">
                    Keterangan Pengiriman
                  </label>
                  <input
                    type="text"
                    value={formKeterangan}
                    onChange={(e) => setFormKeterangan(e.target.value)}
                    placeholder="e.g. dikirim"
                    disabled={!accessToken && connectionMode !== 'apps-script'}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                  />
                  {/* Preset helpers */}
                  <div className="flex flex-wrap gap-1 mt-1.5">
                    {statusPresets.map((s) => (
                      <button
                        key={s}
                        type="button"
                        disabled={!accessToken && connectionMode !== 'apps-script'}
                        onClick={() => setFormKeterangan(s)}
                        className="text-[10px] bg-slate-100 hover:bg-slate-200 disabled:opacity-50 text-slate-600 px-2 py-1 rounded-md transition-colors cursor-pointer font-semibold"
                      >
                        {s}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Pembayaran */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-600">
                    Metode Pembayaran
                  </label>
                  <input
                    type="text"
                    value={formPembayaran}
                    onChange={(e) => setFormPembayaran(e.target.value)}
                    placeholder="e.g. transfer"
                    disabled={!accessToken && connectionMode !== 'apps-script'}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                  />
                  {/* Preset helpers */}
                  <div className="flex flex-wrap gap-1 mt-1.5">
                    {paymentPresets.map((pay) => (
                      <button
                        key={pay}
                        type="button"
                        disabled={!accessToken && connectionMode !== 'apps-script'}
                        onClick={() => setFormPembayaran(pay)}
                        className="text-[10px] bg-slate-100 hover:bg-slate-200 disabled:opacity-50 text-slate-600 px-2 py-1 rounded-md transition-colors cursor-pointer font-semibold"
                      >
                        {pay}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Submit Actions */}
                <button
                  type="submit"
                  disabled={saving || (!accessToken && connectionMode !== 'apps-script')}
                  className="w-full py-3 mt-4 bg-slate-900 text-white rounded-xl font-bold text-sm hover:bg-slate-800 active:scale-[0.98] transition-all disabled:bg-slate-300 disabled:cursor-not-allowed cursor-pointer shadow-lg shadow-slate-100 flex items-center justify-center space-x-2"
                >
                  {saving ? (
                    <>
                      <RefreshCw className="w-4 h-4 animate-spin" />
                      <span>Menyimpan ke Database...</span>
                    </>
                  ) : (
                    <span>{editingCustomer ? 'Perbarui Record di Sheet' : 'Simpan ke Database'}</span>
                  )}
                </button>
              </form>
            </div>
            
            {/* Cloud Status Panel - Matches the premium Dark Sidebar widget exactly */}
            <div className="bg-slate-900 text-white rounded-3xl p-6 shadow-xl">
              <p className="text-xs text-slate-400 mb-1 font-bold uppercase tracking-wider">Cloud Status</p>
              {connectionMode === 'apps-script' ? (
                <>
                  <div className="flex items-center gap-2">
                    <div className={`w-2 h-2 rounded-full ${appsScriptUrl ? 'bg-green-400 animate-pulse' : 'bg-amber-400'}`}></div>
                    <p className="text-sm font-semibold">{appsScriptUrl ? 'Connected (Apps Script)' : 'Belum Terkonfigurasi'}</p>
                  </div>
                  <p className="text-[10px] text-slate-500 mt-2 break-all font-mono">
                    {appsScriptUrl ? appsScriptUrl : 'Masukkan URL web app Apps Script Anda di menu Pengaturan.'}
                  </p>
                </>
              ) : (
                <>
                  <div className="flex items-center gap-2">
                    <div className={`w-2 h-2 rounded-full ${accessToken ? 'bg-green-400 animate-pulse' : 'bg-amber-400'}`}></div>
                    <p className="text-sm font-semibold">{accessToken ? 'Connected to Sheet' : 'Read-Only Mode'}</p>
                  </div>
                  <p className="text-[10px] text-slate-500 mt-2 break-all font-mono">
                    docs.google.com/sheets/d/{spreadsheetId}
                  </p>
                </>
              )}
              {connectionMode !== 'apps-script' && (
                <a
                  href={`https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`}
                  target="_blank"
                  referrerPolicy="no-referrer"
                  className="mt-4 w-full py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white text-xs font-bold rounded-xl text-center block transition-all"
                >
                  Open Google Sheets Workspace &rarr;
                </a>
              )}
            </div>
          </section>

          {/* RIGHT COLUMN: Interactive Customer Table & Filter System */}
          <section className="lg:col-span-7 space-y-6" id="table-section">
            <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-6 flex flex-col">
              
              {/* Header Title section matching 'Live Sheet Data' of design HTML */}
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-sm font-bold text-slate-400 uppercase tracking-widest">Live Sheet Data</h3>
                <span className="text-[10px] bg-slate-100 text-slate-500 px-2.5 py-0.5 rounded-full font-semibold">
                  Last synced 2m ago
                </span>
              </div>

              {/* Integrated Search Bar & Dropdown Filters - Modern layout */}
              <div className="space-y-4 mb-6">
                <div className="grid grid-cols-1 md:grid-cols-12 gap-3">
                  {/* Search Input */}
                  <div className="md:col-span-6 relative">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      placeholder="Cari customer, no telp, alamat..."
                      className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                    />
                  </div>

                  {/* Filter Keterangan */}
                  <div className="md:col-span-3">
                    <select
                      value={filterKeterangan}
                      onChange={(e) => setFilterKeterangan(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-xs font-semibold text-slate-600 outline-none cursor-pointer focus:bg-white focus:border-blue-400 transition-all appearance-none"
                    >
                      <option value="all">Status: Semua</option>
                      <option value="dikirim">Status: Dikirim</option>
                      <option value="pending">Status: Pending</option>
                      <option value="selesai">Status: Selesai</option>
                      <option value="batal">Status: Batal</option>
                    </select>
                  </div>

                  {/* Filter Pembayaran */}
                  <div className="md:col-span-3">
                    <select
                      value={filterPembayaran}
                      onChange={(e) => setFilterPembayaran(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-xs font-semibold text-slate-600 outline-none cursor-pointer focus:bg-white focus:border-blue-400 transition-all appearance-none"
                    >
                      <option value="all">Bayar: Semua</option>
                      <option value="transfer">Bayar: Transfer</option>
                      <option value="cash">Bayar: Cash</option>
                      <option value="pending">Bayar: Pending</option>
                    </select>
                  </div>
                </div>
              </div>

              {/* Data Table */}
              <div className="overflow-x-auto">
                {loading ? (
                  <div className="py-20 text-center space-y-3">
                    <RefreshCw className="w-8 h-8 animate-spin text-blue-600 mx-auto" />
                    <p className="text-sm text-slate-500 font-semibold">Sedang mengambil data terbaru...</p>
                  </div>
                ) : filteredCustomers.length === 0 ? (
                  <div className="py-16 text-center px-4 bg-slate-50/50 rounded-2xl border border-dashed border-slate-200">
                    <Search className="w-8 h-8 text-slate-300 mx-auto mb-3" />
                    <h3 className="text-sm font-bold text-slate-700">Tidak Ada Customer</h3>
                    <p className="text-xs text-slate-500 mt-1 max-w-sm mx-auto">
                      Cari dengan kata kunci lain atau submit record baru.
                    </p>
                  </div>
                ) : (
                  <table className="w-full text-left">
                    <thead>
                      <tr className="border-b border-slate-100">
                        <th className="pb-4 text-xs font-bold text-slate-400 uppercase tracking-wider w-[25%]">Customer</th>
                        <th className="pb-4 text-xs font-bold text-slate-400 uppercase tracking-wider w-[40%]">Alamat</th>
                        <th className="pb-4 text-xs font-bold text-slate-400 uppercase tracking-wider w-[15%]">Harga</th>
                        <th className="pb-4 text-xs font-bold text-slate-400 uppercase tracking-wider text-right w-[20%]">Mode Pembayaran</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {filteredCustomers.map((customer, idx) => {
                        const isKeteranganSelesai = customer.keterangan.toLowerCase().includes('kirim') || customer.keterangan.toLowerCase().includes('selesai');
                        const isPembayaranOk = customer.pembayaran.toLowerCase().includes('transfer') || customer.pembayaran.toLowerCase().includes('cash') || customer.pembayaran.toLowerCase().includes('tunai');

                        return (
                          <tr key={customer.id} className="group hover:bg-slate-50/50 transition-colors">
                            <td className="py-4 pr-3">
                              <div className="flex flex-col space-y-1 text-sm">
                                <span className="font-bold text-slate-900 break-words">{customer.name}</span>
                                {customer.phone ? (
                                  <div className="whitespace-nowrap">
                                    <a 
                                      href={`tel:${customer.phone}`}
                                      className="inline-flex items-center text-xs text-blue-600 hover:text-blue-800 font-bold hover:underline bg-blue-50/70 px-2 py-0.5 rounded-full font-mono whitespace-nowrap"
                                    >
                                      {customer.phone}
                                    </a>
                                  </div>
                                ) : (
                                  <span className="text-xs text-slate-300 italic">-</span>
                                )}
                              </div>
                            </td>
                            
                            <td className="py-4 pr-3 text-xs text-slate-600 break-words">
                              {customer.address ? (
                                <div className="space-y-1">
                                  <p className="whitespace-normal break-words" title={customer.address}>{customer.address}</p>
                                  <a
                                    href={`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(customer.address)}`}
                                    target="_blank"
                                    referrerPolicy="no-referrer"
                                    className="inline-flex items-center space-x-0.5 text-blue-600 hover:text-blue-800 font-bold"
                                  >
                                    <MapPin className="w-3 h-3" />
                                    <span>Google Maps</span>
                                  </a>
                                </div>
                              ) : (
                                <span className="text-slate-400 italic">No address</span>
                              )}
                            </td>

                            <td className="py-4 pr-3 whitespace-nowrap">
                              <span className="text-[11px] bg-slate-100 text-slate-700 font-semibold px-2 py-0.5 rounded whitespace-nowrap">
                                {customer.priceUsed || 'default'}
                              </span>
                            </td>

                            <td className="py-4 text-right">
                              <div className="flex flex-col items-end space-y-1.5">
                                <div className="flex items-center gap-1.5">
                                  {customer.keterangan && (
                                    <span className={`px-2 py-0.5 text-[9px] font-bold rounded-md uppercase tracking-wider ${
                                      isKeteranganSelesai 
                                        ? 'bg-green-50 text-green-600' 
                                        : 'bg-amber-50 text-amber-600'
                                    }`}>
                                      {customer.keterangan}
                                    </span>
                                  )}
                                  {customer.pembayaran && (
                                    <span className={`px-2 py-0.5 text-[9px] font-bold rounded-md uppercase tracking-wider ${
                                      isPembayaranOk
                                        ? 'bg-blue-50 text-blue-600'
                                        : 'bg-rose-50 text-rose-600'
                                    }`}>
                                      {customer.pembayaran}
                                    </span>
                                  )}
                                </div>

                                {/* Hover visible Action Controls */}
                                <div className="flex items-center justify-end space-x-1 mt-1 opacity-80 group-hover:opacity-100 transition-opacity">
                                  <button
                                    onClick={() => handleStartPrint(customer)}
                                    className="p-1 text-slate-500 hover:text-emerald-600 hover:bg-emerald-50 rounded transition-colors cursor-pointer"
                                    title="Print to Thermal Printer"
                                    id={`btn-print-${customer.id}`}
                                  >
                                    <Printer className="w-3.5 h-3.5" />
                                  </button>
                                  {(accessToken || connectionMode === 'apps-script') && (
                                    <>
                                      <button
                                        onClick={() => handleStartEdit(customer)}
                                        className="p-1 text-slate-500 hover:text-blue-600 hover:bg-slate-100 rounded transition-colors cursor-pointer"
                                        title="Edit record"
                                        id={`btn-edit-${customer.id}`}
                                      >
                                        <Edit2 className="w-3.5 h-3.5" />
                                      </button>
                                      <button
                                        onClick={() => handleDelete(customer)}
                                        className="p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors cursor-pointer"
                                        title="Delete record"
                                        id={`btn-delete-${customer.id}`}
                                      >
                                        <Trash2 className="w-3.5 h-3.5" />
                                      </button>
                                    </>
                                  )}
                                </div>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>
              
              {/* Table Footer with custom layout action */}
              <div className="mt-6 pt-6 border-t border-slate-100 flex justify-center">
                <a 
                  href={`https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`}
                  target="_blank"
                  referrerPolicy="no-referrer"
                  className="text-xs font-bold text-blue-600 hover:text-blue-800 flex items-center space-x-1 cursor-pointer"
                >
                  <span>View All Records In Sheet</span>
                  <span>&rarr;</span>
                </a>
              </div>
            </div>
          </section>

        </div>
      </main>

      {/* FOOTER */}
      <footer className="bg-slate-900 text-slate-400 border-t border-slate-800 py-8 mt-16" id="global-footer">
        <div className="max-w-7xl mx-auto px-6 flex flex-col md:flex-row items-center justify-between gap-4">
          <p className="text-xs">
            © 2026 SyncRecap Pro. Fully integrated with Google Sheets API in real-time.
          </p>
          <div className="flex items-center space-x-4 text-xs font-bold">
            <button onClick={() => setShowHelpModal(true)} className="hover:text-white transition-colors cursor-pointer">Petunjuk</button>
            <span className="text-slate-700">|</span>
            <button 
              onClick={() => {
                setTempSpreadsheetId(spreadsheetId);
                setTempSheetName(sheetName);
                setTempConnectionMode(connectionMode);
                setTempAppsScriptUrl(appsScriptUrl);
                setShowConfigModal(true);
              }} 
              className="hover:text-white transition-colors cursor-pointer"
            >
              Ubah Sheet Target
            </button>
          </div>
        </div>
      </footer>

      {/* MODAL 1: SPREADSHEET CONFIGURATION MODAL */}
      <AnimatePresence>
        {showConfigModal && (
          <div className="fixed inset-0 z-50 overflow-y-auto" id="config-modal">
            {/* Backdrop */}
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 0.5 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowConfigModal(false)}
              className="fixed inset-0 bg-black"
            ></motion.div>

            {/* Content Container */}
            <div className="flex min-h-screen items-center justify-center p-4 relative">
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 15 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 15 }}
                className="bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden relative border border-slate-100"
              >
                {/* Header */}
                <div className="bg-slate-900 text-white p-6 flex items-center justify-between">
                  <div className="flex items-center space-x-2.5">
                    <Settings className="w-5 h-5 text-blue-400" />
                    <h3 className="font-bold text-base">Konfigurasi Google Sheets</h3>
                  </div>
                  <button 
                    onClick={() => setShowConfigModal(false)}
                    className="text-slate-400 hover:text-white p-1 rounded-lg transition-colors cursor-pointer"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>

                {/* Form fields */}
                <div className="p-6 space-y-5 max-h-[65vh] overflow-y-auto">
                  {/* Connection Mode Selection */}
                  <div className="space-y-1.5">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                      Metode Sinkronisasi Data
                    </label>
                    <div className="grid grid-cols-2 gap-2 bg-slate-100 p-1 rounded-xl">
                      <button
                        type="button"
                        onClick={() => setTempConnectionMode('sheets-api')}
                        className={`py-2 px-3 text-xs font-bold rounded-lg transition-all cursor-pointer text-center ${
                          tempConnectionMode === 'sheets-api'
                            ? 'bg-white text-slate-900 shadow-sm'
                            : 'text-slate-500 hover:text-slate-900'
                        }`}
                      >
                        Google Sheets API
                      </button>
                      <button
                        type="button"
                        onClick={() => setTempConnectionMode('apps-script')}
                        className={`py-2 px-3 text-xs font-bold rounded-lg transition-all cursor-pointer text-center ${
                          tempConnectionMode === 'apps-script'
                            ? 'bg-emerald-600 text-white shadow-sm'
                            : 'text-slate-500 hover:text-slate-900'
                        }`}
                      >
                        Google Apps Script (Mudah!)
                      </button>
                    </div>
                  </div>

                  {tempConnectionMode === 'apps-script' ? (
                    <div className="space-y-5">
                      {/* Apps Script URL Input */}
                      <div className="space-y-1.5">
                        <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                          Google Apps Script Web App URL
                        </label>
                        <input
                          type="text"
                          value={tempAppsScriptUrl}
                          onChange={(e) => setTempAppsScriptUrl(e.target.value)}
                          placeholder="https://script.google.com/macros/s/.../exec"
                          className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:bg-white focus:border-emerald-500 transition-all"
                        />
                      </div>

                      {/* Setup Instructions */}
                      <div className="bg-slate-50 p-4 rounded-2xl text-xs text-slate-600 leading-relaxed border border-slate-100 space-y-3">
                        <p className="font-bold text-slate-800 flex items-center gap-1.5">
                          <Info className="w-4 h-4 text-emerald-600" />
                          Petunjuk Deploy Apps Script:
                        </p>
                        <ol className="list-decimal pl-4 space-y-1.5 font-medium">
                          <li>Buka dokumen Google Sheets Anda.</li>
                          <li>Di menu atas, pilih <strong>Extensions (Ekstensi) &rarr; Apps Script</strong>.</li>
                          <li>Hapus semua kode bawaan, lalu salin dan tempel kode template di bawah ini.</li>
                          <li>Klik tombol <strong>Save</strong> (ikon disket), lalu klik <strong>Deploy &rarr; New deployment</strong>.</li>
                          <li>Pilih tipe <strong>Web app</strong>. Konfigurasi:
                            <ul className="list-disc pl-4 mt-1 space-y-1">
                              <li>Execute as: <strong>Me (your-email)</strong></li>
                              <li>Who has access: <strong>Anyone</strong> (Penting agar aplikasi ini bisa mengirim data).</li>
                            </ul>
                          </li>
                          <li>Klik <strong>Deploy</strong>, izinkan akses jika diminta, lalu salin <strong>Web App URL</strong>-nya dan tempel ke kolom di atas.</li>
                        </ol>

                        {/* Copy Code Box */}
                        <div className="mt-3 border border-slate-200 rounded-xl bg-slate-900 text-slate-100 p-3.5 space-y-2">
                          <div className="flex items-center justify-between">
                            <span className="font-mono text-[10px] text-slate-400 flex items-center gap-1">
                              <Code className="w-3.5 h-3.5" /> Code.gs Template
                            </span>
                            <button
                              type="button"
                              onClick={() => {
                                const scriptText = `function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    if (!row['id']) {
      row['id'] = i + 1;
    }
    rows.push(row);
  }
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  
  if (action === "create") {
    var lastRow = sheet.getLastRow();
    var newId = lastRow + 1;
    sheet.appendRow([
      newId,
      params.name || "",
      params.phone || "",
      params.address || "",
      params.priceUsed || "",
      params.keterangan || "",
      params.pembayaran || ""
    ]);
    return ContentService.createTextOutput(JSON.stringify({ status: "success", id: newId }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "update") {
    var targetId = parseInt(params.id);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === targetId) {
        var rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(params.name || "");
        sheet.getRange(rowNum, 3).setValue(params.phone || "");
        sheet.getRange(rowNum, 4).setValue(params.address || "");
        sheet.getRange(rowNum, 5).setValue(params.priceUsed || "");
        sheet.getRange(rowNum, 6).setValue(params.keterangan || "");
        sheet.getRange(rowNum, 7).setValue(params.pembayaran || "");
        return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ID not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "delete") {
    var targetId = parseInt(params.id);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === targetId) {
        sheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ID not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action" }))
    .setMimeType(ContentService.MimeType.JSON);
}`;
                                navigator.clipboard.writeText(scriptText);
                                setCopiedScript(true);
                                setTimeout(() => setCopiedScript(false), 2000);
                              }}
                              className="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-[10px] rounded border border-slate-700 flex items-center gap-1 cursor-pointer transition-colors"
                            >
                              {copiedScript ? (
                                <>
                                  <Check className="w-3 h-3 text-green-400" /> Copied!
                                </>
                              ) : (
                                <>
                                  <Copy className="w-3 h-3 text-slate-300" /> Copy Code
                                </>
                              )}
                            </button>
                          </div>
                          <pre className="font-mono text-[9px] text-slate-300 max-h-40 overflow-y-auto bg-black/40 p-2 rounded border border-slate-800">
{`function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    if (!row['id']) {
      row['id'] = i + 1;
    }
    rows.push(row);
  }
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  
  if (action === "create") {
    var lastRow = sheet.getLastRow();
    var newId = lastRow + 1;
    sheet.appendRow([
      newId,
      params.name || "",
      params.phone || "",
      params.address || "",
      params.priceUsed || "",
      params.keterangan || "",
      params.pembayaran || ""
    ]);
    return ContentService.createTextOutput(JSON.stringify({ status: "success", id: newId }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "update") {
    var targetId = parseInt(params.id);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === targetId) {
        var rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(params.name || "");
        sheet.getRange(rowNum, 3).setValue(params.phone || "");
        sheet.getRange(rowNum, 4).setValue(params.address || "");
        sheet.getRange(rowNum, 5).setValue(params.priceUsed || "");
        sheet.getRange(rowNum, 6).setValue(params.keterangan || "");
        sheet.getRange(rowNum, 7).setValue(params.pembayaran || "");
        return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ID not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "delete") {
    var targetId = parseInt(params.id);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === targetId) {
        sheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ID not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action" }))
    .setMimeType(ContentService.MimeType.JSON);
}`}
                          </pre>
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-5">
                      <div className="bg-slate-50 p-4 rounded-2xl text-xs text-slate-600 leading-relaxed border border-slate-100">
                        <p className="font-bold text-slate-800 flex items-center gap-1.5 mb-1">
                          <Info className="w-4 h-4 text-blue-600" />
                          Cara Menghubungkan Google Sheet Kustom Anda:
                        </p>
                        <ol className="list-decimal pl-4 space-y-1.5 mt-2 font-medium">
                          <li>Buka dokumen spreadsheet Anda di browser Google Sheets.</li>
                          <li>Salin ID unik spreadsheet yang ada di URL browser Anda (antara <code className="bg-slate-200/80 px-1 rounded font-mono">/d/</code> dan <code className="bg-slate-200/80 px-1 rounded font-mono">/edit</code>).</li>
                          <li>Tempelkan ke kolom <strong>Spreadsheet ID</strong> di bawah ini.</li>
                          <li>Pastikan Akun Google Anda memiliki izin akses membaca & menulis.</li>
                        </ol>
                      </div>

                      {/* Spreadsheet ID Input */}
                      <div className="space-y-1.5">
                        <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                          Spreadsheet ID
                        </label>
                        <input
                          type="text"
                          value={tempSpreadsheetId}
                          onChange={(e) => setTempSpreadsheetId(e.target.value)}
                          placeholder="Masukkan Spreadsheet ID Anda"
                          className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                        />
                      </div>

                      {/* Sheet Name Input / Selection */}
                      <div className="space-y-1.5">
                        <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                          Nama Sheet (Tab Nama)
                        </label>
                        {availableSheets.length > 0 ? (
                          <div className="grid grid-cols-2 gap-2 mt-1">
                            {availableSheets.map((s) => (
                              <button
                                key={s.sheetId}
                                onClick={() => setTempSheetName(s.title)}
                                className={`px-3 py-2.5 text-xs font-bold rounded-xl border text-left flex items-center justify-between transition-all cursor-pointer ${
                                  tempSheetName.toLowerCase() === s.title.toLowerCase()
                                    ? 'bg-blue-50 text-blue-700 border-blue-200'
                                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
                                }`}
                              >
                                <span className="truncate">{s.title}</span>
                                {tempSheetName.toLowerCase() === s.title.toLowerCase() && (
                                  <Check className="w-3.5 h-3.5 text-blue-600 flex-shrink-0" />
                                )}
                              </button>
                            ))}
                          </div>
                        ) : (
                          <input
                            type="text"
                            value={tempSheetName}
                            onChange={(e) => setTempSheetName(e.target.value)}
                            placeholder="Contoh: Sheet1"
                            className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:bg-white focus:border-blue-400 transition-all"
                          />
                        )}
                      </div>
                    </div>
                  )}
                </div>

                {/* Footer Buttons */}
                <div className="p-5 bg-slate-50 border-t border-slate-100 flex items-center justify-between">
                  <button
                    onClick={handleResetConfig}
                    className="px-4 py-2 text-xs font-bold text-rose-600 hover:bg-rose-50 rounded-xl transition-all cursor-pointer"
                  >
                    Kembalikan Default
                  </button>
                  <div className="flex items-center space-x-2">
                    <button
                      onClick={() => setShowConfigModal(false)}
                      className="px-4 py-2 text-xs font-bold text-slate-600 hover:bg-slate-100 rounded-xl transition-all cursor-pointer"
                    >
                      Batal
                    </button>
                    <button
                      onClick={handleSaveConfig}
                      className="px-4 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-md transition-all cursor-pointer"
                    >
                      Simpan Konfigurasi
                    </button>
                  </div>
                </div>
              </motion.div>
            </div>
          </div>
        )}
      </AnimatePresence>

      {/* MODAL 2: DETAILED INSTRUCTION MODAL */}
      <AnimatePresence>
        {showHelpModal && (
          <div className="fixed inset-0 z-50 overflow-y-auto" id="help-modal">
            {/* Backdrop */}
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 0.5 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowHelpModal(false)}
              className="fixed inset-0 bg-black"
            ></motion.div>

            {/* Content Container */}
            <div className="flex min-h-screen items-center justify-center p-4 relative">
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 15 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 15 }}
                className="bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden relative border border-slate-100"
              >
                {/* Header */}
                <div className="bg-slate-900 text-white p-6 flex items-center justify-between">
                  <div className="flex items-center space-x-2.5">
                    <HelpCircle className="w-5 h-5 text-blue-400" />
                    <h3 className="font-bold text-base">Panduan Penggunaan</h3>
                  </div>
                  <button 
                    onClick={() => setShowHelpModal(false)}
                    className="text-slate-400 hover:text-white p-1 rounded-lg transition-colors cursor-pointer"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>

                {/* Body Content */}
                <div className="p-6 space-y-4 text-sm text-slate-600 leading-relaxed max-h-[70vh] overflow-y-auto">
                  <div className="space-y-2">
                    <h4 className="font-bold text-slate-800 text-sm">1. Kolom Data di Google Sheets</h4>
                    <p className="text-xs leading-relaxed">
                      Spreadsheet wajib memiliki struktur header persis seperti di bawah ini agar dapat terintegrasi tanpa ada error kolom:
                    </p>
                    <div className="bg-slate-50 p-3.5 border border-slate-200 rounded-xl font-mono text-[10px] text-slate-700 overflow-x-auto">
                      NAMA CUSTOMER, NO TELPN, ALAMAT, HARGA DIPAKAI, KETERANGAN, PEMBAYARAN
                    </div>
                  </div>

                  <div className="space-y-2 pt-2">
                    <h4 className="font-bold text-slate-800 text-sm">2. Menambah Data</h4>
                    <p className="text-xs leading-relaxed">
                      Gunakan formulir sebelah kiri untuk menambahkan data baru. Klik tombol "Submit to Database" untuk menulis langsung ke Google Sheet Anda secara realtime. Akun Anda harus sudah login Google Sheets terlebih dahulu.
                    </p>
                  </div>

                  <div className="space-y-2 pt-2">
                    <h4 className="font-bold text-slate-800 text-sm">3. Mengedit & Menghapus</h4>
                    <p className="text-xs leading-relaxed">
                      Klik tombol edit (<Edit2 className="w-3 h-3 inline text-blue-500" />) di baris tabel untuk memuat data ke formulir kiri lalu perbarui. Klik tombol hapus (<Trash2 className="w-3 h-3 inline text-red-500" />) untuk menghapus baris tersebut secara langsung dari Google Sheet Anda secara permanen.
                    </p>
                  </div>

                  <div className="space-y-2 pt-2">
                    <h4 className="font-bold text-slate-800 text-sm">4. Google Maps & Kontak Cepat</h4>
                    <p className="text-xs leading-relaxed">
                      Aplikasi ini mendeteksi alamat customer dan menampilkan tautan langsung ke Google Maps. Anda juga dapat langsung menelepon customer dengan mengeklik kontak telepon yang terdaftar.
                    </p>
                  </div>

                  <div className="space-y-2 pt-2">
                    <h4 className="font-bold text-slate-800 text-sm">5. Ekspor Data Lokal</h4>
                    <p className="text-xs leading-relaxed">
                      Kapan pun dibutuhkan, Anda dapat mengunduh salinan data lokal yang sedang ditampilkan dalam bentuk file CSV siap cetak dengan tombol "Export CSV".
                    </p>
                  </div>
                </div>

                {/* Footer Buttons */}
                <div className="p-5 bg-slate-50 border-t border-slate-100 text-right">
                  <button
                    onClick={() => setShowHelpModal(false)}
                    className="px-6 py-2.5 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-md transition-all cursor-pointer"
                  >
                    Saya Mengerti
                  </button>
                </div>
              </motion.div>
            </div>
          </div>
        )}
      </AnimatePresence>

      {/* MODAL 3: THERMAL PRINTER MODAL */}
      <AnimatePresence>
        {showPrintModal && printingCustomer && (
          <div className="fixed inset-0 z-50 overflow-y-auto" id="print-modal">
            {/* Backdrop */}
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 0.5 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowPrintModal(false)}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm"
            ></motion.div>

            {/* Content Container */}
            <div className="flex min-h-screen items-center justify-center p-4 relative">
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 15 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 15 }}
                className="bg-white w-full max-w-4xl rounded-3xl shadow-2xl overflow-hidden relative border border-slate-100 flex flex-col md:flex-row text-slate-800"
              >
                {/* Left Column: Configuration & Controls */}
                <div className="flex-1 p-6 space-y-5 flex flex-col justify-between max-h-[90vh] overflow-y-auto">
                  <div className="space-y-5">
                    {/* Header */}
                    <div className="flex items-center justify-between pb-4 border-b border-slate-100">
                      <div className="flex items-center space-x-2.5">
                        <Printer className="w-5 h-5 text-emerald-600" />
                        <div>
                          <h3 className="font-bold text-base text-slate-800">Cetak Thermal Printer</h3>
                          <p className="text-[11px] text-slate-400">Atur font, ukuran teks, cetak tebal, dan susunan rincian</p>
                        </div>
                      </div>
                      <button 
                        onClick={() => setShowPrintModal(false)}
                        className="text-slate-400 hover:text-slate-600 p-1.5 rounded-lg transition-colors cursor-pointer"
                      >
                        <X className="w-5 h-5" />
                      </button>
                    </div>

                    {/* Quick Config Presets */}
                    <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200/60 space-y-3">
                      <div className="flex items-center justify-between">
                        <p className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Opsi Cepat Penerima</p>
                        <button
                          type="button"
                          onClick={() => {
                            const defaultLines = generateDefaultPrintLines(printingCustomer, true, true, true);
                            setPrintLines(defaultLines);
                            showToast('Rincian berhasil direset!');
                          }}
                          className="text-[11px] text-emerald-600 hover:text-emerald-800 font-bold hover:underline"
                        >
                          Reset Default
                        </button>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                        <label className="flex items-center space-x-2 text-xs font-semibold text-slate-700 cursor-pointer select-none">
                          <input 
                            type="checkbox" 
                            checked={optDibantuSiapkan} 
                            onChange={(e) => setOptDibantuSiapkan(e.target.checked)}
                            className="w-4 h-4 text-emerald-600 border-slate-300 rounded focus:ring-emerald-500 cursor-pointer"
                          />
                          <span>Siapkan</span>
                        </label>

                        <label className="flex items-center space-x-2 text-xs font-semibold text-slate-700 cursor-pointer select-none">
                          <input 
                            type="checkbox" 
                            checked={optKeterangan} 
                            onChange={(e) => setOptKeterangan(e.target.checked)}
                            className="w-4 h-4 text-emerald-600 border-slate-300 rounded focus:ring-emerald-500 cursor-pointer"
                          />
                          <span>Keterangan</span>
                        </label>

                        <label className="flex items-center space-x-2 text-xs font-semibold text-slate-700 cursor-pointer select-none">
                          <input 
                            type="checkbox" 
                            checked={optPembayaran} 
                            onChange={(e) => setOptPembayaran(e.target.checked)}
                            className="w-4 h-4 text-emerald-600 border-slate-300 rounded focus:ring-emerald-500 cursor-pointer"
                          />
                          <span>Pembayaran</span>
                        </label>
                      </div>
                    </div>

                    {/* WORD-STYLE FORMATTING RIBBON */}
                    <div className="bg-slate-100/90 backdrop-blur-sm border border-slate-200 rounded-2xl p-3 space-y-3 text-slate-700 shadow-sm">
                      {/* Section 1: Font & Base Size selectors */}
                      <div className="flex flex-wrap items-center gap-2.5 pb-2.5 border-b border-slate-200">
                        {/* Font Family Selector */}
                        <div className="flex items-center space-x-1.5 bg-white px-2.5 py-1.5 rounded-lg border border-slate-200 shadow-sm">
                          <Type className="w-3.5 h-3.5 text-slate-400" />
                          <select
                            value={printFontFamily}
                            onChange={(e) => setPrintFontFamily(e.target.value as any)}
                            className="bg-transparent text-xs font-bold text-slate-800 outline-none cursor-pointer pr-1"
                          >
                            <option value="mono">Courier Mono</option>
                            <option value="sans">Inter Sans</option>
                            <option value="serif">Georgia Serif</option>
                          </select>
                        </div>

                        {/* Base Size Selector */}
                        <div className="flex items-center space-x-1.5 bg-white px-2.5 py-1.5 rounded-lg border border-slate-200 shadow-sm">
                          <span className="text-[10px] font-black text-slate-400 font-mono">A</span>
                          <select
                            value={printBaseSize}
                            onChange={(e) => setPrintBaseSize(e.target.value as any)}
                            className="bg-transparent text-xs font-bold text-slate-800 outline-none cursor-pointer pr-1"
                          >
                            <option value="xs">Size XS</option>
                            <option value="sm">Size SM</option>
                            <option value="base">Size MD</option>
                            <option value="lg">Size LG</option>
                          </select>
                        </div>

                        {/* Divider */}
                        <div className="hidden sm:block h-5 w-[1px] bg-slate-300"></div>

                        {/* Paper Width Toggle */}
                        <div className="flex bg-white rounded-lg p-0.5 border border-slate-200 shadow-sm text-[10px] font-bold">
                          <button
                            type="button"
                            onClick={() => setPrintPaperWidth('58mm')}
                            className={`px-2.5 py-1 rounded transition-all cursor-pointer ${
                              printPaperWidth === '58mm' ? 'bg-slate-800 text-white' : 'text-slate-500 hover:text-slate-800'
                            }`}
                          >
                            58mm
                          </button>
                          <button
                            type="button"
                            onClick={() => setPrintPaperWidth('80mm')}
                            className={`px-2.5 py-1 rounded transition-all cursor-pointer ${
                              printPaperWidth === '80mm' ? 'bg-slate-800 text-white' : 'text-slate-500 hover:text-slate-800'
                            }`}
                          >
                            80mm
                          </button>
                        </div>
                      </div>

                      {/* Section 2: Active Line Formatting Controls */}
                      <div className="flex flex-wrap items-center justify-between gap-3 pt-0.5">
                        <div className="flex items-center flex-wrap gap-2">
                          <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mr-1">
                            Gaya Baris {activeLineIndex !== null ? `#${(activeLineIndex + 1).toString().padStart(2, '0')}` : 'Belum Dipilih'} :
                          </span>

                          {/* Bold Toggle */}
                          <button
                            type="button"
                            disabled={activeLineIndex === null}
                            onClick={() => {
                              if (activeLineIndex !== null) {
                                const newLines = [...printLines];
                                newLines[activeLineIndex].isBold = !newLines[activeLineIndex].isBold;
                                setPrintLines(newLines);
                              }
                            }}
                            className={`w-7 h-7 flex items-center justify-center rounded-lg border transition-all cursor-pointer ${
                              activeLineIndex === null
                                ? 'opacity-30 cursor-not-allowed bg-slate-100 border-transparent text-slate-400'
                                : printLines[activeLineIndex]?.isBold
                                ? 'bg-slate-800 border-slate-800 text-white font-black'
                                : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                            }`}
                            title="Tebalkan Teks (Bold)"
                          >
                            <Bold className="w-3.5 h-3.5" />
                          </button>

                          {/* Active Line Size selection group */}
                          <div className={`flex rounded-lg p-0.5 bg-white border border-slate-200 shadow-sm ${activeLineIndex === null ? 'opacity-35 pointer-events-none' : ''}`}>
                            {(['sm', 'md', 'lg', 'xl'] as const).map((szOption) => (
                              <button
                                key={szOption}
                                type="button"
                                onClick={() => {
                                  if (activeLineIndex !== null) {
                                    const newLines = [...printLines];
                                    newLines[activeLineIndex].size = szOption;
                                    setPrintLines(newLines);
                                  }
                                }}
                                className={`px-2 py-1 text-[10px] font-black rounded transition-all uppercase cursor-pointer ${
                                  activeLineIndex !== null && printLines[activeLineIndex]?.size === szOption
                                    ? 'bg-emerald-600 text-white font-bold'
                                    : 'text-slate-400 hover:text-slate-700'
                                }`}
                                title={`Ukuran Teks: ${szOption}`}
                              >
                                {szOption}
                              </button>
                            ))}
                          </div>

                          {/* Active Line Alignment group */}
                          <div className={`flex rounded-lg p-0.5 bg-white border border-slate-200 shadow-sm ${activeLineIndex === null ? 'opacity-35 pointer-events-none' : ''}`}>
                            <button
                              type="button"
                              onClick={() => {
                                if (activeLineIndex !== null) {
                                  const newLines = [...printLines];
                                  newLines[activeLineIndex].align = 'left';
                                  setPrintLines(newLines);
                                }
                              }}
                              className={`p-1 rounded transition-all cursor-pointer ${
                                activeLineIndex !== null && printLines[activeLineIndex]?.align === 'left'
                                  ? 'bg-slate-800 text-white'
                                  : 'text-slate-400 hover:text-slate-700'
                              }`}
                              title="Rata Kiri"
                            >
                              <AlignLeft className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                if (activeLineIndex !== null) {
                                  const newLines = [...printLines];
                                  newLines[activeLineIndex].align = 'center';
                                  setPrintLines(newLines);
                                }
                              }}
                              className={`p-1 rounded transition-all cursor-pointer ${
                                activeLineIndex !== null && printLines[activeLineIndex]?.align === 'center'
                                  ? 'bg-slate-800 text-white'
                                  : 'text-slate-400 hover:text-slate-700'
                              }`}
                              title="Rata Tengah"
                            >
                              <AlignCenter className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                if (activeLineIndex !== null) {
                                  const newLines = [...printLines];
                                  newLines[activeLineIndex].align = 'right';
                                  setPrintLines(newLines);
                                }
                              }}
                              className={`p-1 rounded transition-all cursor-pointer ${
                                activeLineIndex !== null && printLines[activeLineIndex]?.align === 'right'
                                  ? 'bg-slate-800 text-white'
                                  : 'text-slate-400 hover:text-slate-700'
                              }`}
                              title="Rata Kanan"
                            >
                              <AlignRight className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>

                        {/* Re-order & delete controls in Ribbon */}
                        <div className="flex items-center space-x-1">
                          {/* Move active line up */}
                          <button
                            type="button"
                            disabled={activeLineIndex === null || activeLineIndex === 0}
                            onClick={() => {
                              if (activeLineIndex !== null && activeLineIndex > 0) {
                                const newLines = [...printLines];
                                const temp = newLines[activeLineIndex];
                                newLines[activeLineIndex] = newLines[activeLineIndex - 1];
                                newLines[activeLineIndex - 1] = temp;
                                setPrintLines(newLines);
                                setActiveLineIndex(activeLineIndex - 1);
                              }
                            }}
                            className="w-7 h-7 flex items-center justify-center bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-lg disabled:opacity-30 disabled:hover:bg-white transition-all cursor-pointer shadow-sm"
                            title="Naikkan Baris"
                          >
                            <ChevronUp className="w-3.5 h-3.5" />
                          </button>

                          {/* Move active line down */}
                          <button
                            type="button"
                            disabled={activeLineIndex === null || activeLineIndex === printLines.length - 1}
                            onClick={() => {
                              if (activeLineIndex !== null && activeLineIndex < printLines.length - 1) {
                                const newLines = [...printLines];
                                const temp = newLines[activeLineIndex];
                                newLines[activeLineIndex] = newLines[activeLineIndex + 1];
                                newLines[activeLineIndex + 1] = temp;
                                setPrintLines(newLines);
                                setActiveLineIndex(activeLineIndex + 1);
                              }
                            }}
                            className="w-7 h-7 flex items-center justify-center bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-lg disabled:opacity-30 disabled:hover:bg-white transition-all cursor-pointer shadow-sm"
                            title="Turunkan Baris"
                          >
                            <ChevronDown className="w-3.5 h-3.5" />
                          </button>

                          {/* Delete active line */}
                          <button
                            type="button"
                            disabled={activeLineIndex === null}
                            onClick={() => {
                              if (activeLineIndex !== null) {
                                handleDeleteLine(activeLineIndex);
                              }
                            }}
                            className="w-7 h-7 flex items-center justify-center bg-white border border-red-200 hover:bg-red-50 text-red-500 rounded-lg disabled:opacity-30 disabled:hover:bg-white transition-all cursor-pointer shadow-sm"
                            title="Hapus Baris"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    </div>

                    {/* INTERACTIVE DOCUMENT SHEET (NOTES / FREE-FLOW STYLE) */}
                    <div className="space-y-2.5">
                      <div className="flex items-center justify-between">
                        <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider">
                          LEMBAR CATATAN STRUK (NOTES VIEW)
                        </label>
                        <span className="text-[10px] text-emerald-600 font-semibold bg-emerald-50 px-2 py-0.5 rounded-full">
                          Ketik Bebas • Enter = Baris Baru • Arrow Up/Down
                        </span>
                      </div>

                      {/* Luxurious warm cream notebook paper layout */}
                      <div className="bg-[#FCFBF8] border border-slate-300 rounded-2xl shadow-md p-6 min-h-[300px] flex flex-col justify-between max-h-[380px] overflow-y-auto relative ring-1 ring-black/5">
                        
                        {/* Red margin binder line (mimics real notebook paper) */}
                        <div className="absolute left-[44px] top-0 bottom-0 w-[1px] bg-red-200/65 select-none pointer-events-none"></div>

                        <div className="space-y-1 flex-1 relative z-10">
                          {printLines.map((line, idx) => {
                            const isActive = activeLineIndex === idx;
                            return (
                              <div 
                                key={idx}
                                onClick={() => setActiveLineIndex(idx)}
                                className={`group flex items-center space-x-4 py-1.5 px-1 transition-all rounded-lg ${
                                  isActive 
                                    ? 'bg-amber-100/35 shadow-sm' 
                                    : 'hover:bg-slate-100/25'
                                }`}
                              >
                                {/* Minimalist Page Margin Line Number / Indicator */}
                                <span className={`text-[10px] font-mono select-none w-6 text-right font-semibold tracking-tighter ${
                                  isActive ? 'text-emerald-600 font-bold scale-110' : 'text-slate-300'
                                }`}>
                                  {(idx + 1).toString().padStart(2, '0')}
                                </span>

                                {/* Editable Input Field flowing like simple document lines */}
                                <input
                                  type="text"
                                  value={line.text}
                                  onChange={(e) => {
                                    const newLines = [...printLines];
                                    newLines[idx].text = e.target.value;
                                    setPrintLines(newLines);
                                  }}
                                  onFocus={() => setActiveLineIndex(idx)}
                                  onKeyDown={(e) => handleLineKeyDown(e, idx)}
                                  style={{
                                    fontWeight: line.isBold ? 'bold' : 'normal',
                                    textAlign: line.align,
                                    fontSize: line.size === 'sm' ? '13px' : line.size === 'md' ? '15px' : line.size === 'lg' ? '18px' : '21px',
                                    fontFamily: printFontFamily === 'mono' ? 'Courier New, Courier, monospace' : printFontFamily === 'sans' ? 'Inter, sans-serif' : 'Georgia, serif',
                                  }}
                                  className="flex-1 bg-transparent border-none outline-none text-slate-800 p-0 transition-all placeholder:font-normal placeholder:italic placeholder:text-slate-300"
                                  placeholder="Mulai menulis baris struk di sini..."
                                  id={`line-input-${idx}`}
                                />
                              </div>
                            );
                          })}
                        </div>

                        {/* Add line helper button at bottom of notebook */}
                        <div className="pt-4 border-t border-dashed border-slate-200/80 mt-6 relative z-10">
                          <button
                            type="button"
                            onClick={() => handleAddLine(printLines.length - 1)}
                            className="w-full py-2 bg-white/75 hover:bg-emerald-50 text-slate-500 hover:text-emerald-700 border border-slate-200 hover:border-emerald-300 rounded-xl text-xs font-semibold transition-all cursor-pointer flex items-center justify-center space-x-1.5 shadow-sm hover:shadow"
                          >
                            <Plus className="w-4 h-4" />
                            <span>Tambah Baris Catatan</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Action Buttons */}
                  <div className="pt-4 border-t border-slate-100 space-y-2 mt-4 md:mt-0">
                    <div className="grid grid-cols-2 gap-2">
                      <button
                        onClick={async () => {
                          try {
                            const formattedPlain = printLines.map(l => l.text).join('\n');
                            await navigator.clipboard.writeText(formattedPlain);
                            showToast('Teks struk berhasil disalin!');
                          } catch (err) {
                            showToast('Gagal menyalin teks.');
                          }
                        }}
                        className="py-2.5 px-4 text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl transition-all cursor-pointer flex items-center justify-center space-x-1.5"
                      >
                        <Copy className="w-3.5 h-3.5" />
                        <span>Salin Teks</span>
                      </button>
                      <button
                        onClick={() => {
                          window.print();
                        }}
                        className="py-2.5 px-4 text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl shadow-md transition-all cursor-pointer flex items-center justify-center space-x-1.5"
                      >
                        <Printer className="w-3.5 h-3.5" />
                        <span>Cetak Struk</span>
                      </button>
                    </div>
                    <button
                      onClick={() => setShowPrintModal(false)}
                      className="w-full py-2 text-xs font-bold text-slate-500 hover:bg-slate-100 rounded-xl transition-all cursor-pointer text-center"
                    >
                      Tutup
                    </button>
                  </div>
                </div>

                {/* Right Column: Live Struk Preview (The "Paper" view) */}
                <div className="w-full md:w-[320px] bg-slate-50/70 p-6 flex flex-col items-center justify-start border-t md:border-t-0 md:border-l border-slate-100 relative max-h-[90vh] overflow-y-auto">
                  <p className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-3">TAMPILAN LIVE PREVIEW</p>
                  
                  {/* The Thermal Paper Receipt UI Container */}
                  <div className="w-full bg-white shadow-xl border border-slate-200/80 p-5 rounded-b-md relative overflow-hidden flex flex-col justify-between min-h-[360px] max-w-[280px]">
                    {/* Top jagged paper line decoration */}
                    <div className="absolute top-0 left-0 right-0 h-1.5 flex justify-between">
                      {Array.from({ length: 28 }).map((_, i) => (
                        <div key={i} className="w-2.5 h-2.5 bg-slate-50 rotate-45 transform -translate-y-1.5 border-t border-l border-slate-200/30" />
                      ))}
                    </div>

                    {/* Styled Receipt Paper Content */}
                    <div 
                      className="mt-3 flex-1"
                      style={{
                        fontFamily: printFontFamily === 'mono' ? 'Courier New, Courier, monospace' : printFontFamily === 'sans' ? 'Inter, sans-serif' : 'Georgia, serif',
                        fontSize: printBaseSize === 'xs' ? '11px' : printBaseSize === 'sm' ? '13px' : printBaseSize === 'base' ? '15px' : '17px',
                        lineHeight: '1.4'
                      }}
                    >
                      {printLines.map((line, idx) => (
                        <div 
                          key={idx} 
                          style={{
                            fontWeight: line.isBold ? 'bold' : 'normal',
                            fontSize: line.size === 'sm' ? '0.85em' : line.size === 'md' ? '1.1em' : line.size === 'lg' ? '1.3em' : '1.5em',
                            textAlign: line.align,
                            marginBottom: '4px',
                            whiteSpace: 'pre-wrap',
                            wordBreak: 'break-word',
                            lineHeight: '1.3'
                          }}
                          className="text-slate-900"
                        >
                          {line.text}
                        </div>
                      ))}
                    </div>

                    {/* Bottom dashed border */}
                    <div className="border-t border-dashed border-slate-200 mt-6 pt-3 text-center">
                      <p className="text-[8px] font-mono text-slate-400 uppercase tracking-widest">
                        {printPaperWidth} Thermal Printer Mode
                      </p>
                    </div>
                  </div>
                </div>
              </motion.div>
            </div>
          </div>
        )}
      </AnimatePresence>

      {/* Hidden Print Element */}
      <div 
        id="thermal-print-area" 
        className="hidden"
        style={{ 
          width: printPaperWidth === '58mm' ? '58mm' : '80mm',
        }}
      >
        <div style={{
          fontFamily: printFontFamily === 'mono' ? 'Courier New, Courier, monospace' : printFontFamily === 'sans' ? 'Inter, sans-serif' : 'Georgia, serif',
          fontSize: printBaseSize === 'xs' ? '11px' : printBaseSize === 'sm' ? '13px' : printBaseSize === 'base' ? '15px' : '17px',
          lineHeight: '1.4',
          color: '#000',
          backgroundColor: '#fff',
          padding: '4mm'
        }}>
          {printLines.map((line, idx) => (
            <div 
              key={idx} 
              style={{
                fontWeight: line.isBold ? 'bold' : 'normal',
                fontSize: line.size === 'sm' ? '0.85em' : line.size === 'md' ? '1.1em' : line.size === 'lg' ? '1.3em' : '1.5em',
                textAlign: line.align,
                marginBottom: '4px',
                whiteSpace: 'pre-wrap',
                wordBreak: 'break-word',
                lineHeight: '1.3'
              }}
            >
              {line.text}
            </div>
          ))}
        </div>
      </div>

    </div>
  );
}
