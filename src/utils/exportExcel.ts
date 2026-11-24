import * as XLSX from 'xlsx';
import type { Transaction, MenuItem, StoreSettings } from '../types/pos';
import { formatDateTime } from './format';

/**
 * Export all POS transaction and menu records to a well-structured Microsoft Excel (.xlsx) spreadsheet
 */
export const exportReportToExcel = (
  transactions: Transaction[],
  menu: MenuItem[],
  settings: StoreSettings
): void => {
  const wb = XLSX.utils.book_new();

  // 1. Sheet: Riwayat Transaksi (All Transactions)
  const txRows = transactions.map((tx, idx) => ({
    'No': idx + 1,
    'No. Order': `#${tx.orderNumber}`,
    'ID Transaksi': tx.id,
    'Tanggal & Waktu': formatDateTime(tx.timestamp),
    'Kasir': tx.cashierName || '-',
    'Pelanggan': tx.customerName || '-',
    'Meja / Catatan': tx.tableNumber || 'Take Away',
    'Metode Bayar': (tx.paymentMethod || 'cash').toUpperCase(),
    'Total Item': tx.items.reduce((sum, item) => sum + item.quantity, 0),
    'Rincian Pesanan': tx.items.map((i) => `${i.quantity}x ${i.name}`).join(', '),
    'Subtotal (Rp)': tx.subtotal,
    'Diskon (Rp)': tx.discountAmount || 0,
    'Service Charge (Rp)': tx.serviceChargeAmount || 0,
    'Pajak PB1 (Rp)': tx.taxAmount || 0,
    'Total Bayar (Rp)': tx.total,
    'Uang Diterima (Rp)': tx.amountPaid,
    'Kembalian (Rp)': tx.change,
  }));

  const wsTx = txRows.length > 0 
    ? XLSX.utils.json_to_sheet(txRows) 
    : XLSX.utils.aoa_to_sheet([['Belum ada transaksi tercatat.']]);

  wsTx['!cols'] = [
    { wch: 6 },  // No
    { wch: 12 }, // No. Order
    { wch: 22 }, // ID Transaksi
    { wch: 22 }, // Tanggal & Waktu
    { wch: 18 }, // Kasir
    { wch: 18 }, // Pelanggan
    { wch: 16 }, // Meja
    { wch: 15 }, // Metode Bayar
    { wch: 12 }, // Total Item
    { wch: 40 }, // Rincian Pesanan
    { wch: 16 }, // Subtotal
    { wch: 14 }, // Diskon
    { wch: 18 }, // Service Charge
    { wch: 16 }, // Pajak PB1
    { wch: 18 }, // Total Bayar
    { wch: 18 }, // Uang Diterima
    { wch: 16 }, // Kembalian
  ];
  XLSX.utils.book_append_sheet(wb, wsTx, 'Riwayat Transaksi');

  // 2. Sheet: Rincian Menu Terjual (Item Level Breakdown)
  const itemRows: Record<string, string | number>[] = [];
  let itemCounter = 1;
  transactions.forEach((tx) => {
    tx.items.forEach((item) => {
      itemRows.push({
        'No': itemCounter++,
        'No. Order': `#${tx.orderNumber}`,
        'ID Transaksi': tx.id,
        'Waktu': formatDateTime(tx.timestamp),
        'Kategori': item.category,
        'Nama Menu': item.name,
        'Varian': item.selectedVariants?.map((v) => v.optionName).join(', ') || '-',
        'Catatan Item': item.notes || '-',
        'Qty': item.quantity,
        'Harga Satuan (Rp)': item.unitPrice,
        'Subtotal (Rp)': item.subtotal,
      });
    });
  });

  const wsItems = itemRows.length > 0
    ? XLSX.utils.json_to_sheet(itemRows)
    : XLSX.utils.aoa_to_sheet([['Belum ada item terjual.']]);

  wsItems['!cols'] = [
    { wch: 6 },  // No
    { wch: 12 }, // No. Order
    { wch: 22 }, // ID Transaksi
    { wch: 22 }, // Waktu
    { wch: 16 }, // Kategori
    { wch: 28 }, // Nama Menu
    { wch: 24 }, // Varian
    { wch: 22 }, // Catatan
    { wch: 10 }, // Qty
    { wch: 18 }, // Harga Satuan
    { wch: 18 }, // Subtotal
  ];
  XLSX.utils.book_append_sheet(wb, wsItems, 'Rincian Menu Terjual');

  // 3. Sheet: Katalog Menu Toko
  const menuRows = menu.map((m, idx) => ({
    'No': idx + 1,
    'ID Menu': m.id,
    'Nama Menu': m.name,
    'Kategori': m.category,
    'Harga Dasar (Rp)': m.basePrice,
    'Badge Khusus': m.badge || '-',
    'Deskripsi': m.description || '-',
    'Status Stok': m.isAvailable ? 'Tersedia' : 'Habis',
  }));

  const wsMenu = menuRows.length > 0
    ? XLSX.utils.json_to_sheet(menuRows)
    : XLSX.utils.aoa_to_sheet([['Katalog menu kosong.']]);

  wsMenu['!cols'] = [
    { wch: 6 },  // No
    { wch: 16 }, // ID Menu
    { wch: 28 }, // Nama Menu
    { wch: 16 }, // Kategori
    { wch: 18 }, // Harga Dasar
    { wch: 16 }, // Badge
    { wch: 38 }, // Deskripsi
    { wch: 14 }, // Status Stok
  ];
  XLSX.utils.book_append_sheet(wb, wsMenu, 'Katalog Menu Toko');

  // Generate file download
  const dateStr = new Date().toISOString().slice(0, 10);
  const cleanStoreName = (settings.storeName || 'cy-coffee')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '-');
  const filename = `laporan-pos-${cleanStoreName}-${dateStr}.xlsx`;

  XLSX.writeFile(wb, filename);
};
