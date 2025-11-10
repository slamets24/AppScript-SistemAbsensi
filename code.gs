// ==================== CODE.GS ====================
const SPREADSHEET_ID = '12B2cL6yhsgoT17qcD1KaDP1Ksvnub2Zlw_-O3lVCYtA';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Absensi & Jam Kerja')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Spreadsheet tidak ditemukan!');
  }
}

function getMonthlyPeriod(year, month) {
  const startDate = new Date(year, month - 1, 2);
  const endDate = new Date(year, month, 1);
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    throw new Error('Periode tidak valid. Pastikan tahun dan bulan benar.');
  }
  return { startDate, endDate };
}


function getEmployees() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('DATA_KARYAWAN');
  
  if (!sheet) {
    sheet = ss.insertSheet('DATA_KARYAWAN');
    sheet.getRange(1, 1).setValue('NAMA');
    sheet.getRange(1, 1).setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');
    
    const defaultEmployees = [
      ['RESTI RAHMATINA'], ['DEDE RIA'], ['FIKRI'],
      ['DEBBY OKE SETIADI'], ['ADINDA'], ['MELANI'],
      ['DIRA'], ['SLAMET SEPTIAWAN'], ['TASYA MARLINA'],
      ['ANISA ARYANI'], ['DEVI SRI WAHYUNI'], ['TASYA AMATUSSALIM']
    ];
    sheet.getRange(2, 1, defaultEmployees.length, 1).setValues(defaultEmployees);
    return defaultEmployees.map(e => e[0]);
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(row => row[0]).filter(name => name);
}

function saveAttendance(employeeName, year, month, day, type, value) {
  const ss = getSpreadsheet();
  const sheetName = `ABSENSI_${year}_${month}`;
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 4).setValues([['NAMA', 'TANGGAL', 'STATUS', 'JAM_LEMBUR']]);
    sheet.getRange(1, 1, 1, 4).setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  const lastRow = sheet.getLastRow();
  
  let rowIndex = -1;
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowDate = formatDateFromSheet(data[i][1]);
      if (data[i][0] === employeeName && rowDate === dateStr) {
        rowIndex = i + 2;
        break;
      }
    }
  }
  
  if (rowIndex === -1) {
    rowIndex = lastRow + 1;
    sheet.getRange(rowIndex, 1).setValue(employeeName);
    sheet.getRange(rowIndex, 2).setValue(dateStr);
  }
  
  if (type === 'status') {
    const statusValue = value || '';
    const cell = sheet.getRange(rowIndex, 3);
    
    // Jika HADIR atau TIDAK_HADIR, simpan sebagai text
    if (statusValue === 'HADIR' || statusValue === 'TIDAK_HADIR' || statusValue === '') {
      cell.setNumberFormat('@').setValue(statusValue);
    } else {
      // Jika angka (jam kerja parsial), simpan sebagai number dengan prefix '
      // Prefix ' membuat Google Sheets treat sebagai text tapi tetap tampil seperti angka
      cell.setValue("'" + statusValue);
    }
  } else if (type === 'overtime') {
    // Fix: Ganti koma dengan titik untuk format angka yang benar
    const cleanValue = String(value).replace(',', '.');
    const numValue = parseFloat(cleanValue) || 0;
    sheet.getRange(rowIndex, 4).setNumberFormat('0.0').setValue(numValue);
  }
  
  return { success: true };
}

function formatDateFromSheet(dateValue) {
  if (!dateValue) return '';
  
  if (typeof dateValue === 'string' && dateValue.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return dateValue;
  }
  
  const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (!isNaN(date.getTime())) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  }
  
  return '';
}

function getAttendanceData(employeeName, year, month) {
  const ss = getSpreadsheet();
  const sheetName = `ABSENSI_${year}_${month}`;
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const result = {};
  
  data.forEach(row => {
    const dateStr = formatDateFromSheet(row[1]);
    if (row[0] === employeeName && dateStr) {
      // Pastikan status selalu string, bahkan jika null/undefined
      let statusValue = row[2];
      if (statusValue === null || statusValue === undefined || statusValue === '') {
        statusValue = '';
      } else {
        // Konversi ke string dan bersihkan dari format yang aneh
        statusValue = String(statusValue).trim();
        
        // Hapus prefix ' jika ada (dari text format)
        if (statusValue.startsWith("'")) {
          statusValue = statusValue.substring(1);
        }
        
        // Jika Google Sheets mengubah jadi Date object, kembalikan ke string kosong
        if (statusValue.match(/^\d{4}-\d{2}-\d{2}/) || statusValue.includes('GMT')) {
          statusValue = '';
        }
      }
      
      // Parse overtime dengan aman
      let overtimeValue = 0;
      if (row[3] !== null && row[3] !== undefined && row[3] !== '') {
        const cleanOT = String(row[3]).replace(',', '.');
        overtimeValue = parseFloat(cleanOT) || 0;
      }
      
      result[dateStr] = {
        status: statusValue,
        overtime: overtimeValue
      };
    }
  });
  
  return result;
}

// ==================== LOGIKA PERHITUNGAN BARU ====================
function calculateWorkHours(employeeName, year, month) {
  const attendanceData = getAttendanceData(employeeName, year, month);
  const period = getMonthlyPeriod(year, month);
  const today = new Date();

  let totalWorkDaysElapsed = 0;
  let totalTargetHours = 0;
  let totalActualWorkHours = 0;
  let totalOvertimeHours = 0;
  let totalSundayHours = 0;
  let daysAbsent = 0;

  const dailyData = [];
  let currentDate = new Date(period.startDate);

  // Loop semua hari dalam periode
  while (currentDate <= period.endDate) {
    const dayOfWeek = currentDate.getDay();
    const dateStr = formatDateString(currentDate);
    const dayData = attendanceData[dateStr] || {};
    const isPastOrToday = currentDate <= today;

    // Tentukan jam kerja target per hari
    let targetHours = 0;
    if (dayOfWeek >= 1 && dayOfWeek <= 5) targetHours = 8.5; // Senin–Jumat
    else if (dayOfWeek === 6) targetHours = 7; // Sabtu
    // Minggu (0) = 0 jam target

    // Ambil data kehadiran
    let actualHours = 0;
    let attendanceStatus = "BELUM_INPUT"; // Default: belum ada input

    if (dayData.status === "HADIR") {
      actualHours = targetHours > 0 ? targetHours : 0;
      attendanceStatus = "HADIR_PENUH";
    } else if (dayData.status === "TIDAK_HADIR") {
      actualHours = 0;
      attendanceStatus = "TIDAK_HADIR";
    } else if (!isNaN(parseFloat(dayData.status)) && dayData.status !== "" && dayData.status !== null && dayData.status !== undefined) {
      actualHours = parseFloat(dayData.status);
      attendanceStatus = "HADIR_SEBAGIAN";
    }

    const overtime = parseFloat(dayData.overtime) || 0;
    
    // Khusus untuk hari Minggu: jika ada lembur tapi status kosong, anggap tidak ada jam kerja normal
    if (dayOfWeek === 0 && overtime > 0 && actualHours === 0) {
      attendanceStatus = "LEMBUR_MINGGU";
    }

    dailyData.push({
      date: dateStr,
      dayOfWeek,
      targetHours,
      actualHours,
      overtime,
      status: attendanceStatus,
      isPastOrToday
    });

    // Hitung hanya untuk hari yang sudah berlalu atau hari ini
    if (isPastOrToday) {
      if (dayOfWeek === 0) {
        // Minggu: jam kerja di Minggu dianggap lembur
        totalSundayHours += actualHours + overtime;
      } else {
        // Hari kerja (Senin-Sabtu)
        totalWorkDaysElapsed++;
        totalTargetHours += targetHours;
        totalActualWorkHours += actualHours;
        totalOvertimeHours += overtime;
        
        if (attendanceStatus === "TIDAK_HADIR") {
          daysAbsent++;
        }
      }
    }

    currentDate.setDate(currentDate.getDate() + 1);
  }

  // ==================== PERHITUNGAN INTI ====================
  
  // Total jam tersedia = jam kerja aktual + jam lembur (Senin-Sabtu) + jam Minggu
  const totalAvailableHours = totalActualWorkHours + totalOvertimeHours + totalSundayHours;
  
  // Hitung kekurangan jam
  let shortageHours = Math.max(0, totalTargetHours - totalActualWorkHours);
  
  // Jam yang bisa digunakan untuk menutup kekurangan
  const hoursToFillShortage = totalOvertimeHours + totalSundayHours;
  
  // Hitung jam kerja bersih dan lembur bersih
  let finalWorkHours = totalActualWorkHours;
  let cleanOvertime = 0;
  
  if (shortageHours > 0) {
    // Ada kekurangan jam kerja
    if (hoursToFillShortage >= shortageHours) {
      // Lembur cukup untuk menutup kekurangan
      finalWorkHours = totalTargetHours;
      cleanOvertime = hoursToFillShortage - shortageHours;
      shortageHours = 0;
    } else {
      // Lembur tidak cukup menutup kekurangan
      finalWorkHours = totalActualWorkHours + hoursToFillShortage;
      cleanOvertime = 0;
      shortageHours = totalTargetHours - finalWorkHours;
    }
  } else {
    // Tidak ada kekurangan, semua lembur adalah lembur bersih
    finalWorkHours = totalActualWorkHours;
    cleanOvertime = hoursToFillShortage;
    shortageHours = 0;
  }
  
  // Hitung hari kerja bersih (dalam satuan hari)
  // Asumsi: 1 hari = 8.5 jam (rata-rata)
  const shortfallDays = shortageHours / 8.5;
  const cleanWorkDays = totalWorkDaysElapsed - shortfallDays;
  
  // Pisahkan hari penuh dan sisa jam untuk tampilan
  const cleanWorkDaysFull = Math.floor(cleanWorkDays);
  const cleanWorkDaysRemainder = cleanWorkDays - cleanWorkDaysFull;
  const cleanWorkHoursRemainder = cleanWorkDaysRemainder * 8.5;
  
  // Pisahkan kekurangan menjadi hari penuh dan sisa jam
  const shortfallDaysFull = Math.floor(shortfallDays);
  const shortfallDaysRemainder = shortfallDays - shortfallDaysFull;
  const shortfallHoursRemainder = shortfallDaysRemainder * 8.5;
  
  // Tentukan apakah target sudah terpenuhi
  const isComplete = shortageHours === 0;
  
  // Format tampilan hari kerja bersih (hanya tampilkan hari penuh, sisa jam ada di kartu terpisah)
  let cleanWorkDaysDisplay = cleanWorkDaysFull.toString();
  
  // Format tampilan kekurangan
  let shortfallDisplay = '0';
  if (shortageHours > 0) {
    if (shortfallHoursRemainder > 0) {
      shortfallDisplay = `${shortfallDaysFull} hari + ${shortfallHoursRemainder.toFixed(1)} jam`;
    } else {
      shortfallDisplay = `${shortfallDaysFull}`;
    }
  }

  return {
    workDaysElapsed: totalWorkDaysElapsed,
    cleanWorkDays: cleanWorkDaysDisplay,
    cleanWorkDaysFull: cleanWorkDaysFull,
    cleanWorkHoursRemainder: cleanWorkHoursRemainder.toFixed(1),
    shortfallDays: shortfallDisplay,
    shortfallHours: shortageHours.toFixed(1),
    cleanOvertime: cleanOvertime.toFixed(1),
    remainingHours: cleanWorkHoursRemainder.toFixed(1),
    isComplete: isComplete,
    absentDays: daysAbsent,
    totalWorkHours: finalWorkHours.toFixed(1),
    totalOvertimeInput: (totalOvertimeHours + totalSundayHours).toFixed(1),
    totalAvailableHours: totalAvailableHours.toFixed(1),
    totalTargetHours: totalTargetHours.toFixed(1),
    periodInfo: `${formatDateLocal(period.startDate)} - ${formatDateLocal(period.endDate)}`
  };
}



function formatDateString(date) {
  if (!date) return '';
  const d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}


function formatDateLocal(date) {
  if (!date) return '-';
  const d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '-';
  return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
}

function testConnection() {
  try {
    const ss = getSpreadsheet();
    Logger.log('✅ Terhubung: ' + ss.getName());
    return 'SUCCESS';
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    return 'ERROR';
  }
}

//
 ==================== RANGKUMAN SEMUA KARYAWAN ====================
function getAllEmployeesSummary(year, month) {
  const employees = getEmployees();
  const summary = [];
  
  employees.forEach(employeeName => {
    const stats = calculateWorkHours(employeeName, year, month);
    summary.push({
      name: employeeName,
      workDaysElapsed: stats.workDaysElapsed,
      cleanWorkDays: stats.cleanWorkDays,
      remainingHours: stats.remainingHours,
      shortfallDays: stats.shortfallDays,
      shortfallHours: stats.shortfallHours,
      cleanOvertime: stats.cleanOvertime,
      absentDays: stats.absentDays,
      isComplete: stats.isComplete
    });
  });
  
  return summary;
}
