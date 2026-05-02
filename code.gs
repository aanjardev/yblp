  // ==================== KONFIGURASI ====================
  const SHEET_ID = '..'; 

  // Nama sheet yang digunakan
  const SHEET_USERS = 'Users';
  const SHEET_ANAK_YATIM = 'AnakYatim';
  const SHEET_DUAFA = 'Duafa';
  const SHEET_BEDAH_RUMAH = 'BedahRumah';
  const SHEET_KHITAN_MASAL = 'KhitanMasal';

  // Nama sheet program sessions (schema final)
  const SHEET_BEDAH_SESSIONS = 'BedahSessions';

  // Sheet tambahan untuk relasi
  const SHEET_SANTUNAN_SESSIONS = 'SantunanSessions';
  const SHEET_SANTUNAN_PENERIMA = 'SantunanPenerima';
  const SHEET_KHITAN_SESSIONS = 'KhitanSessions';
  const SHEET_KHITAN_PESERTA = 'KhitanPeserta';

  // Headers untuk relasi
  const HEADERS_SANTUNAN_SESSIONS = ['ID', 'NamaKegiatan', 'Tanggal', 'Lokasi', 'TotalAnggaran', 'Keterangan'];
  const HEADERS_SANTUNAN_PENERIMA = ['ID', 'SantunanID', 'PenerimaID', 'JenisPenerima', 'Status', 'Keterangan'];
  const HEADERS_KHITAN_SESSIONS = ['ID', 'NamaKegiatan', 'Tanggal', 'Lokasi', 'Dokter', 'TotalAnggaran', 'Keterangan'];
  const HEADERS_KHITAN_PESERTA = ['ID', 'KhitanSessionID', 'NamaAnak', 'Umur', 'NamaOrtu', 'NoHPOrtu', 'Alamat', 'Status', 'Keterangan'];

  const HEADERS_BEDAH_SESSIONS = ['ID', 'NamaPenerima', 'NoKTP', 'NoHP', 'AlamatLengkap', 'LuasTanah', 'LuasBangunan', 'StatusTanah', 'JenisBedah', 'Anggaran', 'TanggalMulai', 'TanggalSelesai', 'StatusPengerjaan', 'Keterangan'];

  const SHEET_HEADERS_MAP = {
    [SHEET_SANTUNAN_PENERIMA]: HEADERS_SANTUNAN_PENERIMA,
    [SHEET_KHITAN_PESERTA]: HEADERS_KHITAN_PESERTA,
    [SHEET_BEDAH_SESSIONS]: HEADERS_BEDAH_SESSIONS,
    [SHEET_SANTUNAN_SESSIONS]: HEADERS_SANTUNAN_SESSIONS,
    [SHEET_KHITAN_SESSIONS]: HEADERS_KHITAN_SESSIONS
  };

  // ==================== HALAMAN UTAMA ====================
  function doGet(e) {
    // Cek parameter page dari URL
    const page = e?.parameter?.page;
    
    // Render dashboard jika page=dashboard
    if (page === 'dashboard') {
      return HtmlService.createHtmlOutputFromFile('dashboard')
        .setTitle('Dashboard - Yayasan Bumi Langgat')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    // Default: tampilkan index (unified page dengan login + dashboard)
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Yayasan Bumi Langgat Peduli')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // ==================== FUNGSI LOGIN ====================
  function loginUser(email, password) {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_USERS);
      if (!sheet) {
        return { success: false, message: 'Sheet Users tidak ditemukan! Periksa nama sheet.' };
      }
      
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return { success: false, message: 'Tidak ada data di sheet Users.' };
      }
      
      for (let i = 1; i < data.length; i++) {
        const sheetEmail = data[i][0];
        const sheetPassword = data[i][1];
        
        if (sheetEmail === email && sheetPassword === password) {
          // Generate session token
          const token = Utilities.getUuid();
          setSession(token, email);
          return { success: true, message: 'Login berhasil', email: email, token: token };
        }
      }
      
      return { success: false, message: 'Email atau password salah' };
    } catch(e) {
      return { success: false, message: 'Error: ' + e.toString() };
    }
  }

  // Simpan session di PropertiesService (lebih reliable per user)
  function setSession(token, email) {
    const props = PropertiesService.getScriptProperties();
    const expiry = Date.now() + (3600 * 1000); // 1 jam
    props.setProperty('session_' + token, JSON.stringify({ email: email, expiry: expiry }));
    return token;
  }

  function getSession(token) {
    if (!token) return null;
    const props = PropertiesService.getScriptProperties();
    const data = props.getProperty('session_' + token);
    if (!data) return null;
    const session = JSON.parse(data);
    if (Date.now() > session.expiry) {
      props.deleteProperty('session_' + token);
      return null;
    }
    return session;
  }

  // Fungsi untuk cek status login
  function checkLogin(token) {
    const session = getSession(token);
    return { loggedIn: !!session, email: session?.email };
  }

  // Fungsi logout
  function logout(token) {
    if (token) {
      const props = PropertiesService.getScriptProperties();
      props.deleteProperty('session_' + token);
    }
    return { success: true };
  }

  // ==================== FUNGSI UNTUK DASHBOARD ====================
  function getDashboardHTML() {
    return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('Dashboard - Yayasan Bumi Langgat')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .getContent();
  }

  // ==================== FUNGSI GENERIK CRUD ====================
  function getSheetData(sheetName, headers) {
    // Defensive: ensure we never return null
    if (!sheetName || !headers || !Array.isArray(headers)) {
      console.error('getSheetData: invalid params', { sheetName, headers });
      return { headers: [], data: [], error: 'Invalid parameters' };
    }

    try {
      // Standalone script: HARUS pakai openById
      const ss = SpreadsheetApp.openById(SHEET_ID);
      if (!ss) {
        console.error('getSheetData: Cannot open spreadsheet by ID: ' + SHEET_ID);
        return { headers: [], data: [], error: 'Cannot open spreadsheet' };
      }

      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        console.error('getSheetData: Sheet not found: ' + sheetName);
        return { headers: [], data: [], error: 'Sheet not found: ' + sheetName };
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        console.log('getSheetData: Sheet is empty or headers only', { sheetName, lastRow });
        return { headers: headers, data: [] };
      }

      const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      const data = rows.map((row, index) => {
        const obj = { rowIndex: index + 2 };
        headers.forEach((header, colIndex) => {
          let value = row[colIndex];
          // Convert Date objects to string for serialization
          if (value instanceof Date) {
            value = value.toISOString().split('T')[0]; // YYYY-MM-DD format
          }
          obj[header] = value;
        });
        return obj;
      });

      console.log('getSheetData: success', { sheetName, count: data.length });
      return { headers: headers, data: data };
    } catch(e) {
      console.error('getSheetData: exception', { sheetName, error: e.toString(), stack: e.stack });
      return { headers: [], data: [], error: e.toString() };
    }
  }

function addDataToSheet(sheetName, headers, formData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: 'Error: Sheet tidak ditemukan: ' + sheetName };
    }
    const lastRow = sheet.getLastRow();
      const newId = lastRow;
      
      const newRow = headers.map(header => {
        if (header === 'ID') return newId;
        if (header === 'Umur' && formData.TanggalLahir) return hitungUmur(formData.TanggalLahir);
        return formData[header] || '';
      });
      
      sheet.appendRow(newRow);
      return { success: true, message: 'Data berhasil ditambahkan', insertedId: newId };
    } catch(e) {
      return { success: false, message: 'Error: ' + e.toString() };
    }
  }

  function updateDataInSheet(sheetName, headers, updatedData, originalRowIndex) {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(sheetName);
      
      headers.forEach((header, colIndex) => {
        let value = updatedData[header];
        if (header === 'Umur' && updatedData.TanggalLahir) {
          value = hitungUmur(updatedData.TanggalLahir);
        }
        if (value !== undefined) {
          sheet.getRange(originalRowIndex, colIndex + 1).setValue(value);
        }
      });
      
      return { success: true, message: 'Data berhasil diupdate' };
    } catch(e) {
      return { success: false, message: 'Error: ' + e.toString() };
    }
  }

  function deleteDataFromSheet(sheetName, rowIndex) {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(sheetName);
      sheet.deleteRow(rowIndex);
      return { success: true, message: 'Data berhasil dihapus' };
    } catch(e) {
      return { success: false, message: 'Error: ' + e.toString() };
    }
  }

  function hitungUmur(tanggalLahir) {
    if (!tanggalLahir) return '';
    const lahir = new Date(tanggalLahir);
    const today = new Date();
    let umur = today.getFullYear() - lahir.getFullYear();
    const m = today.getMonth() - lahir.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < lahir.getDate())) {
      umur--;
    }
    return umur + ' tahun';
  }

  // ==================== ANAK YATIM ====================
  const HEADERS_ANAK_YATIM = [
    'ID', 'NamaLengkap', 'TanggalLahir', 'Umur', 'JenisKelamin', 
    'NamaWali', 'NoHPWali', 'Alamat', 'StatusSekolah', 'NamaSekolah', 
    'Kelas', 'StatusAktif', 'Keterangan'
  ];

  function getAnakYatim() {
    // Inline implementation untuk standalone script
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_ANAK_YATIM);
      if (!sheet) {
        console.error('getAnakYatim: Sheet not found');
        return { headers: HEADERS_ANAK_YATIM, data: [], error: 'Sheet not found' };
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        console.log('getAnakYatim: Empty sheet');
        return { headers: HEADERS_ANAK_YATIM, data: [] };
      }

      const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS_ANAK_YATIM.length).getValues();
      const data = rows.map((row, index) => {
        const obj = { rowIndex: index + 2 };
        HEADERS_ANAK_YATIM.forEach((header, colIndex) => {
          let value = row[colIndex];
          // Convert Date objects to string for serialization
          if (value instanceof Date) {
            value = value.toISOString().split('T')[0]; // YYYY-MM-DD format
          }
          obj[header] = value;
        });
        return obj;
      });

      console.log('getAnakYatim: success, count:', data.length);
      return { headers: HEADERS_ANAK_YATIM, data: data };
    } catch(e) {
      console.error('getAnakYatim: exception', e.toString(), e.stack);
      return { headers: HEADERS_ANAK_YATIM, data: [], error: e.toString() };
    }
  }

  function addAnakYatim(data) {
    return addDataToSheet(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM, data);
  }
  function updateAnakYatim(data, rowIndex) {
    return updateDataInSheet(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM, data, rowIndex);
  }
  function deleteAnakYatim(rowIndex) {
    return deleteDataFromSheet(SHEET_ANAK_YATIM, rowIndex);
  }

  // ==================== DUAFA ====================
  const HEADERS_DUAFA = [
    'ID', 'NamaLengkap', 'TanggalLahir', 'Umur', 'NoKTP', 
    'NamaKK', 'JmlTanggungan', 'Alamat', 'Kondisi', 'Penghasilan', 
    'SumberPenghasilan', 'StatusAktif', 'Keterangan'
  ];

  function getDuafa() {
    return getSheetData(SHEET_DUAFA, HEADERS_DUAFA);
  }
  function addDuafa(data) {
    return addDataToSheet(SHEET_DUAFA, HEADERS_DUAFA, data);
  }
  function updateDuafa(data, rowIndex) {
    return updateDataInSheet(SHEET_DUAFA, HEADERS_DUAFA, data, rowIndex);
  }
  function deleteDuafa(rowIndex) {
    return deleteDataFromSheet(SHEET_DUAFA, rowIndex);
  }

  // ==================== BEDAH RUMAH ====================
  const HEADERS_BEDAH_RUMAH = [
    'ID', 'NamaPemilik', 'NoKTP', 'NoHP', 'AlamatLengkap', 
    'LuasTanah', 'LuasBangunan', 'StatusTanah', 'JenisBedah', 
    'Anggaran', 'TanggalPelaksanaan', 'StatusPengerjaan', 'Keterangan'
  ];

  function getBedahRumah() {
    return getSheetData(SHEET_BEDAH_RUMAH, HEADERS_BEDAH_RUMAH);
  }
  function addBedahRumah(data) {
    return addDataToSheet(SHEET_BEDAH_RUMAH, HEADERS_BEDAH_RUMAH, data);
  }
  function updateBedahRumah(data, rowIndex) {
    return updateDataInSheet(SHEET_BEDAH_RUMAH, HEADERS_BEDAH_RUMAH, data, rowIndex);
  }
  function deleteBedahRumah(rowIndex) {
    return deleteDataFromSheet(SHEET_BEDAH_RUMAH, rowIndex);
  }

  // ==================== KHITAN MASAL ====================
  const HEADERS_KHITAN_MASAL = [
    'ID', 'NamaAnak', 'TanggalLahir', 'Umur', 'NamaOrtu', 
    'NoHPOrtu', 'Alamat', 'TanggalKhitan', 'Waktu', 'Status', 'Keterangan'
  ];

  function getKhitanMasal() {
    return getSheetData(SHEET_KHITAN_MASAL, HEADERS_KHITAN_MASAL);
  }
  function addKhitanMasal(data) {
    return addDataToSheet(SHEET_KHITAN_MASAL, HEADERS_KHITAN_MASAL, data);
  }
  function updateKhitanMasal(data, rowIndex) {
    return updateDataInSheet(SHEET_KHITAN_MASAL, HEADERS_KHITAN_MASAL, data, rowIndex);
  }
  function deleteKhitanMasal(rowIndex) {
    return deleteDataFromSheet(SHEET_KHITAN_MASAL, rowIndex);
  }

  // ==================== SANTUNAN SESSIONS & PENERIMA ====================
function getSantunanSessions() {
  const sessions = getSheetData(SHEET_SANTUNAN_SESSIONS, HEADERS_SANTUNAN_SESSIONS);
  const penerima = getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data || [];
  const data = (sessions.data || []).map(session => ({
    ...session,
    JumlahPeserta: penerima.filter(p => String(p.SantunanID) === String(session.ID)).length
  }));
  return { headers: [...HEADERS_SANTUNAN_SESSIONS, 'JumlahPeserta'], data: data };
}
  function addSantunanSession(data) {
    return addDataToSheet(SHEET_SANTUNAN_SESSIONS, HEADERS_SANTUNAN_SESSIONS, data);
  }
  function updateSantunanSession(data, rowIndex) {
    return updateDataInSheet(SHEET_SANTUNAN_SESSIONS, HEADERS_SANTUNAN_SESSIONS, data, rowIndex);
  }
  function deleteSantunanSession(rowIndex) {
    return deleteDataFromSheet(SHEET_SANTUNAN_SESSIONS, rowIndex);
  }

  // Get session dengan penerima lengkap (lookup ke master data)
  function getSantunanSessionWithPenerima(sessionId) {
    const session = getSheetData(SHEET_SANTUNAN_SESSIONS, HEADERS_SANTUNAN_SESSIONS).data
      .find(row => row.ID === sessionId);
    
    if (!session) return null;
    
    // Ambil penerima untuk session ini
    const penerima = getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data
      .filter(row => row.SantunanID === sessionId);
    
    // Lookup data master untuk setiap penerima
    const anakData = getSheetData(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM).data;
    const duafaData = getSheetData(SHEET_DUAFA, HEADERS_DUAFA).data;
    
    const penerimaLengkap = penerima.map(p => {
      const masterData = p.JenisPenerima === 'AnakYatim' 
        ? anakData.find(a => a.ID === p.PenerimaID)
        : duafaData.find(d => d.ID === p.PenerimaID);
      
      return {
        ...p,
        Nama: masterData?.NamaLengkap || '-',
        Alamat: masterData?.Alamat || '-',
        Usia: masterData?.Umur || masterData?.Usia || '-',
        Kontak: masterData?.NoHPWali || masterData?.NoHP || '-'
      };
    });
    
    return {
      ...session,
      Penerima: penerimaLengkap,
      JumlahPenerima: penerimaLengkap.length
    };
  }

  function getSantunanPenerima() {
    return getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA);
  }
  function addSantunanPenerima(data) {
    return addDataToSheet(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA, data);
  }
  function addMultipleSantunanPenerima(sessionId, penerimaList) {
    // penerimaList = [{PenerimaID, JenisPenerima, Status, Keterangan}, ...]
    const existing = getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data || [];
    const existingKeySet = new Set(
      existing
        .filter(row => String(row.SantunanID) === String(sessionId))
        .map(row => `${row.JenisPenerima}::${String(row.PenerimaID)}`)
    );

    const results = [];
    let skipped = 0;
    penerimaList.forEach(p => {
      const key = `${p.JenisPenerima}::${String(p.PenerimaID)}`;
      if (existingKeySet.has(key)) {
        skipped++;
        return;
      }

      const data = {
        SantunanID: sessionId,
        PenerimaID: p.PenerimaID,
        JenisPenerima: p.JenisPenerima,
        Status: p.Status || 'Diterima',
        Keterangan: p.Keterangan || ''
      };
      results.push(addSantunanPenerima(data));
      existingKeySet.add(key);
    });
    return {
      success: true,
      results: results,
      inserted: results.length,
      skipped: skipped,
      message: skipped > 0
        ? `${results.length} peserta ditambahkan, ${skipped} peserta duplikat dilewati`
        : `${results.length} peserta ditambahkan`
    };
  }
  function updateSantunanPenerima(data, rowIndex) {
    return updateDataInSheet(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA, data, rowIndex);
  }
  function deleteSantunanPenerima(rowIndex) {
    return deleteDataFromSheet(SHEET_SANTUNAN_PENERIMA, rowIndex);
  }

  // ==================== KHITAN SESSIONS & PESERTA ====================
function getKhitanSessions() {
  const sessions = getSheetData(SHEET_KHITAN_SESSIONS, HEADERS_KHITAN_SESSIONS);
  const peserta = getSheetData(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA).data || [];
  const data = (sessions.data || []).map(session => ({
    ...session,
    TotalPeserta: peserta.filter(p => String(p.KhitanSessionID) === String(session.ID)).length
  }));
  return { headers: [...HEADERS_KHITAN_SESSIONS, 'TotalPeserta'], data: data };
}
  function addKhitanSession(data) {
    return addDataToSheet(SHEET_KHITAN_SESSIONS, HEADERS_KHITAN_SESSIONS, data);
  }
  function updateKhitanSession(data, rowIndex) {
    return updateDataInSheet(SHEET_KHITAN_SESSIONS, HEADERS_KHITAN_SESSIONS, data, rowIndex);
  }
  function deleteKhitanSession(rowIndex) {
    return deleteDataFromSheet(SHEET_KHITAN_SESSIONS, rowIndex);
  }

  function getKhitanPeserta() {
    return getSheetData(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA);
  }
  function addKhitanPeserta(data) {
    return addDataToSheet(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA, data);
  }
  function addMultipleKhitanPeserta(sessionId, pesertaList) {
    // pesertaList = [{NamaAnak, Umur, NamaOrtu, NoHPOrtu, Alamat, Status, Keterangan}, ...]
    const results = [];
    pesertaList.forEach(p => {
      const data = {
        KhitanSessionID: sessionId,
        NamaAnak: p.NamaAnak,
        Umur: p.Umur,
        NamaOrtu: p.NamaOrtu,
        NoHPOrtu: p.NoHPOrtu,
        Alamat: p.Alamat,
        Status: p.Status || 'Terdaftar',
        Keterangan: p.Keterangan || ''
      };
      results.push(addKhitanPeserta(data));
    });
    return { success: true, results: results };
  }
  function updateKhitanPeserta(data, rowIndex) {
    return updateDataInSheet(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA, data, rowIndex);
  }
  function deleteKhitanPeserta(rowIndex) {
    return deleteDataFromSheet(SHEET_KHITAN_PESERTA, rowIndex);
  }

  // ==================== BEDAH SESSIONS ====================
  function getBedahSessions() {
    return getSheetData(SHEET_BEDAH_SESSIONS, HEADERS_BEDAH_SESSIONS);
  }
  function addBedahSession(data) {
    return addDataToSheet(SHEET_BEDAH_SESSIONS, HEADERS_BEDAH_SESSIONS, data);
  }
  function updateBedahSession(data, rowIndex) {
    return updateDataInSheet(SHEET_BEDAH_SESSIONS, HEADERS_BEDAH_SESSIONS, data, rowIndex);
  }
  function deleteBedahSession(rowIndex) {
    return deleteDataFromSheet(SHEET_BEDAH_SESSIONS, rowIndex);
  }

  // ==================== DASHBOARD STATS ====================
  function getDashboardStats() {
    // Data master - hitung hanya yang aktif
    const anakData = getSheetData(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM).data;
    const duafaData = getSheetData(SHEET_DUAFA, HEADERS_DUAFA).data;
    
    const anakAktif = anakData?.filter(row => row.StatusAktif === 'Aktif').length || 0;
    const anakNonAktif = anakData?.filter(row => row.StatusAktif === 'Non-Aktif').length || 0;
    const duafaAktif = duafaData?.filter(row => row.StatusAktif === 'Aktif').length || 0;
    const duafaNonAktif = duafaData?.filter(row => row.StatusAktif === 'Non-Aktif').length || 0;
    
    // Data program sessions
    let bedahBerjalan = 0;
    try {
      const bedahData = getSheetData(SHEET_BEDAH_SESSIONS, HEADERS_BEDAH_SESSIONS).data;
      bedahBerjalan = bedahData?.filter(row => row.StatusPengerjaan === 'Berjalan').length || 0;
    } catch(e) {
      const bedahData = getSheetData(SHEET_BEDAH_RUMAH, HEADERS_BEDAH_RUMAH).data;
      bedahBerjalan = bedahData?.filter(row => row.StatusPengerjaan === 'Berjalan').length || 0;
    }
    
    // Hitung program sessions
    let santunanCount = 0, khitanCount = 0, bedahCount = 0;
    try {
      santunanCount = getSheetData(SHEET_SANTUNAN_SESSIONS, HEADERS_SANTUNAN_SESSIONS).data?.length || 0;
    } catch(e) {}
    try {
      khitanCount = getSheetData(SHEET_KHITAN_SESSIONS, HEADERS_KHITAN_SESSIONS).data?.length || 0;
    } catch(e) {}
    try {
      bedahCount = getSheetData(SHEET_BEDAH_SESSIONS, HEADERS_BEDAH_SESSIONS).data?.length || 0;
    } catch(e) {
      bedahCount = getSheetData(SHEET_BEDAH_RUMAH, HEADERS_BEDAH_RUMAH).data?.length || 0;
    }
    
    return {
      anakAktif: anakAktif,
      anakNonAktif: anakNonAktif,
      duafaAktif: duafaAktif,
      duafaNonAktif: duafaNonAktif,
      bedahBerjalan: bedahBerjalan,
      santunanCount: santunanCount,
      khitanCount: khitanCount,
      bedahCount: bedahCount
    };
  }

  // ==================== RECENT PROGRAMS ====================
  function getRecentPrograms() {
    const results = [];
    const santunanPenerima = getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data || [];
    const khitanPeserta = getSheetData(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA).data || [];
    
    // Ambil recent santunan sessions
    try {
      const santunanData = getSheetData(SHEET_SANTUNAN_SESSIONS, HEADERS_SANTUNAN_SESSIONS).data || [];
      santunanData.slice(-3).forEach(item => {
        results.push({
          type: 'santunan',
          id: item.ID,
          nama: item.NamaKegiatan,
          tanggal: item.Tanggal,
          lokasi: item.Lokasi,
          jumlahPeserta: santunanPenerima.filter(p => String(p.SantunanID) === String(item.ID)).length,
          status: 'Selesai'
        });
      });
    } catch(e) {}
    
    // Ambil recent khitan sessions
    try {
      const khitanData = getSheetData(SHEET_KHITAN_SESSIONS, HEADERS_KHITAN_SESSIONS).data || [];
      khitanData.slice(-3).forEach(item => {
        results.push({
          type: 'khitan',
          id: item.ID,
          nama: item.NamaKegiatan,
          tanggal: item.Tanggal,
          lokasi: item.Lokasi,
          jumlahPeserta: khitanPeserta.filter(p => String(p.KhitanSessionID) === String(item.ID)).length,
          status: 'Selesai'
        });
      });
    } catch(e) {}
    
    // Ambil recent bedah sessions
    try {
      const bedahData = getSheetData(SHEET_BEDAH_SESSIONS, HEADERS_BEDAH_SESSIONS).data || [];
      bedahData.slice(-3).forEach(item => {
        results.push({
          type: 'bedah',
          id: item.ID,
          nama: item.NamaPenerima,
          tanggal: item.TanggalMulai,
          tanggalSelesai: item.TanggalSelesai,
          lokasi: item.AlamatLengkap,
          status: item.StatusPengerjaan
        });
      });
    } catch(e) {
      const bedahData = getSheetData(SHEET_BEDAH_RUMAH, HEADERS_BEDAH_RUMAH).data || [];
      bedahData.slice(-3).forEach(item => {
        results.push({
          type: 'bedah',
          id: item.ID,
          nama: item.NamaPemilik,
          tanggal: item.TanggalPelaksanaan,
          tanggalSelesai: item.TanggalSelesai || '',
          lokasi: item.AlamatLengkap,
          status: item.StatusPengerjaan
        });
      });
    }
    
    // Sort by tanggal (descending)
    results.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
    
    return results.slice(0, 5); // Return max 5 recent items
  }

  // ==================== LAPORAN ====================
  function getAllDataForLaporan() {
    return {
      anakYatim: getSheetData(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM).data,
      duafa: getSheetData(SHEET_DUAFA, HEADERS_DUAFA).data,
      santunanSessions: getSheetData(SHEET_SANTUNAN_SESSIONS, HEADERS_SANTUNAN_SESSIONS).data,
      santunanPenerima: getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data,
      khitanSessions: getSheetData(SHEET_KHITAN_SESSIONS, HEADERS_KHITAN_SESSIONS).data,
      khitanPeserta: getSheetData(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA).data,
      bedahSessions: getSheetData(SHEET_BEDAH_SESSIONS, HEADERS_BEDAH_SESSIONS).data,
      bedahRumah: getSheetData(SHEET_BEDAH_RUMAH, HEADERS_BEDAH_RUMAH).data
    };
  }

  // ==================== GENERIC PARTICIPANT APIs ====================
  function getParticipants(programType, sessionId) {
  if (programType === 'santunan' || programType === 'SantunanPenerima') {
      const penerima = getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data
        .filter(row => row.SantunanID === sessionId);
      
      // Lookup master data untuk display
      const anakData = getSheetData(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM).data;
      const duafaData = getSheetData(SHEET_DUAFA, HEADERS_DUAFA).data;
      
      return penerima.map(p => {
        const masterData = p.JenisPenerima === 'AnakYatim' 
          ? anakData.find(a => a.ID === p.PenerimaID)
          : duafaData.find(d => d.ID === p.PenerimaID);
        
      return {
        ...p,
        NamaLengkap: masterData?.NamaLengkap || '-',
        Kategori: p.JenisPenerima === 'AnakYatim' ? 'Anak Yatim' : 'Duafa',
        Usia: masterData?.Umur || masterData?.Usia || '-',
        Alamat: masterData?.Alamat || '-',
        Kontak: masterData?.NoHPWali || masterData?.NoHP || '-'
      };
    });
  }
    
    if (programType === 'khitan' || programType === 'KhitanPeserta') {
      return getSheetData(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA).data
        .filter(row => row.KhitanSessionID === sessionId);
    }
    
    return [];
  }

  function addParticipant(programType, data) {
    if (programType === 'santunan' || programType === 'SantunanPenerima') {
      const existing = getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data || [];
      const isDuplicate = existing.some(row =>
        String(row.SantunanID) === String(data.SantunanID) &&
        String(row.PenerimaID) === String(data.PenerimaID) &&
        String(row.JenisPenerima) === String(data.JenisPenerima)
      );
      if (isDuplicate) {
        return { success: false, message: 'Penerima ini sudah ada pada program santunan ini' };
      }
      return addDataToSheet(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA, data);
    }
    
    if (programType === 'khitan' || programType === 'KhitanPeserta') {
      return addDataToSheet(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA, data);
    }
    
    return { success: false, message: 'Invalid program type' };
  }

  function updateParticipant(programType, data, rowIndex) {
    if (programType === 'santunan' || programType === 'SantunanPenerima') {
      return updateDataInSheet(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA, data, rowIndex);
    }
    
    if (programType === 'khitan' || programType === 'KhitanPeserta') {
      return updateDataInSheet(SHEET_KHITAN_PESERTA, HEADERS_KHITAN_PESERTA, data, rowIndex);
    }
    
    return { success: false, message: 'Invalid program type' };
  }

  function deleteParticipant(programType, rowIndex) {
    if (programType === 'santunan' || programType === 'SantunanPenerima') {
      return deleteDataFromSheet(SHEET_SANTUNAN_PENERIMA, rowIndex);
    }
    
    if (programType === 'khitan' || programType === 'KhitanPeserta') {
      return deleteDataFromSheet(SHEET_KHITAN_PESERTA, rowIndex);
    }
    
    return { success: false, message: 'Invalid program type' };
  }

  function getMasterDataForDropdown(type) {
    if (type === 'AnakYatim') {
      return getSheetData(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM).data
        .filter(row => row.StatusAktif === 'Aktif')
        .map(row => ({
          ID: row.ID,
          NamaLengkap: row.NamaLengkap,
          Alamat: row.Alamat
        }));
    }
    
    if (type === 'Duafa') {
      return getSheetData(SHEET_DUAFA, HEADERS_DUAFA).data
        .filter(row => row.StatusAktif === 'Aktif')
        .map(row => ({
          ID: row.ID,
          NamaLengkap: row.NamaLengkap,
          Alamat: row.Alamat
        }));
    }
    
    return [];
  }

  function getAktifPenerima(sessionId) {
    const anak = getSheetData(SHEET_ANAK_YATIM, HEADERS_ANAK_YATIM).data
      .filter(row => row.StatusAktif === 'Aktif')
      .map(row => ({
        ID: row.ID,
        Jenis: 'AnakYatim',
        Nama: row.NamaLengkap || '-',
        Alamat: row.Alamat || '-',
        Umur: row.Umur || ''
      }));

    const duafa = getSheetData(SHEET_DUAFA, HEADERS_DUAFA).data
      .filter(row => row.StatusAktif === 'Aktif')
      .map(row => ({
        ID: row.ID,
        Jenis: 'Duafa',
        Nama: row.NamaLengkap || '-',
        Alamat: row.Alamat || '-',
        Umur: row.Umur || ''
      }));

    let merged = [...anak, ...duafa];
    if (sessionId !== undefined && sessionId !== null && String(sessionId) !== '') {
      const existing = getSheetData(SHEET_SANTUNAN_PENERIMA, HEADERS_SANTUNAN_PENERIMA).data || [];
      const existingKeySet = new Set(
        existing
          .filter(row => String(row.SantunanID) === String(sessionId))
          .map(row => `${row.JenisPenerima}::${String(row.PenerimaID)}`)
      );
      merged = merged.filter(row => !existingKeySet.has(`${row.Jenis}::${String(row.ID)}`));
    }

    return merged;
  }
