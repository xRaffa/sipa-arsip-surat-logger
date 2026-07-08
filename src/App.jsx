import React, { useState, useEffect } from 'react';
import { supabase } from './supabaseClient';
import { uploadFileToDrive, generateRandom16CharFilename } from './googleDriveHelper';
import * as XLSX from 'xlsx';
import {
  Folder,
  FileText,
  Plus,
  Search,
  Download,
  LogOut,
  AlertTriangle,
  CheckCircle,
  Settings,
  Layers,
  Database,
  Cloud,
  ChevronRight,
  Loader2,
  Calendar,
  User,
  Info
} from 'lucide-react';

// Dropdown Constants
const JENIS_SURAT_OPTIONS = [
  { value: 'O', label: 'O - Lain-lain' },
  { value: 'SK', label: 'SK - Surat Keputusan/Kebijakan' },
  { value: 'SE', label: 'SE - Surat Edaran' },
  { value: 'SH', label: 'SH - Surat Himbauan' },
  { value: 'SB', label: 'SB - Surat Balasan' },
  { value: 'Sket', label: 'Sket - Surat Keterangan' },
  { value: 'SR', label: 'SR - Surat Riset & Magang' },
  { value: 'SPM', label: 'SPM - Surat Permintaan' },
  { value: 'SP', label: 'SP - Surat Peringatan' },
  { value: 'SPD', label: 'SPD - Surat Permintaan Dana' },
  { value: 'SKK', label: 'SKK - Surat Kontrak Kerjasama' },
  { value: 'SOP', label: 'SOP - Surat Standar Operasional Prosedur' },
  { value: 'SKP', label: 'SKP - Surat Kontrak Pengadaan Barang/Jasa' },
  { value: 'SU', label: 'SU - Surat Umum' }
];

const PENUNJUKAN_SURAT_OPTIONS = [
  { value: 'A', label: 'A - Yayasan Pendidikan Harapan' },
  { value: 'B', label: 'B - Kopertis wilayah I' },
  { value: 'C', label: 'C - BR-PTSI & Aptisi' },
  { value: 'D', label: 'D - Dikti, BAN-PT, Didjen Perti' },
  { value: 'E', label: 'E - Umum (Instansi Pemerintah)' },
  { value: 'F', label: 'F - Umum (Instansi Non Pemerintah)' },
  { value: 'G', label: 'G - Dosen Luar Biasa' },
  { value: 'H', label: 'H - Direktorat' },
  { value: 'I', label: 'I - Fakultas' },
  { value: 'J', label: 'J - Dosen Tetap Yayasan' },
  { value: 'K', label: 'K - Dosen Tetap Kopertis' },
  { value: 'L', label: 'L - Mahasiswa' },
  { value: 'M', label: 'M - Pegawai, Karyawan & Satpam' },
  { value: 'N', label: 'N - Prodi' }
];

const INSTANSI_PENERBIT_OPTIONS = [
  { value: 'R.UnHar', label: 'R.UnHar - Rektorat' },
  { value: 'SPMI.UnHar', label: 'SPMI.UnHar - Satuan Penjaminan Mutu Internal' },
  { value: 'LPWPI.UnHar', label: 'LPWPI.UnHar - Lembaga Penjaminan Mutu Internal' },
  { value: 'LPPJ.UnHar', label: 'LPPJ.UnHar - Lembaga Penelitian & Publikasi Jurnal' },
  { value: 'LPPM.UnHar', label: 'LPPM.UnHar - Lembaga Pengabdian pada Masyarakat' },
  { value: 'FEB.UnHar', label: 'FEB.UnHar - Fakultas Ekonomi Bisnis' },
  { value: 'FTK.UnHar', label: 'FTK.UnHar - Fakultas Teknik dan Komputer' },
  { value: 'FBK.UnHar', label: 'FBK.UnHar - Fakultas Bahasa dan Komunikasi' },
  { value: 'FH.UnHar', label: 'FH.UnHar - Fakultas Hukum' }
];

export default function App() {
  // Auth States
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [isSignUp, setIsSignUp] = useState(false);
  const [authError, setAuthError] = useState('');
  const [authSuccessMsg, setAuthSuccessMsg] = useState('');

  // App Layout States
  const [activeTab, setActiveTab] = useState('dashboard');
  const [searchQuery, setSearchQuery] = useState('');

  // Data Lists
  const [suratMasukList, setSuratMasukList] = useState([]);
  const [suratKeluarList, setSuratKeluarList] = useState([]);
  const [loadingData, setLoadingData] = useState(false);

  // Google Drive Credentials & Token
  const [googleClientId, setGoogleClientId] = useState(import.meta.env.VITE_GOOGLE_CLIENT_ID || '');
  const [driveToken, setDriveToken] = useState(null);
  const [driveTokenExpires, setDriveTokenExpires] = useState(null);

  // Weekly Backup Reminders
  const [backupReminderFormat, setBackupReminderFormat] = useState('both');
  const [lastBackupDate, setLastBackupDate] = useState(null);

  // Add Entry Modal States
  const [showAddModal, setShowAddModal] = useState(false);
  const [modalType, setModalType] = useState('masuk'); // 'masuk' or 'keluar'
  const [submittingEntry, setSubmittingEntry] = useState(false);

  // File Upload State
  const [selectedFile, setSelectedFile] = useState(null);
  const [uploadProgress, setUploadProgress] = useState(''); // 'idle', 'uploading', 'success', 'error'

  // Generated Letter fields
  const [genNomorSurat, setGenNomorSurat] = useState('001');
  const [genJenisSurat, setGenJenisSurat] = useState('O');
  const [genPenunjukanSurat, setGenPenunjukanSurat] = useState('A');
  const [genTanggal, setGenTanggal] = useState(new Date().toISOString().split('T')[0]);
  const [genInstansiPenerbit, setGenInstansiPenerbit] = useState('FTK.UnHar');

  // Database Entry Fields
  const [pengirim, setPengirim] = useState('');
  const [nomorTanggal, setNomorTanggal] = useState('');
  const [isiRingkas, setIsiRingkas] = useState('');
  const [agendaBerikut, setAgendaBerikut] = useState('');
  const [keterangan, setKeterangan] = useState('');
  const [alamatTanggal, setAlamatTanggal] = useState('');

  // Bulk Delete States
  const [isDeleteMode, setIsDeleteMode] = useState(false);
  const [selectedDeleteIds, setSelectedDeleteIds] = useState(new Set());

  // Initial Auth hook
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setAuthLoading(false);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });

    // Check Google Drive cache
    const cachedToken = localStorage.getItem('drive_token');
    const cachedExpires = localStorage.getItem('drive_token_expires');
    if (cachedToken && cachedExpires && Date.now() < Number(cachedExpires)) {
      setDriveToken(cachedToken);
      setDriveTokenExpires(Number(cachedExpires));
    }

    // Check Backup Settings & Date
    const cachedBackupDate = localStorage.getItem('last_backup_date');
    if (cachedBackupDate) {
      setLastBackupDate(Number(cachedBackupDate));
    }
    const cachedFormat = localStorage.getItem('backup_reminder_format');
    if (cachedFormat) {
      setBackupReminderFormat(cachedFormat);
    }

    return () => subscription.unsubscribe();
  }, []);

  // Fetch letters when session is active
  useEffect(() => {
    if (session) {
      fetchLetters();
    }
  }, [session]);

  const fetchLetters = async () => {
    setLoadingData(true);
    try {
      // Fetch Surat Masuk
      const { data: masukData, error: masukError } = await supabase
        .from('surat_masuk')
        .select('*')
        .order('created_at', { ascending: false });

      if (masukError) throw masukError;
      setSuratMasukList(masukData || []);

      // Fetch Surat Keluar
      const { data: keluarData, error: keluarError } = await supabase
        .from('surat_keluar')
        .select('*')
        .order('created_at', { ascending: false });

      if (keluarError) throw keluarError;
      setSuratKeluarList(keluarData || []);
    } catch (err) {
      console.error('Error fetching data:', err.message);
    } finally {
      setLoadingData(false);
    }
  };

  // Bulk Selection Helpers
  const toggleSelectRow = (id) => {
    const nextSet = new Set(selectedDeleteIds);
    if (nextSet.has(id)) {
      nextSet.delete(id);
    } else {
      nextSet.add(id);
    }
    setSelectedDeleteIds(nextSet);
  };

  const toggleSelectAll = (list) => {
    const allIds = list.map(item => item.id);
    const hasAllSelected = allIds.every(id => selectedDeleteIds.has(id));

    const nextSet = new Set(selectedDeleteIds);
    if (hasAllSelected) {
      allIds.forEach(id => nextSet.delete(id));
    } else {
      allIds.forEach(id => nextSet.add(id));
    }
    setSelectedDeleteIds(nextSet);
  };

  const handleBulkDelete = async (table) => {
    const count = selectedDeleteIds.size;
    if (count === 0) return;

    if (window.confirm(`Apakah Anda yakin ingin menghapus ${count} data surat yang terpilih?`)) {
      setLoadingData(true);
      try {
        const { error } = await supabase
          .from(table)
          .delete()
          .in('id', Array.from(selectedDeleteIds));

        if (error) throw error;

        alert(`Berhasil menghapus ${count} data surat.`);
        fetchLetters();
        setIsDeleteMode(false);
        setSelectedDeleteIds(new Set());
      } catch (err) {
        console.error(err);
        alert('Gagal menghapus data: ' + err.message);
      } finally {
        setLoadingData(false);
      }
    }
  };

  // Auth Operations
  const handleAuth = async (e) => {
    e.preventDefault();
    setAuthError('');
    setAuthSuccessMsg('');
    if (!authEmail || !authPassword) {
      setAuthError('Email and Password are required.');
      return;
    }

    try {
      if (isSignUp) {
        if (authPassword.length < 6) {
          setAuthError('Password must be at least 6 characters.');
          return;
        }
        const { error } = await supabase.auth.signUp({
          email: authEmail,
          password: authPassword
        });
        if (error) throw error;
        setAuthSuccessMsg('Sign up successful! Please check your email for confirmation, or login.');
      } else {
        const { error } = await supabase.auth.signInWithPassword({
          email: authEmail,
          password: authPassword
        });
        if (error) throw error;
      }
    } catch (err) {
      setAuthError(err.message);
    }
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setDriveToken(null);
    setDriveTokenExpires(null);
    localStorage.removeItem('drive_token');
    localStorage.removeItem('drive_token_expires');
  };

  // Google Drive Connect
  const connectGoogleDrive = () => {
    if (!googleClientId) {
      alert('Please configure VITE_GOOGLE_CLIENT_ID in your settings or .env file.');
      return;
    }

    if (!window.google) {
      alert('Google API library not loaded. Check your internet connection.');
      return;
    }

    try {
      const client = window.google.accounts.oauth2.initTokenClient({
        client_id: googleClientId,
        scope: 'https://www.googleapis.com/auth/drive.file',
        callback: (response) => {
          if (response.error) {
            console.error('Google OAuth Error:', response.error);
            alert(`Authentication failed: ${response.error_description || response.error}`);
            return;
          }
          setDriveToken(response.access_token);
          const expiresAt = Date.now() + (parseInt(response.expires_in) * 1000);
          setDriveTokenExpires(expiresAt);
          localStorage.setItem('drive_token', response.access_token);
          localStorage.setItem('drive_token_expires', expiresAt.toString());
        }
      });
      client.requestAccessToken();
    } catch (err) {
      console.error(err);
      alert('Error initiating Google connection: ' + err.message);
    }
  };

  const disconnectGoogleDrive = () => {
    setDriveToken(null);
    setDriveTokenExpires(null);
    localStorage.removeItem('drive_token');
    localStorage.removeItem('drive_token_expires');
  };

  // Helper for Roman numerals
  const toRoman = (num) => {
    const romanMap = {
      1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI',
      7: 'VII', 8: 'VIII', 9: 'IX', 10: 'X', 11: 'XI', 12: 'XII'
    };
    return romanMap[num] || '';
  };

  // Generated Letter formatting
  // Nomor Surat (XXX) / Jenis Surat (XX)-Penunjukan Surat (X) / Bulan Romawi (XX) / Instansi Penerbit (X.XXXX) / Tahun (XXXX)
  // Final: XXX/XX-X/XX/X.XXXX/XXXX
  const computedLetterNumber = () => {
    const paddedNomor = String(genNomorSurat).padStart(3, '0');
    const dateObj = genTanggal ? new Date(genTanggal) : new Date();
    const monthVal = isNaN(dateObj.getTime()) ? new Date().getMonth() + 1 : dateObj.getMonth() + 1;
    const yearVal = isNaN(dateObj.getTime()) ? new Date().getFullYear() : dateObj.getFullYear();
    const bulanRomawi = toRoman(monthVal);
    return `${paddedNomor}/${genJenisSurat}-${genPenunjukanSurat}/${bulanRomawi}/${genInstansiPenerbit}/${yearVal}`;
  };

  // Handle Adding Entry
  const handleAddEntrySubmit = async (e) => {
    e.preventDefault();
    setSubmittingEntry(true);
    setUploadProgress('');

    let uploadedFileName = null;
    let driveFileId = null;

    try {
      // 1. Check file upload first (if file is selected)
      if (selectedFile) {
        if (!driveToken || Date.now() >= driveTokenExpires) {
          throw new Error('Google Drive is not connected or connection expired. Please reconnect in settings.');
        }

        setUploadProgress('Uploading...');
        // Generate random 16 character name preserving extension
        const randomName = generateRandom16CharFilename(selectedFile.name);

        const uploadResult = await uploadFileToDrive(driveToken, selectedFile, randomName);
        uploadedFileName = uploadResult.name;
        driveFileId = uploadResult.id;
        setUploadProgress('Success!');
      }

      // 2. Submit to Supabase
      const generatedNo = computedLetterNumber();

      if (modalType === 'masuk') {
        // Automatically calculate next nomor berturut
        const nextNomorBerturut = (suratMasukList.length > 0)
          ? Math.max(...suratMasukList.map(item => item.nomor_berturut || 0)) + 1
          : 1;

        const { error } = await supabase.from('surat_masuk').insert([{
          nomor_berturut: nextNomorBerturut,
          pengirim: pengirim,
          nomor_tanggal: nomorTanggal,
          isi_ringkas: isiRingkas,
          agenda_berikut: agendaBerikut,
          keterangan: keterangan,
          nomor_surat_generated: generatedNo,
          uploaded_files: uploadedFileName,
          google_drive_file_id: driveFileId,
          user_id: session.user.id
        }]);

        if (error) throw error;
      } else {
        const { error } = await supabase.from('surat_keluar').insert([{
          isi_ringkas: isiRingkas,
          alamat_tanggal: alamatTanggal,
          agenda_berikut: agendaBerikut,
          keterangan: keterangan,
          nomor_surat_generated: generatedNo,
          uploaded_files: uploadedFileName,
          google_drive_file_id: driveFileId,
          user_id: session.user.id
        }]);

        if (error) throw error;
      }

      // Refresh list, close modal, and reset form
      await fetchLetters();
      setShowAddModal(false);
      resetForm();
    } catch (err) {
      console.error(err);
      setUploadProgress('Error!');
      alert('Error inserting entry: ' + err.message);
    } finally {
      setSubmittingEntry(false);
    }
  };

  const resetForm = () => {
    setSelectedFile(null);
    setUploadProgress('');
    setPengirim('');
    setNomorTanggal('');
    setIsiRingkas('');
    setAgendaBerikut('');
    setKeterangan('');
    setAlamatTanggal('');
    setGenNomorSurat('001');
    setGenJenisSurat('O');
    setGenPenunjukanSurat('A');
    setGenTanggal(new Date().toISOString().split('T')[0]);
    setGenInstansiPenerbit('FTK.UnHar');
  };

  // Export Utilities
  const handleExport = (format) => {
    // Generate filename based on date
    const dateStr = new Date().toISOString().split('T')[0];

    // Choose dataset based on active tab
    const isMasuk = activeTab === 'surat_masuk' || (activeTab === 'dashboard' && suratMasukList.length > 0);
    const data = isMasuk ? suratMasukList : suratKeluarList;
    const name = isMasuk ? 'Surat_Masuk' : 'Surat_Keluar';

    if (data.length === 0) {
      alert('Tidak ada data untuk diekspor!');
      return;
    }

    // Format data for export
    const formattedData = data.map((item, index) => {
      if (isMasuk) {
        return {
          'No. Urut': index + 1,
          'Nomor Urut Berturut': item.nomor_berturut,
          'Nomor Surat (Generated)': item.nomor_surat_generated,
          'Si Pengirim': item.pengirim,
          'Nomor & Tanggal dari Instansi Pengirim': item.nomor_tanggal,
          'Isi Ringkas': item.isi_ringkas,
          'Hubungan Agenda Berikut': item.agenda_berikut,
          'Keterangan': item.keterangan,
          'File Google Drive ID': item.google_drive_file_id || '',
          'Nama File': item.uploaded_files || '',
          'Tanggal Dibuat': new Date(item.created_at).toLocaleString()
        };
      } else {
        return {
          'No. Urut': index + 1,
          'Nomor Surat (Generated)': item.nomor_surat_generated,
          'Isi Ringkas': item.isi_ringkas,
          'Alamat & Tanggal': item.alamat_tanggal,
          'Hubungan Agenda Berikut': item.agenda_berikut,
          'Keterangan': item.keterangan,
          'File Google Drive ID': item.google_drive_file_id || '',
          'Nama File': item.uploaded_files || '',
          'Tanggal Dibuat': new Date(item.created_at).toLocaleString()
        };
      }
    });

    if (format === 'csv') {
      const csvContent = convertToCSV(formattedData);
      downloadFile(csvContent, `${name}_${dateStr}.csv`, 'text/csv;charset=utf-8;');
    } else {
      const worksheet = XLSX.utils.json_to_sheet(formattedData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
      XLSX.writeFile(workbook, `${name}_${dateStr}.xlsx`);
    }

    // Update last backup date
    const now = Date.now();
    setLastBackupDate(now);
    localStorage.setItem('last_backup_date', now.toString());
  };

  const convertToCSV = (objArray) => {
    const array = typeof objArray !== 'object' ? JSON.parse(objArray) : objArray;
    let str = '';

    // Headers
    const headers = Object.keys(array[0]);
    str += headers.join(',') + '\r\n';

    // Rows
    for (let i = 0; i < array.length; i++) {
      let line = '';
      for (const index in array[i]) {
        if (line !== '') line += ',';
        // Handle values containing commas
        const val = array[i][index] !== null ? String(array[i][index]) : '';
        line += `"${val.replace(/"/g, '""')}"`;
      }
      str += line + '\r\n';
    }
    return str;
  };

  const downloadFile = (content, fileName, contentType) => {
    const blob = new Blob([content], { type: contentType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Helper: check backup warning state
  const isBackupWarningActive = () => {
    if (!lastBackupDate) return true;
    const sevenDaysInMs = 7 * 24 * 60 * 60 * 1000;
    return (Date.now() - lastBackupDate) > sevenDaysInMs;
  };

  const updateBackupFormat = (format) => {
    setBackupReminderFormat(format);
    localStorage.setItem('backup_reminder_format', format);
  };

  const saveSettings = (e) => {
    e.preventDefault();
    localStorage.setItem('google_client_id_config', googleClientId);
    alert('Settings saved successfully!');
  };

  // Filter lists based on search
  const filteredSuratMasuk = suratMasukList.filter(item =>
    (item.nomor_surat_generated && item.nomor_surat_generated.toLowerCase().includes(searchQuery.toLowerCase())) ||
    (item.pengirim && item.pengirim.toLowerCase().includes(searchQuery.toLowerCase())) ||
    (item.isi_ringkas && item.isi_ringkas.toLowerCase().includes(searchQuery.toLowerCase())) ||
    (item.keterangan && item.keterangan.toLowerCase().includes(searchQuery.toLowerCase()))
  );

  const filteredSuratKeluar = suratKeluarList.filter(item =>
    (item.nomor_surat_generated && item.nomor_surat_generated.toLowerCase().includes(searchQuery.toLowerCase())) ||
    (item.isi_ringkas && item.isi_ringkas.toLowerCase().includes(searchQuery.toLowerCase())) ||
    (item.alamat_tanggal && item.alamat_tanggal.toLowerCase().includes(searchQuery.toLowerCase())) ||
    (item.keterangan && item.keterangan.toLowerCase().includes(searchQuery.toLowerCase()))
  );

  // Render Auth screen if not logged in
  if (authLoading) {
    return (
      <div className="auth-wrapper">
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '16px' }}>
          <Loader2 className="animate-spin" size={48} color="#4F46E5" />
          <p style={{ color: '#64748B', fontWeight: 500 }}>Memuat aplikasi...</p>
        </div>
      </div>
    );
  }

  if (!session) {
    return (
      <div className="auth-wrapper">
        <div className="auth-card">
          <div className="auth-logo">
            <Layers color="#4F46E5" size={28} />
            <span>SIPA FTK UnHar</span>
          </div>
          <p className="auth-subtitle">Sistem Informasi Pengelolaan Arsip Surat FTK Universitas Harapan Medan</p>

          <form onSubmit={handleAuth} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            {authError && (
              <div style={{ padding: '12px', background: 'var(--error-light)', border: '1px solid var(--error)', color: 'var(--error)', borderRadius: 'var(--border-radius-sm)', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <AlertTriangle size={16} />
                <span>{authError}</span>
              </div>
            )}

            {authSuccessMsg && (
              <div style={{ padding: '12px', background: 'var(--success-light)', border: '1px solid var(--success)', color: 'var(--success)', borderRadius: 'var(--border-radius-sm)', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <CheckCircle size={16} />
                <span>{authSuccessMsg}</span>
              </div>
            )}

            <div className="form-group">
              <label>Alamat Email</label>
              <input
                type="email"
                className="input-control"
                placeholder="nama@unhar.ac.id"
                value={authEmail}
                onChange={e => setAuthEmail(e.target.value)}
                required
              />
            </div>

            <div className="form-group">
              <label>Password</label>
              <input
                type="password"
                className="input-control"
                placeholder="••••••••"
                value={authPassword}
                onChange={e => setAuthPassword(e.target.value)}
                required
              />
            </div>

            <button type="submit" className="btn btn-primary" style={{ width: '100%', marginTop: '8px', padding: '12px' }}>
              {isSignUp ? 'Daftar Akun Baru' : 'Masuk ke Sistem'}
            </button>

            <button
              type="button"
              className="btn-text"
              style={{ fontSize: '13px', textAlign: 'center', marginTop: '4px' }}
              onClick={() => {
                setIsSignUp(!isSignUp);
                setAuthError('');
                setAuthSuccessMsg('');
              }}
            >
              {isSignUp ? 'Sudah punya akun? Masuk' : 'Belum punya akun? Daftar Baru'}
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="app-container">
      {/* Sidebar navigation */}
      <aside className="sidebar">
        <div className="sidebar-logo">
          <Layers color="#4F46E5" size={28} />
          <h1>SIPA FTK UnHar</h1>
        </div>

        <nav className="sidebar-nav">
          <button
            className={`nav-item ${activeTab === 'dashboard' ? 'active' : ''}`}
            onClick={() => { setActiveTab('dashboard'); setSearchQuery(''); setIsDeleteMode(false); setSelectedDeleteIds(new Set()); }}
          >
            <Database size={20} />
            <span>Dashboard</span>
          </button>

          <button
            className={`nav-item ${activeTab === 'surat_masuk' ? 'active' : ''}`}
            onClick={() => { setActiveTab('surat_masuk'); setSearchQuery(''); setIsDeleteMode(false); setSelectedDeleteIds(new Set()); }}
          >
            <FileText size={20} />
            <span>Surat Masuk</span>
          </button>

          <button
            className={`nav-item ${activeTab === 'surat_keluar' ? 'active' : ''}`}
            onClick={() => { setActiveTab('surat_keluar'); setSearchQuery(''); setIsDeleteMode(false); setSelectedDeleteIds(new Set()); }}
          >
            <FileText size={20} />
            <span>Surat Keluar</span>
          </button>

          <button
            className={`nav-item ${activeTab === 'settings' ? 'active' : ''}`}
            onClick={() => { setActiveTab('settings'); setSearchQuery(''); setIsDeleteMode(false); setSelectedDeleteIds(new Set()); }}
          >
            <Settings size={20} />
            <span>Pengaturan</span>
          </button>
        </nav>

        <footer className="sidebar-footer">
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{ padding: '8px', background: 'var(--primary-light)', borderRadius: 'var(--border-radius-sm)' }}>
              <User size={18} color="var(--primary)" />
            </div>
            <div className="user-badge">
              <span className="user-email">{session.user.email}</span>
              <span className="user-role">Administrator FTK</span>
            </div>
          </div>
          <button onClick={handleLogout} className="btn btn-secondary btn-sm" style={{ width: '100%' }}>
            <LogOut size={16} />
            <span>Keluar</span>
          </button>
        </footer>
      </aside>

      {/* Main Content Area */}
      <main className="main-content">

        {/* Weekly backup warning toast/banner */}
        {isBackupWarningActive() && activeTab !== 'settings' && (
          <div className="alert-banner warning">
            <div className="alert-content">
              <AlertTriangle size={20} />
              <div>
                <strong>Pengingat Cadangan Mingguan!</strong>
                <p style={{ margin: 0, fontSize: '13px' }}>
                  Anda belum mengekspor/mencadangkan daftar arsip surat minggu ini. Format pilihan Anda: <strong>{backupReminderFormat.toUpperCase()}</strong>.
                </p>
              </div>
            </div>
            <div className="alert-actions">
              {(backupReminderFormat === 'csv' || backupReminderFormat === 'both') && (
                <button onClick={() => handleExport('csv')} className="btn btn-secondary btn-sm" style={{ padding: '6px 12px', background: '#ffffff', color: '#92400e' }}>
                  <Download size={14} />
                  <span>Unduh CSV</span>
                </button>
              )}
              {(backupReminderFormat === 'xlsx' || backupReminderFormat === 'both') && (
                <button onClick={() => handleExport('xlsx')} className="btn btn-primary btn-sm" style={{ padding: '6px 12px', background: '#92400e', color: '#ffffff' }}>
                  <Download size={14} />
                  <span>Unduh Excel</span>
                </button>
              )}
            </div>
          </div>
        )}

        {/* Dashboard Tab */}
        {activeTab === 'dashboard' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '32px' }}>
            <div className="header-bar">
              <div>
                <h2>Dashboard Administrasi</h2>
                <p style={{ color: 'var(--text-muted)' }}>Fakultas Teknik dan Komputer Universitas Harapan Medan</p>
              </div>
              <div style={{ display: 'flex', gap: '12px' }}>
                <button className="btn btn-primary" onClick={() => { setModalType('masuk'); setShowAddModal(true); }}>
                  <Plus size={16} />
                  <span>Surat Masuk</span>
                </button>
                <button className="btn btn-accent" onClick={() => { setModalType('keluar'); setShowAddModal(true); }}>
                  <Plus size={16} />
                  <span>Surat Keluar</span>
                </button>
              </div>
            </div>

            <div className="dashboard-grid">
              <div className="stat-card">
                <div className="stat-header">
                  <span>TOTAL SURAT MASUK</span>
                  <div className="stat-icon-wrapper">
                    <FileText size={18} />
                  </div>
                </div>
                <span className="stat-value">{suratMasukList.length}</span>
                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>Arsip Surat Masuk terdaftar</span>
              </div>

              <div className="stat-card">
                <div className="stat-header">
                  <span>TOTAL SURAT KELUAR</span>
                  <div className="stat-icon-wrapper" style={{ color: 'var(--accent)', background: 'var(--accent-light)' }}>
                    <FileText size={18} />
                  </div>
                </div>
                <span className="stat-value" style={{ color: 'var(--accent)' }}>{suratKeluarList.length}</span>
                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>Arsip Surat Keluar terdaftar</span>
              </div>

              <div className="stat-card">
                <div className="stat-header">
                  <span>GOOGLE DRIVE</span>
                  <div className="stat-icon-wrapper" style={{
                    color: driveToken ? 'var(--success)' : 'var(--error)',
                    background: driveToken ? 'var(--success-light)' : 'var(--error-light)'
                  }}>
                    <Cloud size={18} />
                  </div>
                </div>
                <span className="stat-value" style={{ fontSize: '20px', fontWeight: 700, color: driveToken ? 'var(--success)' : 'var(--error)' }}>
                  {driveToken ? 'Terkoneksi' : 'Terputus'}
                </span>
                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                  {driveToken ? 'Drive API siap untuk upload file' : 'Hubungkan Drive di tab pengaturan'}
                </span>
              </div>
            </div>

            {/* Quick overview of latest letters */}
            <div className="card-table-container">
              <div className="table-toolbar">
                <h3 style={{ fontSize: '16px' }}>Arsip Surat Terbaru</h3>
                <div style={{ display: 'flex', gap: '12px' }}>
                  <button onClick={() => setActiveTab('surat_masuk')} className="btn btn-secondary btn-sm">
                    <span>Lihat Semua Surat Masuk</span>
                    <ChevronRight size={14} />
                  </button>
                  <button onClick={() => setActiveTab('surat_keluar')} className="btn btn-secondary btn-sm">
                    <span>Lihat Semua Surat Keluar</span>
                    <ChevronRight size={14} />
                  </button>
                </div>
              </div>
              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>Jenis</th>
                      <th>Nomor Surat Generated</th>
                      <th>Isi Ringkas / Keterangan</th>
                      <th>Tanggal Masuk/Keluar</th>
                    </tr>
                  </thead>
                  <tbody>
                    {suratMasukList.slice(0, 3).map(item => (
                      <tr key={item.id}>
                        <td><span style={{ color: '#10b981', fontWeight: 600 }}>Masuk</span></td>
                        <td><span className="letter-badge">{item.nomor_surat_generated}</span></td>
                        <td>
                          <div><strong>{item.isi_ringkas || '-'}</strong></div>
                          <div style={{ fontSize: '12px', color: 'var(--text-muted)' }}>Pengirim: {item.pengirim || '-'}</div>
                        </td>
                        <td>{new Date(item.created_at).toLocaleDateString()}</td>
                      </tr>
                    ))}
                    {suratKeluarList.slice(0, 3).map(item => (
                      <tr key={item.id}>
                        <td><span style={{ color: 'var(--accent)', fontWeight: 600 }}>Keluar</span></td>
                        <td><span className="letter-badge" style={{ background: 'var(--accent-light)', color: 'var(--accent)' }}>{item.nomor_surat_generated}</span></td>
                        <td>
                          <div><strong>{item.isi_ringkas || '-'}</strong></div>
                          <div style={{ fontSize: '12px', color: 'var(--text-muted)' }}>Alamat: {item.alamat_tanggal || '-'}</div>
                        </td>
                        <td>{new Date(item.created_at).toLocaleDateString()}</td>
                      </tr>
                    ))}
                    {suratMasukList.length === 0 && suratKeluarList.length === 0 && (
                      <tr>
                        <td colSpan="4" className="empty-state">
                          <Info size={32} />
                          <span>Belum ada data arsip surat terdaftar.</span>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Surat Masuk Tab */}
        {activeTab === 'surat_masuk' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            <div className="header-bar">
              <div>
                <h2>Arsip Surat Masuk</h2>
                <p style={{ color: 'var(--text-muted)' }}>Daftar lengkap berkas surat masuk administrasi FTK</p>
              </div>
              <div style={{ display: 'flex', gap: '12px' }}>
                {isDeleteMode ? (
                  <>
                    <button
                      className="btn btn-secondary"
                      onClick={() => { setIsDeleteMode(false); setSelectedDeleteIds(new Set()); }}
                    >
                      Batalkan Hapus Surat
                    </button>
                    {selectedDeleteIds.size > 0 && (
                      <button
                        className="btn btn-primary"
                        style={{ backgroundColor: 'var(--error)', borderColor: 'var(--error)' }}
                        onClick={() => handleBulkDelete('surat_masuk')}
                      >
                        Confirm Hapus ({selectedDeleteIds.size})
                      </button>
                    )}
                  </>
                ) : (
                  <>
                    <button className="btn btn-primary" onClick={() => { setModalType('masuk'); setShowAddModal(true); }}>
                      <Plus size={16} />
                      <span>Tambah Surat Masuk</span>
                    </button>
                    <button
                      className="btn btn-secondary"
                      style={{ color: 'var(--error)', borderColor: 'var(--error)' }}
                      onClick={() => { setIsDeleteMode(true); setSelectedDeleteIds(new Set()); }}
                      disabled={suratMasukList.length === 0}
                    >
                      Hapus Surat
                    </button>
                  </>
                )}
              </div>
            </div>

            <div className="card-table-container">
              <div className="table-toolbar">
                <div className="search-box">
                  <Search size={18} />
                  <input
                    type="text"
                    className="input-control"
                    placeholder="Cari surat masuk..."
                    value={searchQuery}
                    onChange={e => setSearchQuery(e.target.value)}
                  />
                </div>
                <div className="toolbar-actions">
                  <button onClick={() => handleExport('csv')} className="btn btn-secondary btn-sm">
                    <Download size={14} />
                    <span>Ekspor CSV</span>
                  </button>
                  <button onClick={() => handleExport('xlsx')} className="btn btn-secondary btn-sm">
                    <Download size={14} />
                    <span>Ekspor Excel (.xlsx)</span>
                  </button>
                </div>
              </div>

              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      {isDeleteMode && (
                        <th style={{ width: '40px' }}>
                          <input
                            type="checkbox"
                            checked={filteredSuratMasuk.length > 0 && filteredSuratMasuk.every(item => selectedDeleteIds.has(item.id))}
                            onChange={() => toggleSelectAll(filteredSuratMasuk)}
                          />
                        </th>
                      )}
                      <th>No. Urut</th>
                      <th>Nomor Generated</th>
                      <th>Nomor Berturut</th>
                      <th>Pengirim</th>
                      <th>Nomor & Tanggal dari Instansi Pengirim</th>
                      <th>Isi Ringkas</th>
                      <th>Agenda Berikut</th>
                      <th>Keterangan</th>
                      <th>Berkas</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSuratMasuk.map((item, idx) => (
                      <tr key={item.id} className={selectedDeleteIds.has(item.id) ? 'row-selected' : ''}>
                        {isDeleteMode && (
                          <td>
                            <input
                              type="checkbox"
                              checked={selectedDeleteIds.has(item.id)}
                              onChange={() => toggleSelectRow(item.id)}
                            />
                          </td>
                        )}
                        <td>{idx + 1}</td>
                        <td><span className="letter-badge">{item.nomor_surat_generated}</span></td>
                        <td style={{ fontWeight: 600 }}>{item.nomor_berturut}</td>
                        <td>{item.pengirim || '-'}</td>
                        <td>{item.nomor_tanggal}</td>
                        <td>{item.isi_ringkas || '-'}</td>
                        <td>{item.agenda_berikut || '-'}</td>
                        <td>{item.keterangan || '-'}</td>
                        <td>
                          {item.google_drive_file_id ? (
                            <a
                              href={`https://drive.google.com/file/d/${item.google_drive_file_id}/view?usp=drivesdk`}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="file-link"
                            >
                              <Cloud size={14} />
                              <span style={{ fontSize: '11px' }}>{item.uploaded_files}</span>
                            </a>
                          ) : '-'}
                        </td>
                      </tr>
                    ))}
                    {filteredSuratMasuk.length === 0 && (
                      <tr>
                        <td colSpan="9" className="empty-state">
                          <Search size={32} />
                          <span>Data surat masuk tidak ditemukan atau kosong.</span>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Surat Keluar Tab */}
        {activeTab === 'surat_keluar' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            <div className="header-bar">
              <div>
                <h2>Arsip Surat Keluar</h2>
                <p style={{ color: 'var(--text-muted)' }}>Daftar lengkap berkas surat keluar administrasi FTK</p>
              </div>
              <div style={{ display: 'flex', gap: '12px' }}>
                {isDeleteMode ? (
                  <>
                    <button
                      className="btn btn-secondary"
                      onClick={() => { setIsDeleteMode(false); setSelectedDeleteIds(new Set()); }}
                    >
                      Batalkan Hapus Surat
                    </button>
                    {selectedDeleteIds.size > 0 && (
                      <button
                        className="btn btn-accent"
                        style={{ backgroundColor: 'var(--error)', borderColor: 'var(--error)' }}
                        onClick={() => handleBulkDelete('surat_keluar')}
                      >
                        Confirm Hapus ({selectedDeleteIds.size})
                      </button>
                    )}
                  </>
                ) : (
                  <>
                    <button className="btn btn-accent" onClick={() => { setModalType('keluar'); setShowAddModal(true); }}>
                      <Plus size={16} />
                      <span>Tambah Surat Keluar</span>
                    </button>
                    <button
                      className="btn btn-secondary"
                      style={{ color: 'var(--error)', borderColor: 'var(--error)' }}
                      onClick={() => { setIsDeleteMode(true); setSelectedDeleteIds(new Set()); }}
                      disabled={suratKeluarList.length === 0}
                    >
                      Hapus Surat
                    </button>
                  </>
                )}
              </div>
            </div>

            <div className="card-table-container">
              <div className="table-toolbar">
                <div className="search-box">
                  <Search size={18} />
                  <input
                    type="text"
                    className="input-control"
                    placeholder="Cari surat keluar..."
                    value={searchQuery}
                    onChange={e => setSearchQuery(e.target.value)}
                  />
                </div>
                <div className="toolbar-actions">
                  <button onClick={() => handleExport('csv')} className="btn btn-secondary btn-sm">
                    <Download size={14} />
                    <span>Ekspor CSV</span>
                  </button>
                  <button onClick={() => handleExport('xlsx')} className="btn btn-secondary btn-sm">
                    <Download size={14} />
                    <span>Ekspor Excel (.xlsx)</span>
                  </button>
                </div>
              </div>

              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      {isDeleteMode && (
                        <th style={{ width: '40px' }}>
                          <input
                            type="checkbox"
                            checked={filteredSuratKeluar.length > 0 && filteredSuratKeluar.every(item => selectedDeleteIds.has(item.id))}
                            onChange={() => toggleSelectAll(filteredSuratKeluar)}
                          />
                        </th>
                      )}
                      <th>No. Urut</th>
                      <th>Nomor Generated</th>
                      <th>Isi Ringkas</th>
                      <th>Alamat & Tanggal</th>
                      <th>Agenda Berikut</th>
                      <th>Keterangan</th>
                      <th>Berkas</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSuratKeluar.map((item, idx) => (
                      <tr key={item.id} className={selectedDeleteIds.has(item.id) ? 'row-selected' : ''}>
                        {isDeleteMode && (
                          <td>
                            <input
                              type="checkbox"
                              checked={selectedDeleteIds.has(item.id)}
                              onChange={() => toggleSelectRow(item.id)}
                            />
                          </td>
                        )}
                        <td>{idx + 1}</td>
                        <td><span className="letter-badge" style={{ background: 'var(--accent-light)', color: 'var(--accent)' }}>{item.nomor_surat_generated}</span></td>
                        <td>{item.isi_ringkas || '-'}</td>
                        <td>{item.alamat_tanggal || '-'}</td>
                        <td>{item.agenda_berikut || '-'}</td>
                        <td>{item.keterangan || '-'}</td>
                        <td>
                          {item.google_drive_file_id ? (
                            <a
                              href={`https://drive.google.com/file/d/${item.google_drive_file_id}/view?usp=drivesdk`}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="file-link"
                            >
                              <Cloud size={14} />
                              <span style={{ fontSize: '11px' }}>{item.uploaded_files}</span>
                            </a>
                          ) : '-'}
                        </td>
                      </tr>
                    ))}
                    {filteredSuratKeluar.length === 0 && (
                      <tr>
                        <td colSpan="7" className="empty-state">
                          <Search size={32} />
                          <span>Data surat keluar tidak ditemukan atau kosong.</span>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Settings and Config tab */}
        {activeTab === 'settings' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '32px' }}>
            <div>
              <h2>Google Drive & Pengaturan Setup</h2>
              <p style={{ color: 'var(--text-muted)' }}>Konfigurasi dan integrasi Google Drive pihak ketiga untuk user</p>
            </div>

            <div className="settings-card">
              <h3 style={{ fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Cloud size={20} color="var(--accent)" />
                <span>Koneksi Google Drive</span>
              </h3>

              <div className="connection-status">
                <div className={`status-indicator ${driveToken ? 'connected' : 'disconnected'}`}></div>
                <div>
                  <strong>{driveToken ? 'Google Drive Terkoneksi' : 'Google Drive Terputus'}</strong>
                  <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-muted)' }}>
                    {driveToken
                      ? `Token aktif. Kedaluwarsa pada: ${new Date(driveTokenExpires).toLocaleTimeString()}`
                      : 'Koneksikan Google Drive Anda untuk mengaktifkan fitur upload berkas.'}
                  </p>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '12px' }}>
                {driveToken ? (
                  <button onClick={disconnectGoogleDrive} className="btn btn-secondary">
                    Putuskan Akses
                  </button>
                ) : (
                  <button onClick={connectGoogleDrive} className="btn btn-accent">
                    Hubungkan Google Drive
                  </button>
                )}
              </div>

              <form onSubmit={saveSettings} style={{ borderTop: '1px solid var(--border)', paddingTop: '24px' }}>
                <div className="form-group">
                  <label>Google OAuth Client ID (Konfigurasi Lokal/Dev)</label>
                  <input
                    type="text"
                    className="input-control"
                    placeholder="Masukkan Google Client ID dari Google Cloud Console"
                    value={googleClientId}
                    onChange={e => setGoogleClientId(e.target.value)}
                  />
                  <span style={{ fontSize: '12px', color: 'var(--text-muted)', marginTop: '4px' }}>
                    Diperlukan untuk memunculkan modal login Google Drive client-side. Pastikan domain asal terdaftar di Google Cloud Console.
                  </span>
                </div>
                <button type="submit" className="btn btn-primary btn-sm" style={{ marginTop: '8px' }}>
                  Simpan Client ID
                </button>
              </form>
            </div>

            <div className="settings-card">
              <h3 style={{ fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <AlertTriangle size={20} color="var(--warning)" />
                <span>Pengaturan Backup Mingguan</span>
              </h3>

              <div className="switch-group">
                <div className="switch-info">
                  <span className="switch-title">Pilihan Ekspor Pengingat</span>
                  <span className="switch-desc">Tentukan berkas ekspor mana yang akan diminta di pengingat dashboard</span>
                </div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button
                    onClick={() => updateBackupFormat('csv')}
                    className={`btn btn-sm ${backupReminderFormat === 'csv' ? 'btn-primary' : 'btn-secondary'}`}
                  >
                    CSV Saja
                  </button>
                  <button
                    onClick={() => updateBackupFormat('xlsx')}
                    className={`btn btn-sm ${backupReminderFormat === 'xlsx' ? 'btn-primary' : 'btn-secondary'}`}
                  >
                    Excel (.xlsx) Saja
                  </button>
                  <button
                    onClick={() => updateBackupFormat('both')}
                    className={`btn btn-sm ${backupReminderFormat === 'both' ? 'btn-primary' : 'btn-secondary'}`}
                  >
                    Keduanya
                  </button>
                </div>
              </div>

              <div className="switch-group">
                <div className="switch-info">
                  <span className="switch-title">Riwayat Pencadangan Terakhir</span>
                  <span className="switch-desc">
                    {lastBackupDate
                      ? `Terakhir kali Anda mengunduh list arsip adalah: ${new Date(lastBackupDate).toLocaleString()}`
                      : 'Anda belum pernah melakukan ekspor data di perangkat ini.'}
                  </span>
                </div>
                {lastBackupDate && (
                  <button
                    onClick={() => {
                      localStorage.removeItem('last_backup_date');
                      setLastBackupDate(null);
                    }}
                    className="btn btn-secondary btn-sm"
                  >
                    Reset Status Backup
                  </button>
                )}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Add letter pop-up Modal */}
      {showAddModal && (
        <div className="modal-overlay">
          <div className="modal-container">
            <div className="modal-header">
              <h3>Tambah Data Surat {modalType === 'masuk' ? 'Masuk' : 'Keluar'}</h3>
              <button className="btn-text" onClick={() => { setShowAddModal(false); resetForm(); }}>
                ✕
              </button>
            </div>

            <form onSubmit={handleAddEntrySubmit}>
              <div className="modal-body">

                {/* Generated Letter Preview */}
                <div className="live-preview-box">
                  <span className="live-preview-label">PREVIEW NOMOR SURAT GENERATED</span>
                  <span className="live-preview-value">{computedLetterNumber()}</span>
                </div>

                <div className="form-row">
                  {/* Nomor Surat (001 - 999) */}
                  <div className="form-group">
                    <label>Nomor Surat (001 - 999)</label>
                    <input
                      type="number"
                      min="1"
                      max="999"
                      className="input-control"
                      value={genNomorSurat}
                      onChange={e => setGenNomorSurat(e.target.value)}
                      required
                    />
                  </div>
                  {/* Jenis Surat */}
                  <div className="form-group">
                    <label>Jenis Surat</label>
                    <select
                      className="input-control"
                      value={genJenisSurat}
                      onChange={e => setGenJenisSurat(e.target.value)}
                    >
                      {JENIS_SURAT_OPTIONS.map(opt => (
                        <option key={opt.value} value={opt.value}>{opt.label}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="form-row">
                  {/* Penunjukan Surat */}
                  <div className="form-group">
                    <label>Penunjukan Surat</label>
                    <select
                      className="input-control"
                      value={genPenunjukanSurat}
                      onChange={e => setGenPenunjukanSurat(e.target.value)}
                    >
                      {PENUNJUKAN_SURAT_OPTIONS.map(opt => (
                        <option key={opt.value} value={opt.value}>{opt.label}</option>
                      ))}
                    </select>
                  </div>

                  {/* Bulan Romawi & Tahun Penerbitan replaced with a Date Selector */}
                  <div className="form-group">
                    <label>Tanggal Penerbitan</label>
                    <input
                      type="date"
                      className="input-control"
                      value={genTanggal}
                      onChange={e => setGenTanggal(e.target.value)}
                      required
                    />
                  </div>
                </div>

                <div className="form-row">
                  {/* Instansi Penerbit */}
                  <div className="form-group" style={{ gridColumn: 'span 2' }}>
                    <label>Instansi Penerbit</label>
                    <select
                      className="input-control"
                      value={genInstansiPenerbit}
                      onChange={e => setGenInstansiPenerbit(e.target.value)}
                    >
                      {INSTANSI_PENERBIT_OPTIONS.map(opt => (
                        <option key={opt.value} value={opt.value}>{opt.label}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <hr style={{ margin: '16px 0', borderColor: 'var(--border)' }} />

                {/* Surat Masuk Specific fields */}
                {modalType === 'masuk' ? (
                  <>
                    <div className="form-row">
                      <div className="form-group" style={{ gridColumn: 'span 2' }}>
                        <label>Si Pengirim</label>
                        <input
                          type="text"
                          className="input-control"
                          value={pengirim}
                          onChange={e => setPengirim(e.target.value)}
                          placeholder="cth: LPPM UnHar"
                        />
                      </div>
                    </div>

                    <div className="form-group">
                      <label>Nomor dan Tanggal dari Instansi Pengirim</label>
                      <input
                        type="text"
                        className="input-control"
                        value={nomorTanggal}
                        onChange={e => setNomorTanggal(e.target.value)}
                        placeholder="cth: 025/LPPM/VI/2026 - 15 Juni 2026"
                      />
                    </div>
                  </>
                ) : (
                  /* Surat Keluar Specific fields */
                  <div className="form-group">
                    <label>Alamat dan Tanggal Pengiriman</label>
                    <input
                      type="text"
                      className="input-control"
                      value={alamatTanggal}
                      onChange={e => setAlamatTanggal(e.target.value)}
                      placeholder="cth: Jl. Merdeka No. 10 - 2 Juli 2026"
                    />
                  </div>
                )}

                <div className="form-group">
                  <label>Isi Ringkas</label>
                  <textarea
                    className="input-control"
                    rows="3"
                    value={isiRingkas}
                    onChange={e => setIsiRingkas(e.target.value)}
                    placeholder="Tuliskan ringkasan isi surat..."
                  />
                </div>

                <div className="form-row">
                  <div className="form-group">
                    <label>Hubungan Agenda Berikutnya</label>
                    <input
                      type="text"
                      className="input-control"
                      value={agendaBerikut}
                      onChange={e => setAgendaBerikut(e.target.value)}
                      placeholder="Hubungan dengan nomor agenda berikutnya"
                    />
                  </div>
                  <div className="form-group">
                    <label>Keterangan</label>
                    <input
                      type="text"
                      className="input-control"
                      value={keterangan}
                      onChange={e => setKeterangan(e.target.value)}
                      placeholder="Catatan keterangan tambahan"
                    />
                  </div>
                </div>

                {/* Google Drive upload section */}
                <div style={{ background: '#f8fafc', padding: '16px', borderRadius: 'var(--border-radius-md)', border: '1px dashed var(--border)', marginTop: '8px' }}>
                  <span style={{ fontSize: '13px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
                    <Cloud size={16} color={driveToken ? 'var(--success)' : 'var(--text-muted)'} />
                    <span>Upload File Ke Google Drive Pribadi</span>
                  </span>

                  {driveToken ? (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      <input
                        type="file"
                        onChange={e => setSelectedFile(e.target.files[0])}
                        style={{ fontSize: '13px' }}
                      />
                      <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                        File akan diunggah ke folder: <strong>SIPA FTK UnHar Files</strong> dengan nama acak 16-karakter.
                      </span>
                      {uploadProgress && (
                        <div style={{ fontSize: '12px', fontWeight: 600, color: uploadProgress.includes('Error') ? 'var(--error)' : 'var(--success)' }}>
                          Status upload: {uploadProgress}
                        </div>
                      )}
                    </div>
                  ) : (
                    <div style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                      Google Drive belum terkoneksi. Silakan hubungkan Google Drive di tab Pengaturan terlebih dahulu untuk mengunggah berkas surat.
                    </div>
                  )}
                </div>

              </div>

              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" onClick={() => { setShowAddModal(false); resetForm(); }}>
                  Batal
                </button>
                <button type="submit" className="btn btn-primary" disabled={submittingEntry}>
                  {submittingEntry ? (
                    <>
                      <Loader2 className="animate-spin" size={16} />
                      <span>Menyimpan...</span>
                    </>
                  ) : (
                    <span>Simpan Entri</span>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
