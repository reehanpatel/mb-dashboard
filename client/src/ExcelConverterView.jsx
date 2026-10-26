import React, { useState, useEffect, useRef, useCallback } from 'react';
import api from './api';

// Official Media Buzz Software Terms & Conditions (Strict Single Source of Truth)
const OFFICIAL_SOFTWARE_TERMS = [
  '1. Before confirming any site, please check and confirm its availability.',
  '2. 18% GST will be charged extra, as applicable.',
  '3. Printing charges will be charged extra, as actual.',
  '4. Mounting charges will be charged extra at ₹4 per sq. ft.',
  '5. Payment Terms: Advance / within 30 days from the date of invoice.',
  '6. Artwork/Flex must be provided within 2 days of site confirmation, otherwise the site may be released.',
  "7. Cancellations require a minimum 10 days' notice; immediate cancellations are not accepted."
];

// Target fields corresponding to the 11 columns in Jade Blue 16-10-2026.xlsx master template
const MASTER_TARGET_FIELDS = [
  { key: 'sr_no', label: 'Serial / Site No', masterCol: 'SR NO', required: false, desc: 'Site sequence number (e.g. 1, 2, 3...)' },
  { key: 'city', label: 'City', masterCol: 'City', required: false, desc: 'City name (e.g. Ahmedabad, Surat)' },
  { key: 'location', label: 'Location / Address', masterCol: 'LOCATION', required: true, desc: 'Prime location or billboard address' },
  { key: 'media', label: 'Media Type', masterCol: 'MEDIA', required: false, desc: 'Unipole, Hoarding, Gantry, etc.' },
  { key: 'light', label: 'Illumination', masterCol: 'LIGHT', required: false, desc: 'BL, FL, NL (Backlit, Frontlit, Non-lit)' },
  { key: 'w', label: 'Width (W)', masterCol: 'W', required: false, desc: 'Width in feet' },
  { key: 'h', label: 'Height (H)', masterCol: 'H', required: false, desc: 'Height in feet' },
  { key: 'sq_ft', label: 'Total Sq Ft', masterCol: 'SQ FT', required: false, desc: 'Calculated area in square feet' },
  { key: 'avail', label: 'Availability', masterCol: 'AVAILABLITY', required: false, desc: 'Available or Booked till date' },
  { key: 'rate', label: 'Selling Amount', masterCol: 'Selling Amount', required: false, desc: 'Monthly advertising fee (INR)' },
  { key: 'coords', label: 'Location Coordinates', masterCol: 'Latitude Longitude', required: false, desc: 'GPS Lat, Lng coordinates' }
];

const POPULAR_CITIES = ['Ahmedabad', 'Surat', 'Vadodara', 'Rajkot', 'Gandhinagar', 'Mumbai', 'Pune', 'Delhi NCR'];

export default function ExcelConverterView() {
  // Master template state
  const [templateStatus, setTemplateStatus] = useState({
    loading: true,
    available: false,
    filename: '',
    size: 0,
    masterExcel: null,
    error: null
  });

  // File upload state
  const [file, setFile] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const fileInputRef = useRef(null);

  // Manual configuration
  const [manualCity, setManualCity] = useState('');
  const [customFileName, setCustomFileName] = useState('');

  // Inspection & Column Mapping state
  const [inspecting, setInspecting] = useState(false);
  const [excelInspection, setExcelInspection] = useState(null);
  const [columnMapping, setColumnMapping] = useState({});
  const [selectedSheet, setSelectedSheet] = useState('');
  const [showMappingModal, setShowMappingModal] = useState(false);
  const [showTermsModal, setShowTermsModal] = useState(false);

  // Conversion process
  const [converting, setConverting] = useState(false);
  const [progressStep, setProgressStep] = useState(0);
  const [progressPercent, setProgressPercent] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [conversionResult, setConversionResult] = useState(null);
  const [error, setError] = useState(null);

  // Slide preview & inline editing state
  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const [viewMode, setViewMode] = useState('grid'); // 'grid' | 'slideshow'
  const [filterReviewOnly, setFilterReviewOnly] = useState(false);
  const [editingSlideIndex, setEditingSlideIndex] = useState(null);
  const [editedFields, setEditedFields] = useState({});

  // Fetch master template status on mount
  const checkTemplateStatus = useCallback(async () => {
    try {
      setTemplateStatus(prev => ({ ...prev, loading: true, error: null }));
      const res = await api.get('/excel-converter/status');
      setTemplateStatus({
        loading: false,
        available: !!res.data?.available,
        filename: res.data?.filename || 'MediaBuzz_Master_Template.pptx',
        size: res.data?.size || 0,
        masterExcel: res.data?.masterExcel || {
          available: true,
          filename: 'Jade Blue 16-10-2026.xlsx',
          masterColumns: MASTER_TARGET_FIELDS.map(f => f.masterCol)
        },
        error: null
      });
    } catch (err) {
      setTemplateStatus({
        loading: false,
        available: false,
        filename: 'MediaBuzz_Master_Template.pptx',
        size: 0,
        masterExcel: {
          available: true,
          filename: 'Jade Blue 16-10-2026.xlsx',
          masterColumns: MASTER_TARGET_FIELDS.map(f => f.masterCol)
        },
        error: err.response?.data?.error || err.message || 'Could not verify master template'
      });
    }
  }, []);

  useEffect(() => {
    checkTemplateStatus();
  }, [checkTemplateStatus]);

  // Handle Drag & Drop
  const handleDrag = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  }, []);

  // Inspect uploaded Excel workbook structure
  const inspectExcelWorkbook = async (selectedFile) => {
    try {
      setInspecting(true);
      setError(null);
      const fd = new FormData();
      fd.append('file', selectedFile);
      const res = await api.post('/excel-converter/inspect-excel', fd, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setExcelInspection(res.data);
      if (res.data?.mapping) {
        setColumnMapping(res.data.mapping);
      }
      if (res.data?.activeSheet) {
        setSelectedSheet(res.data.activeSheet);
      }
      // If headers do not match master Jade Blue template, show mapping modal
      if (res.data?.matchesMaster === false) {
        setShowMappingModal(true);
      }
    } catch (err) {
      console.error('Excel inspection error:', err);
      setError('Failed to inspect Excel workbook: ' + (err.response?.data?.message || err.message));
    } finally {
      setInspecting(false);
    }
  };

  const validateAndSetFile = (selectedFile) => {
    setError(null);
    if (!selectedFile) return;
    const name = selectedFile.name || '';
    const lower = name.toLowerCase();
    const isXls = lower.endsWith('.xlsx') || lower.endsWith('.xls');

    if (!isXls) {
      setError('Please select a valid Excel workbook (.xlsx or .xls).');
      return;
    }
    setFile(selectedFile);
    setConversionResult(null);
    const defaultName = name.replace(/\.[^/.]+$/, '') + '_MediaBuzz';
    if (!customFileName) setCustomFileName(defaultName);
    inspectExcelWorkbook(selectedFile);
  };

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      validateAndSetFile(e.dataTransfer.files[0]);
    }
  }, [customFileName]);

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      validateAndSetFile(e.target.files[0]);
    }
  };

  // Convert Excel workflow
  const handleConvert = async () => {
    if (!file || !templateStatus.available || converting) return;

    setError(null);
    setConverting(true);
    setProgressStep(1);
    setProgressPercent(15);
    setProgressMessage('Uploading Excel workbook for conversion...');

    const pInterval = setInterval(() => {
      setProgressPercent(prev => {
        if (prev < 40) return prev + 8;
        if (prev < 75) return prev + 4;
        if (prev < 90) return prev + 1;
        return prev;
      });
    }, 400);

    setTimeout(() => {
      setProgressStep(2);
      setProgressMessage('Parsing billboard rows, specs & mapped columns...');
    }, 1200);

    setTimeout(() => {
      setProgressStep(3);
      setProgressMessage('Populating Media Buzz Master Template PPT and standardized Excel...');
    }, 2800);

    try {
      const fd = new FormData();
      fd.append('file', file);
      if (manualCity.trim()) {
        fd.append('manualCity', manualCity.trim());
      }
      if (customFileName.trim()) {
        fd.append('customName', customFileName.trim());
      }
      if (columnMapping && Object.keys(columnMapping).length > 0) {
        fd.append('columnMapping', JSON.stringify(columnMapping));
      }
      if (selectedSheet) {
        fd.append('sheetName', selectedSheet);
      }

      const res = await api.post('/excel-converter/convert', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 180000
      });

      clearInterval(pInterval);
      setProgressPercent(100);
      setProgressStep(4);
      setProgressMessage('Excel conversion completed successfully!');

      setTimeout(() => {
        setConversionResult(res.data);
        if (res.data?.customName) {
          setCustomFileName(res.data.customName);
        }
        setActiveSlideIndex(0);
        setConverting(false);
      }, 500);

    } catch (err) {
      clearInterval(pInterval);
      setConverting(false);
      setProgressStep(0);
      setProgressPercent(0);
      const msg = err.response?.data?.message || err.message || 'An error occurred during Excel conversion.';
      setError(msg);
    }
  };

  // Download converted presentation (.pptx)
  const handleDownloadPpt = async () => {
    if (!conversionResult?.downloadId) return;
    try {
      const cleanBase = (customFileName || 'Converted_MediaBuzz_Master_Format').trim().replace(/\.pptx$/i, '');
      const cleanName = `${cleanBase}.pptx`;
      const downloadUrl = `/excel-converter/download/${encodeURIComponent(conversionResult.downloadId)}?name=${encodeURIComponent(cleanName)}`;
      const res = await api.get(downloadUrl, { responseType: 'blob' });
      const blob = new Blob([res.data], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = cleanName;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      alert('Failed to download presentation: ' + (err.response?.data?.message || err.message));
    }
  };

  // Download standardized Excel (.xlsx)
  const handleDownloadExcel = async () => {
    if (!conversionResult?.downloadId && !conversionResult?.excelDownloadId) return;
    try {
      const cleanBase = (customFileName || 'Converted_MediaBuzz_Master_Format').trim().replace(/\.xlsx$/i, '').replace(/\.pptx$/i, '');
      const cleanName = `${cleanBase}.xlsx`;
      const downloadId = conversionResult.excelDownloadId || conversionResult.downloadId.replace(/\.pptx$/i, '.xlsx');
      const downloadUrl = `/excel-converter/download-excel/${encodeURIComponent(downloadId)}?name=${encodeURIComponent(cleanName)}`;
      const res = await api.get(downloadUrl, { responseType: 'blob' });
      const blob = new Blob([res.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = cleanName;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      alert('Failed to download Excel file: ' + (err.response?.data?.message || err.message));
    }
  };

  // Reset conversion
  const handleReset = () => {
    setFile(null);
    setExcelInspection(null);
    setColumnMapping({});
    setShowMappingModal(false);
    setConversionResult(null);
    setError(null);
    setConverting(false);
    setProgressStep(0);
    setProgressPercent(0);
    setActiveSlideIndex(0);
    setEditingSlideIndex(null);
    setEditedFields({});
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const formatBytes = (bytes) => {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const slides = conversionResult?.slides || [];
  const filteredSlides = filterReviewOnly
    ? slides.filter(s => s.reviewNeeded)
    : slides;

  const activeSlide = slides[activeSlideIndex] || null;

  return (
    <div className="scooh-ppt-converter-module">
      {/* Page Header */}
      <div className="scooh-pagehead">
        <div>
          <h1 style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <span>Excel Converter</span>
            <span className="scooh-pill active" style={{ fontSize: '11px', textTransform: 'uppercase' }}>
              Jade Blue Master Standard
            </span>
          </h1>
          <p>
            Convert Excel workbooks (<code>.xlsx</code> / <code>.xls</code>) into official Media Buzz Master Presentation format and clean 11-column standardized Master Excel.
            Enter city manually, customize file name, and map non-standard columns with real-time preview.
          </p>
        </div>
        <div className="scooh-headactions">
          {conversionResult && (
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleDownloadPpt}
                id="btn-download-excel-conv-ppt-top"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                </svg>
                Download Presentation (.pptx)
              </button>
              <button
                type="button"
                className="scooh-btn secondary"
                onClick={handleDownloadExcel}
                id="btn-download-excel-conv-xls-top"
                style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.35)' }}
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                Download Standardized Excel (.xlsx)
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={handleReset}
                id="btn-reset-excel-converter-top"
              >
                Convert Another Workbook
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Mandatory Reference Standards Status Card */}
      <div className="scooh-panel scooh-ppt-template-bar" style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '14px' }}>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
            <div className="scooh-ppt-template-icon" style={{ background: 'rgba(34, 197, 94, 0.12)', color: '#4ade80' }}>
              <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
                <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
              </svg>
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                <strong style={{ color: '#fff', fontSize: '14px' }}>
                  Excel Master Reference Standard:
                </strong>
                <span className="scooh-pill" style={{ background: 'rgba(255, 194, 0, 0.15)', color: '#FFC200', border: '1px solid rgba(255, 194, 0, 0.3)', fontSize: '11px' }}>
                  Jade Blue 16-10-2026.xlsx (11 Columns)
                </span>
                <span className="scooh-pill active" style={{ fontSize: '11px' }}>
                  Theme: {templateStatus.filename || 'MediaBuzz_Master_Template.pptx'}
                </span>
              </div>
              <p className="scooh-ppt-template-desc" style={{ marginTop: '4px', marginBottom: 0 }}>
                11 Master Fields in Order: <code>SR NO</code>, <code>City</code>, <code>LOCATION</code>, <code>MEDIA</code>, <code>LIGHT</code>, <code>W</code>, <code>H</code>, <code>SQ FT</code>, <code>AVAILABLITY</code>, <code>Selling Amount</code>, <code>Latitude Longitude</code>.
              </p>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <button
              type="button"
              className="scooh-btn secondary"
              onClick={() => setShowTermsModal(true)}
              style={{ fontSize: '12px', padding: '4px 10px' }}
            >
              🛡️ View Software Terms (7 Pts)
            </button>
          </div>
        </div>

        {/* Governance Notice */}
        <div style={{ background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(34, 197, 94, 0.25)', borderRadius: '8px', padding: '8px 12px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px', fontSize: '12px', color: '#cbd5e1' }}>
          <div>
            <span style={{ color: '#4ade80', fontWeight: 700, marginRight: '6px' }}>EXCEL POLICY:</span>
            <span>External terms inside uploaded workbooks are ignored. Official 7-point software terms are attached. Missing values remain blank.</span>
          </div>
          <span style={{ color: '#4ade80', fontWeight: 600, fontSize: '11px' }}>✓ Policy Active</span>
        </div>
      </div>

      {/* Error Message */}
      {error && (
        <div className="scooh-banner error" style={{ position: 'static', marginBottom: '20px', maxWidth: '100%' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <strong>Conversion Error:</strong> {error}
            </div>
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ padding: '2px 8px', minHeight: '26px', fontSize: '11px', color: '#ffc2c2' }}
              onClick={() => setError(null)}
            >
              Dismiss
            </button>
          </div>
        </div>
      )}

      {/* Main Workflow View: Upload State vs Converted Result */}
      {!conversionResult ? (
        <div className="scooh-panel scooh-ppt-upload-panel">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '16px', flexWrap: 'wrap', gap: '10px' }}>
            <div>
              <h3 style={{ margin: '0 0 4px', color: '#fff' }}>Upload Excel Workbook (.xlsx / .xls)</h3>
              <p className="scooh-footnote" style={{ margin: 0 }}>
                Upload any vendor or client Excel sheet. Structure will be inspected against the 11-column master format.
              </p>
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <span className="scooh-tag" style={{ background: 'rgba(34, 197, 94, 0.1)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.25)' }}>
                .xlsx Supported
              </span>
              <span className="scooh-tag" style={{ background: 'rgba(56, 189, 248, 0.1)', color: '#38bdf8', border: '1px solid rgba(56, 189, 248, 0.25)' }}>
                .xls Supported
              </span>
            </div>
          </div>

          {/* Configuration Inputs: Manual City & Rename File */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px', marginBottom: '20px' }}>
            {/* Manual City Input */}
            <div style={{ background: '#0a1424', border: '1px solid #1e293b', borderRadius: '10px', padding: '14px' }}>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, color: '#f1f5f9', marginBottom: '6px' }}>
                🏙️ Enter City Manually (Default / Override)
              </label>
              <input
                type="text"
                value={manualCity}
                onChange={(e) => setManualCity(e.target.value)}
                placeholder="e.g. Ahmedabad, Surat, Vadodara..."
                style={{
                  width: '100%',
                  boxSizing: 'border-box',
                  background: '#111e33',
                  border: '1px solid #2d3e58',
                  borderRadius: '6px',
                  color: '#fff',
                  fontSize: '13px',
                  padding: '8px 12px',
                  outline: 'none',
                  marginBottom: '8px'
                }}
              />
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', alignItems: 'center' }}>
                <span style={{ fontSize: '11px', color: '#64748b' }}>Quick Pick:</span>
                {POPULAR_CITIES.slice(0, 5).map(c => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => setManualCity(c)}
                    style={{
                      background: manualCity === c ? 'rgba(34, 197, 94, 0.25)' : 'rgba(255, 255, 255, 0.05)',
                      color: manualCity === c ? '#4ade80' : '#94a3b8',
                      border: manualCity === c ? '1px solid #4ade80' : '1px solid #334155',
                      borderRadius: '4px',
                      padding: '2px 7px',
                      fontSize: '10.5px',
                      cursor: 'pointer'
                    }}
                  >
                    {c}
                  </button>
                ))}
                {manualCity && (
                  <button
                    type="button"
                    onClick={() => setManualCity('')}
                    style={{ background: 'transparent', border: 'none', color: '#ef4444', fontSize: '10.5px', cursor: 'pointer', padding: '0 4px' }}
                  >
                    Clear
                  </button>
                )}
              </div>
            </div>

            {/* Rename File Name Input */}
            <div style={{ background: '#0a1424', border: '1px solid #1e293b', borderRadius: '10px', padding: '14px' }}>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, color: '#f1f5f9', marginBottom: '6px' }}>
                ✏️ Output Name (Presentation & Excel)
              </label>
              <div style={{ display: 'flex', alignItems: 'center', background: '#111e33', border: '1px solid #2d3e58', borderRadius: '6px', padding: '2px 10px' }}>
                <input
                  type="text"
                  value={customFileName}
                  onChange={(e) => setCustomFileName(e.target.value)}
                  placeholder="e.g. Jade_Blue_Proposal_MediaBuzz"
                  style={{
                    width: '100%',
                    background: 'transparent',
                    border: 'none',
                    color: '#fff',
                    fontSize: '13px',
                    padding: '6px 0',
                    outline: 'none'
                  }}
                />
                <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>.pptx / .xlsx</span>
              </div>
              <p style={{ margin: '6px 0 0', fontSize: '11px', color: '#64748b' }}>
                You can change the output file title here before converting or on the download screen.
              </p>
            </div>
          </div>

          {/* Drag & Drop Upload Zone */}
          <div
            className={`scooh-ppt-dropzone ${dragActive ? 'drag-over' : ''} ${file ? 'has-file' : ''}`}
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
            onClick={() => !file && fileInputRef.current?.click()}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              accept=".xlsx,.xls"
              style={{ display: 'none' }}
            />

            {!file ? (
              <div className="scooh-ppt-drop-prompt">
                <div className="scooh-ppt-drop-icon" style={{ background: 'rgba(34, 197, 94, 0.08)', color: '#4ade80', borderColor: 'rgba(34, 197, 94, 0.2)' }}>
                  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                  </svg>
                </div>
                <h4>Drag and drop Excel workbook here</h4>
                <p>or click to browse (.xlsx or .xls formats)</p>
                <div className="scooh-ppt-drop-badge">
                  <span>Excel structure verified against master Jade Blue 16-10-2026.xlsx</span>
                </div>
              </div>
            ) : (
              <div className="scooh-ppt-file-card">
                <div className="scooh-ppt-file-icon" style={{ background: 'rgba(34, 197, 94, 0.12)' }}>
                  <svg width="36" height="36" viewBox="0 0 24 24" fill="#22c55e">
                    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                  </svg>
                </div>
                <div className="scooh-ppt-file-details">
                  <strong className="scooh-ppt-file-name">{file.name}</strong>
                  <span className="scooh-ppt-file-size">{formatBytes(file.size)}</span>
                  {inspecting && (
                    <span className="scooh-pill" style={{ marginLeft: '12px', background: 'rgba(56, 189, 248, 0.2)', color: '#38bdf8' }}>
                      Inspecting structure...
                    </span>
                  )}
                  {!inspecting && excelInspection?.matchesMaster && (
                    <span className="scooh-pill active" style={{ marginLeft: '12px' }}>
                      ✓ Matches Jade Blue Reference Standard
                    </span>
                  )}
                  {!inspecting && excelInspection && !excelInspection.matchesMaster && (
                    <span className="scooh-pill" style={{ marginLeft: '12px', background: 'rgba(245, 158, 11, 0.2)', color: '#fbbf24', border: '1px solid rgba(245, 158, 11, 0.4)' }}>
                      ⚠️ Custom Structure (Mapping Available)
                    </span>
                  )}
                </div>
                <button
                  type="button"
                  className="scooh-iconbtn danger-icon"
                  onClick={(e) => {
                    e.stopPropagation();
                    handleReset();
                  }}
                  title="Remove file"
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
                  </svg>
                </button>
              </div>
            )}
          </div>

          {/* Excel Inspection Overview & Column Mapping Access */}
          {file && excelInspection && (
            <div style={{ marginTop: '16px', background: '#0a1628', border: '1px solid #1e293b', borderRadius: '10px', padding: '16px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', flexWrap: 'wrap', gap: '8px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '13px', fontWeight: 800, color: excelInspection.matchesMaster ? '#4ade80' : '#fbbf24' }}>
                    {excelInspection.matchesMaster ? '✓ Standard Master Structure Detected' : '⚠️ Non-Standard Structure Detected'}
                  </span>
                  <span style={{ fontSize: '12px', color: '#94a3b8' }}>
                    ({excelInspection.totalSites} sites detected in sheet "{excelInspection.activeSheet}")
                  </span>
                </div>
                <button
                  type="button"
                  className="scooh-btn secondary"
                  onClick={() => setShowMappingModal(true)}
                  style={{ fontSize: '12px', padding: '4px 12px' }}
                >
                  ⚙️ {excelInspection.matchesMaster ? 'Review Column Mapping' : 'Open Column Mapping Interface'}
                </button>
              </div>

              {/* Quick Mapped Headers Pill Bar */}
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', fontSize: '11px', color: '#cbd5e1' }}>
                {MASTER_TARGET_FIELDS.map(f => {
                  const colIdx = columnMapping[f.key];
                  const colName = colIdx && excelInspection.headers[colIdx - 1] ? excelInspection.headers[colIdx - 1] : null;
                  return (
                    <span
                      key={f.key}
                      style={{
                        padding: '3px 8px',
                        borderRadius: '5px',
                        background: colName ? 'rgba(34, 197, 94, 0.12)' : 'rgba(255, 255, 255, 0.05)',
                        border: colName ? '1px solid rgba(34, 197, 94, 0.3)' : '1px solid #334155',
                        color: colName ? '#e2e8f0' : '#64748b'
                      }}
                      title={colName ? `Mapped to Column ${colIdx}: "${colName}"` : 'Not mapped (will be left blank)'}
                    >
                      <strong>{f.masterCol}:</strong> {colName || 'Blank'}
                    </span>
                  );
                })}
              </div>

              <div style={{ marginTop: '12px', padding: '8px 12px', background: 'rgba(15, 23, 42, 0.6)', borderLeft: '3px solid #ffc200', borderRadius: '4px', fontSize: '11.5px', color: '#94a3b8' }}>
                <strong style={{ color: '#ffc200' }}>Note on Terms & Conditions:</strong> Any Terms & Conditions found in the uploaded workbook (e.g. row 18+) are automatically ignored. The official 7-point Media Buzz software terms will be attached to the final presentation.
              </div>
            </div>
          )}

          {/* Conversion Progress Indicator */}
          {converting && (
            <div className="scooh-ppt-progress-container" style={{ marginTop: '20px' }}>
              <div className="scooh-ppt-progress-header">
                <span className="scooh-ppt-progress-msg">{progressMessage}</span>
                <span className="scooh-ppt-progress-pct">{progressPercent}%</span>
              </div>
              <div className="scooh-ppt-progress-bar">
                <div
                  className="scooh-ppt-progress-fill"
                  style={{ width: `${progressPercent}%`, background: 'linear-gradient(90deg, #22c55e, #4ade80)' }}
                />
              </div>
              <div className="scooh-ppt-steps">
                <div className={`scooh-ppt-step ${progressStep >= 1 ? 'active' : ''} ${progressStep > 1 ? 'done' : ''}`}>
                  <span>1</span> Upload Workbook
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 2 ? 'active' : ''} ${progressStep > 2 ? 'done' : ''}`}>
                  <span>2</span> Parse Listings
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 3 ? 'active' : ''} ${progressStep > 3 ? 'done' : ''}`}>
                  <span>3</span> Apply Master Format
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 4 ? 'active' : ''} ${progressStep >= 4 ? 'done' : ''}`}>
                  <span>4</span> Ready to Export
                </div>
              </div>
            </div>
          )}

          {/* Convert Trigger Button */}
          <div className="scooh-ppt-upload-actions" style={{ marginTop: '20px' }}>
            <button
              type="button"
              className="scooh-btn primary large"
              onClick={handleConvert}
              disabled={!file || !templateStatus.available || converting || inspecting}
              id="btn-convert-excel"
              style={{ minWidth: '260px', padding: '12px 24px', fontSize: '14px', fontWeight: 700 }}
            >
              {converting ? (
                <>
                  <span className="scooh-spin">🔄</span> Converting Excel...
                </>
              ) : (
                <>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                  </svg>
                  Convert Excel to Media Buzz PPT & Master Excel
                </>
              )}
            </button>
            {!templateStatus.available && (
              <span className="scooh-footnote" style={{ color: 'var(--mb-warning)' }}>
                ⚠ Cannot convert: Master template is unavailable.
              </span>
            )}
          </div>
        </div>
      ) : (
        /* Results and Slide Previews Area */
        <div className="scooh-ppt-results-view">
          {/* Conversion Success Overview Card */}
          <div className="scooh-panel scooh-ppt-success-card">
            <div className="scooh-ppt-success-left">
              <div className="scooh-ppt-success-badge" style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80' }}>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                </svg>
              </div>
              <div>
                <h3 style={{ margin: '0 0 4px', color: '#fff' }}>
                  {conversionResult.processedSlides} Billboard Sites Converted from Excel
                </h3>
                <p style={{ margin: 0, color: 'var(--mb-muted)', fontSize: '12px' }}>
                  Official 16:9 widescreen layout with verified software Terms & Conditions.
                  {conversionResult.manualCity && (
                    <span style={{ color: '#4ade80', marginLeft: '6px' }}>
                      City: <strong>{conversionResult.manualCity}</strong>.
                    </span>
                  )}
                  Total presentation size: <strong>{conversionResult.totalSlidesInFile} slides</strong>.
                </p>
              </div>
            </div>

            <div className="scooh-ppt-success-actions" style={{ display: 'flex', flexDirection: 'column', gap: '8px', minWidth: '340px' }}>
              <div style={{ display: 'flex', alignItems: 'center', background: 'rgba(255,255,255,0.06)', borderRadius: '6px', border: '1px solid var(--mb-border)', padding: '2px 10px' }}>
                <span style={{ fontSize: '11px', color: 'var(--mb-primary)', fontWeight: 700, marginRight: '8px', letterSpacing: '0.5px' }}>RENAME:</span>
                <input
                  type="text"
                  value={customFileName}
                  onChange={(e) => setCustomFileName(e.target.value)}
                  placeholder="Output File Name"
                  title="Click to rename files before downloading"
                  style={{
                    background: 'transparent',
                    border: 'none',
                    color: '#ffffff',
                    fontSize: '12px',
                    fontWeight: 600,
                    outline: 'none',
                    flex: 1,
                    padding: '6px 0'
                  }}
                />
                <span style={{ fontSize: '11px', color: 'var(--mb-muted)', fontWeight: 600 }}>.pptx / .xlsx</span>
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button
                  type="button"
                  className="scooh-btn primary"
                  onClick={handleDownloadPpt}
                  id="btn-download-excel-conv-ppt-main"
                  style={{ flex: 1 }}
                >
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                  </svg>
                  Download PPT
                </button>
                <button
                  type="button"
                  className="scooh-btn secondary"
                  onClick={handleDownloadExcel}
                  id="btn-download-excel-conv-xls-main"
                  style={{ flex: 1, background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.35)' }}
                >
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                  </svg>
                  Download Excel
                </button>
              </div>
            </div>
          </div>

          {/* Slide Deck Grid / Cards */}
          <div className="scooh-panel scooh-ppt-gallery-panel">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '10px' }}>
              <div>
                <h3 style={{ margin: '0 0 2px', color: '#fff' }}>Slide Deck & Site Cards ({slides.length})</h3>
                <span className="scooh-footnote">Includes Title Covers, Billboard Sites, Terms & Conditions, and Closing Slide</span>
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button
                  type="button"
                  className={`scooh-btn ${filterReviewOnly ? 'primary' : 'ghost'}`}
                  onClick={() => setFilterReviewOnly(!filterReviewOnly)}
                  style={{ fontSize: '12px', padding: '4px 10px' }}
                >
                  {filterReviewOnly ? 'Showing Attention Needed' : 'Filter Needs Attention'}
                </button>
              </div>
            </div>

            <div className="scooh-ppt-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '16px' }}>
              {filteredSlides.map((slide, idx) => (
                <div
                  key={idx}
                  className="scooh-panel"
                  style={{
                    background: '#0d1522',
                    border: slide.reviewNeeded ? '1px solid rgba(245, 158, 11, 0.5)' : '1px solid #1e293b',
                    borderRadius: '10px',
                    padding: '12px',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '10px'
                  }}
                >
                  {/* Thumbnail / Visual Header */}
                  <div style={{ position: 'relative', borderRadius: '6px', overflow: 'hidden', background: '#050b14', aspectRatio: '16 / 9', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                    {slide.previewUrl ? (
                      <img
                        src={slide.previewUrl}
                        alt={`Site ${slide.siteNumber}`}
                        style={{ width: '100%', height: '100%', objectFit: 'cover' }}
                      />
                    ) : (
                      <div style={{ color: '#64748b', fontSize: '12px', textAlign: 'center', padding: '10px' }}>
                        <span>Site {slide.siteNumber}</span>
                        <div style={{ fontSize: '11px', marginTop: '4px' }}>No photo attached in sheet</div>
                      </div>
                    )}
                    <span
                      style={{
                        position: 'absolute',
                        top: '8px',
                        left: '8px',
                        background: '#FFC200',
                        color: '#000',
                        fontWeight: 800,
                        fontSize: '11px',
                        padding: '2px 8px',
                        borderRadius: '4px'
                      }}
                    >
                      Site {slide.siteNumber}
                    </span>
                  </div>

                  {/* Site Details */}
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '12px' }}>
                    <div style={{ fontWeight: 700, color: '#f1f5f9', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      {slide.location || 'Location Not Specified'}
                    </div>
                    <div style={{ color: '#94a3b8', fontSize: '11px' }}>
                      <span>City: <strong style={{ color: '#cbd5e1' }}>{slide.city || manualCity || 'Ahmedabad'}</strong></span>
                      <span style={{ margin: '0 6px' }}>•</span>
                      <span>Size: <strong style={{ color: '#cbd5e1' }}>{slide.size || '30x10 ft'}</strong></span>
                      <span style={{ margin: '0 6px' }}>•</span>
                      <span>Type: <strong style={{ color: '#cbd5e1' }}>{slide.type || 'Hoarding'}</strong></span>
                    </div>
                    <div style={{ color: '#94a3b8', fontSize: '11px' }}>
                      <span>Availability: <strong style={{ color: '#4ade80' }}>{slide.availability || 'Available'}</strong></span>
                      {slide.rate && (
                        <>
                          <span style={{ margin: '0 6px' }}>•</span>
                          <span>Rate: <strong style={{ color: '#FFC200' }}>{slide.rate}</strong></span>
                        </>
                      )}
                    </div>
                    {slide.coords && (
                      <div style={{ color: '#64748b', fontSize: '10.5px' }}>
                        📍 {slide.coords}
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* COLUMN MAPPING MODAL (Rule 2 & 4: Mandatory Master Excel Column Alignment) */}
      {/* ========================================================================= */}
      {showMappingModal && excelInspection && (
        <div
          className="scooh-modal-backdrop"
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 9999,
            background: 'rgba(0, 0, 0, 0.85)',
            backdropFilter: 'blur(6px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px'
          }}
          onClick={() => setShowMappingModal(false)}
        >
          <div
            className="scooh-panel"
            style={{
              width: '100%',
              maxWidth: '960px',
              maxHeight: '90vh',
              overflowY: 'auto',
              border: '1px solid rgba(34, 197, 94, 0.4)',
              boxShadow: '0 25px 60px rgba(0,0,0,0.85), 0 0 25px rgba(34, 197, 94, 0.15)',
              background: '#071224',
              padding: '24px',
              borderRadius: '14px',
              color: '#f8fafc'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #1e293b', paddingBottom: '14px', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '18px', fontWeight: 800 }}>
                  <span style={{ color: '#4ade80' }}>⚙️</span>
                  <span>Excel Column Mapping Interface</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  Target standard: <strong>Jade Blue 16-10-2026.xlsx</strong>. Map each required field to the corresponding column header in your uploaded workbook.
                </p>
              </div>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => setShowMappingModal(false)}
                style={{ padding: '4px 10px', fontSize: '14px', minHeight: '30px' }}
              >
                ✕
              </button>
            </div>

            {/* Sheet Selector */}
            {excelInspection.sheets && excelInspection.sheets.length > 1 && (
              <div style={{ marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '12px', background: '#0a1628', padding: '10px 14px', borderRadius: '8px' }}>
                <span style={{ fontSize: '12.5px', fontWeight: 700, color: '#e2e8f0' }}>Target Worksheet:</span>
                <select
                  value={selectedSheet}
                  onChange={(e) => setSelectedSheet(e.target.value)}
                  style={{ background: '#111e33', color: '#fff', border: '1px solid #2d3e58', borderRadius: '6px', padding: '6px 12px', fontSize: '12.5px' }}
                >
                  {excelInspection.sheets.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              </div>
            )}

            {/* Mapping Grid */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(420px, 1fr))', gap: '12px', marginBottom: '20px' }}>
              {MASTER_TARGET_FIELDS.map((field) => {
                const currentIdx = columnMapping[field.key] || '';
                return (
                  <div
                    key={field.key}
                    style={{
                      background: '#0a1628',
                      border: currentIdx ? '1px solid rgba(34, 197, 94, 0.3)' : '1px solid #1e293b',
                      borderRadius: '8px',
                      padding: '12px 14px'
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: '6px' }}>
                      <strong style={{ color: '#fff', fontSize: '13px' }}>
                        {field.masterCol}
                        {field.required && <span style={{ color: '#ef4444', marginLeft: '4px' }}>*</span>}
                      </strong>
                      <span style={{ fontSize: '11px', color: '#64748b' }}>{field.desc}</span>
                    </div>

                    <select
                      value={currentIdx}
                      onChange={(e) => {
                        const val = e.target.value ? parseInt(e.target.value, 10) : '';
                        setColumnMapping(prev => ({
                          ...prev,
                          [field.key]: val
                        }));
                      }}
                      style={{
                        width: '100%',
                        background: '#111e33',
                        border: '1px solid #2d3e58',
                        color: currentIdx ? '#4ade80' : '#94a3b8',
                        borderRadius: '6px',
                        padding: '8px 10px',
                        fontSize: '12.5px',
                        fontWeight: currentIdx ? 700 : 400
                      }}
                    >
                      <option value="">-- Leave Blank (No match in sheet) --</option>
                      {excelInspection.headers.map((h, i) => (
                        <option key={i} value={i + 1}>
                          Col {i + 1}: "{h}"
                        </option>
                      ))}
                    </select>
                  </div>
                );
              })}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px', borderTop: '1px solid #1e293b', paddingTop: '14px' }}>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => setShowMappingModal(false)}
              >
                Cancel
              </button>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={() => setShowMappingModal(false)}
                style={{ minHeight: '38px', padding: '0 20px', fontSize: '12.5px', fontWeight: 800 }}
              >
                ✓ Confirm & Apply Mapping
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Software Terms Modal */}
      {showTermsModal && (
        <div
          className="scooh-modal-backdrop"
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 9999,
            background: 'rgba(0, 0, 0, 0.85)',
            backdropFilter: 'blur(6px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px'
          }}
          onClick={() => setShowTermsModal(false)}
        >
          <div
            className="scooh-panel"
            style={{
              width: '100%',
              maxWidth: '680px',
              border: '1px solid #FFC200',
              boxShadow: '0 25px 60px rgba(0,0,0,0.85), 0 0 25px rgba(255, 194, 0, 0.2)',
              background: '#071224',
              padding: '24px',
              borderRadius: '14px',
              color: '#f8fafc'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #1e293b', paddingBottom: '12px', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '18px', fontWeight: 800 }}>
                  <span style={{ color: '#FFC200' }}>🛡️</span>
                  <span>Official Media Buzz Software Terms & Conditions</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  Mandatory governance: Retrieved exclusively from Media Buzz software. Any terms in master workbook or uploaded workbook are ignored.
                </p>
              </div>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => setShowTermsModal(false)}
                style={{ padding: '4px 10px', fontSize: '14px', minHeight: '30px' }}
              >
                ✕
              </button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginBottom: '20px' }}>
              {OFFICIAL_SOFTWARE_TERMS.map((term, tIdx) => {
                const dotPos = term.indexOf('.');
                const num = term.slice(0, dotPos + 1);
                const txt = term.slice(dotPos + 1).trim();
                return (
                  <div
                    key={tIdx}
                    style={{
                      display: 'flex',
                      gap: '12px',
                      alignItems: 'flex-start',
                      padding: '10px 14px',
                      background: '#0a1628',
                      border: '1px solid #1e293b',
                      borderRadius: '8px'
                    }}
                  >
                    <span style={{ color: '#FFC200', fontWeight: 800, fontSize: '14px', minWidth: '24px' }}>
                      {num}
                    </span>
                    <span style={{ color: '#e2e8f0', fontSize: '13px', lineHeight: 1.5 }}>
                      {txt}
                    </span>
                  </div>
                );
              })}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', borderTop: '1px solid #1e293b', paddingTop: '14px' }}>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={() => setShowTermsModal(false)}
                style={{ minHeight: '36px', padding: '0 20px', fontSize: '12.5px' }}
              >
                Close Terms
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
