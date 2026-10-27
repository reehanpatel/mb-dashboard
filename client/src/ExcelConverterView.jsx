import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import api from './api';
import { MEDIA_BUZZ_LOGO_BASE64 } from './excelLogo';

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
  { key: 'avail', label: 'Availability', masterCol: 'AVAILABLITY', required: false, desc: 'Available' },
  { key: 'rate', label: 'Selling Amount', masterCol: 'Selling Amount', required: false, desc: 'Monthly advertising fee (INR)' },
  { key: 'coords', label: 'Location Coordinates', masterCol: 'Latitude Longitude', required: false, desc: 'GPS Lat, Lng coordinates' }
];

const POPULAR_CITIES = ['Ahmedabad', 'Surat', 'Vadodara', 'Rajkot', 'Gandhinagar', 'Mumbai', 'Pune', 'Delhi NCR'];

export default function ExcelConverterView() {
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

  // Table preview search filter
  const [tableFilter, setTableFilter] = useState('');

  // Pre-conversion Excel Sheet Preview state
  const [prePreviewTab, setPrePreviewTab] = useState('mapped'); // 'mapped' | 'raw'
  const [preFilter, setPreFilter] = useState('');
  const [showPrePreview, setShowPrePreview] = useState(true);

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
  const inspectExcelWorkbook = async (selectedFile, targetSheet = null, customMap = null) => {
    try {
      setInspecting(true);
      setError(null);
      const fd = new FormData();
      fd.append('file', selectedFile);
      const sheetToUse = targetSheet !== null ? targetSheet : selectedSheet;
      if (sheetToUse) {
        fd.append('sheetName', sheetToUse);
      }
      const mapToUse = customMap !== null ? customMap : columnMapping;
      if (mapToUse && Object.keys(mapToUse).length > 0) {
        fd.append('columnMapping', JSON.stringify(mapToUse));
      }
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
    } catch (err) {
      console.error('Excel inspection error:', err);
      setError('Failed to inspect Excel workbook: ' + (err.response?.data?.message || err.message));
    } finally {
      setInspecting(false);
    }
  };

  const handleSheetChange = (newSheet) => {
    setSelectedSheet(newSheet);
    if (file) {
      inspectExcelWorkbook(file, newSheet, {});
    }
  };

  const handleApplyCustomMapping = () => {
    setShowMappingModal(false);
    if (file) {
      inspectExcelWorkbook(file, selectedSheet, columnMapping);
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
    if (!file || converting) return;

    setError(null);
    setConverting(true);
    setProgressStep(1);
    setProgressPercent(20);
    setProgressMessage('Uploading Excel workbook for conversion...');

    const pInterval = setInterval(() => {
      setProgressPercent(prev => {
        if (prev < 45) return prev + 10;
        if (prev < 80) return prev + 5;
        if (prev < 95) return prev + 2;
        return prev;
      });
    }, 300);

    setTimeout(() => {
      setProgressStep(2);
      setProgressMessage('Parsing billboard rows, specs & mapped columns...');
    }, 800);

    setTimeout(() => {
      setProgressStep(3);
      setProgressMessage('Generating clean 11-column Master Excel with Media Buzz branding...');
    }, 1800);

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
      setProgressMessage('Standardized Master Excel ready to export!');

      setTimeout(() => {
        setConversionResult(res.data);
        if (res.data?.customName) {
          setCustomFileName(res.data.customName);
        }
        setConverting(false);
      }, 400);

    } catch (err) {
      clearInterval(pInterval);
      setConverting(false);
      setProgressStep(0);
      setProgressPercent(0);
      const msg = err.response?.data?.message || err.message || 'An error occurred during Excel conversion.';
      setError(msg);
    }
  };

  // Download standardized Excel (.xlsx)
  const handleDownloadExcel = async () => {
    const downloadId = conversionResult?.excelDownloadId || conversionResult?.downloadId;
    if (!downloadId) return;

    try {
      const cleanBase = (customFileName || 'Converted_MediaBuzz_Master_Format')
        .trim()
        .replace(/\.xlsx$/i, '')
        .replace(/\.pptx$/i, '');
      const cleanName = `${cleanBase}.xlsx`;
      const downloadUrl = `/excel-converter/download/${encodeURIComponent(downloadId)}?name=${encodeURIComponent(cleanName)}`;
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
    setTableFilter('');
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const formatBytes = (bytes) => {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // Extracted sites normalized for Excel preview table
  const sites = useMemo(() => {
    if (!conversionResult) return [];
    if (Array.isArray(conversionResult.sites) && conversionResult.sites.length > 0) {
      return conversionResult.sites;
    }
    // Fallback if legacy slides format is provided
    if (Array.isArray(conversionResult.slides)) {
      return conversionResult.slides
        .filter(s => !s.isTermsSlide)
        .map((s, idx) => {
          const sizeTxt = String(s.size || '').replace(/ft|FT/g, '').trim();
          const parts = sizeTxt.split(/[xX*×]/);
          const w = parts[0] ? parts[0].trim() : '';
          const h = parts[1] ? parts[1].trim() : '';
          const sqFt = (w && h && !isNaN(w) && !isNaN(h)) ? String(Math.round(parseFloat(w) * parseFloat(h))) : '';
          return {
            srNo: idx + 1,
            city: s.city || manualCity || 'Ahmedabad',
            location: s.location || '',
            media: s.type || 'Hoarding',
            light: s.illum || 'BL',
            w: w,
            h: h,
            sqFt: sqFt,
            availability: s.availability || 'Available',
            rate: s.rate || '',
            coords: s.coords || ''
          };
        });
    }
    return [];
  }, [conversionResult, manualCity]);

  const filteredSites = useMemo(() => {
    if (!tableFilter.trim()) return sites;
    const q = tableFilter.toLowerCase();
    return sites.filter(s =>
      String(s.location || '').toLowerCase().includes(q) ||
      String(s.city || '').toLowerCase().includes(q) ||
      String(s.media || '').toLowerCase().includes(q) ||
      String(s.light || '').toLowerCase().includes(q) ||
      String(s.rate || '').toLowerCase().includes(q) ||
      String(s.coords || '').toLowerCase().includes(q) ||
      String(s.srNo || '').includes(q)
    );
  }, [sites, tableFilter]);

  // Aggregate stats
  const totalSqFt = useMemo(() => {
    if (conversionResult?.totalSqFt) return conversionResult.totalSqFt.toLocaleString();
    const sum = sites.reduce((acc, s) => acc + (parseInt(String(s.sqFt).replace(/,/g, ''), 10) || 0), 0);
    return sum.toLocaleString();
  }, [conversionResult, sites]);

  const distinctCities = useMemo(() => {
    if (conversionResult?.cities && conversionResult.cities.length > 0) {
      return conversionResult.cities.join(', ');
    }
    const set = new Set(sites.map(s => s.city).filter(Boolean));
    return Array.from(set).join(', ') || (manualCity || 'All Markets');
  }, [conversionResult, sites, manualCity]);

  // Pre-conversion detected sites normalized for preview
  const preDetectedSites = useMemo(() => {
    if (!excelInspection) return [];
    if (Array.isArray(excelInspection.detectedRows) && excelInspection.detectedRows.length > 0) {
      return excelInspection.detectedRows.map(r => ({
        ...r,
        city: manualCity || r.city || 'Ahmedabad'
      }));
    }
    if (Array.isArray(excelInspection.sampleRows)) {
      return excelInspection.sampleRows.map((r, idx) => ({
        srNo: idx + 1,
        city: manualCity || r.city || 'Ahmedabad',
        location: r.location || '',
        media: r.media || 'Hoarding',
        light: r.light || 'BL',
        w: r.w || '',
        h: r.h || '',
        sqFt: r.sqFt || '',
        availability: r.avail || 'Available',
        rate: r.rate || '',
        coords: r.coords || ''
      }));
    }
    return [];
  }, [excelInspection, manualCity]);

  const filteredPreSites = useMemo(() => {
    if (!preFilter.trim()) return preDetectedSites;
    const q = preFilter.toLowerCase();
    return preDetectedSites.filter(s =>
      String(s.location || '').toLowerCase().includes(q) ||
      String(s.city || '').toLowerCase().includes(q) ||
      String(s.media || '').toLowerCase().includes(q) ||
      String(s.light || '').toLowerCase().includes(q) ||
      String(s.rate || '').toLowerCase().includes(q) ||
      String(s.coords || '').toLowerCase().includes(q) ||
      String(s.srNo || '').includes(q)
    );
  }, [preDetectedSites, preFilter]);

  const filteredRawRows = useMemo(() => {
    if (!excelInspection || !Array.isArray(excelInspection.rawRows)) return [];
    if (!preFilter.trim()) return excelInspection.rawRows;
    const q = preFilter.toLowerCase();
    return excelInspection.rawRows.filter(r =>
      r.cells && r.cells.some(c => String(c || '').toLowerCase().includes(q))
    );
  }, [excelInspection, preFilter]);

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
            Convert any vendor or client Excel workbook (<code>.xlsx</code> / <code>.xls</code>) directly into clean 11-column standardized Master Excel format.
            Includes Media Buzz brand header banner, official logo, and standard 7-point software Terms &amp; Conditions.
          </p>
        </div>
        <div className="scooh-headactions">
          {conversionResult && (
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleDownloadExcel}
                id="btn-download-master-excel-top"
                style={{ background: 'linear-gradient(135deg, #16a34a, #22c55e)', color: '#ffffff', fontWeight: 700 }}
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                Download Master Excel (.xlsx)
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
                  Standardized .xlsx Export
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
            <span>External terms inside uploaded workbooks are ignored. Official 7-point software terms are attached at bottom. Missing values remain blank.</span>
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
            {/* Manual City Override */}
            <div className="scooh-formgroup" style={{ margin: 0 }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', fontWeight: 600, color: '#e2e8f0', marginBottom: '6px' }}>
                <span>🏙️</span>
                <span>Enter City Manually (Default / Override)</span>
              </label>
              <input
                type="text"
                value={manualCity}
                onChange={(e) => setManualCity(e.target.value)}
                placeholder="e.g. Ahmedabad, Surat, Vadodara..."
                style={{
                  width: '100%',
                  background: '#0d1522',
                  border: '1px solid #1e293b',
                  borderRadius: '6px',
                  padding: '9px 12px',
                  color: '#fff',
                  fontSize: '13px'
                }}
              />
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '6px' }}>
                <span style={{ fontSize: '10.5px', color: '#64748b', alignSelf: 'center' }}>Quick Pick:</span>
                {POPULAR_CITIES.slice(0, 5).map(c => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => setManualCity(c)}
                    style={{
                      background: manualCity === c ? 'rgba(34, 197, 94, 0.2)' : 'rgba(255, 255, 255, 0.04)',
                      border: manualCity === c ? '1px solid rgba(34, 197, 94, 0.4)' : '1px solid #1e293b',
                      color: manualCity === c ? '#4ade80' : '#94a3b8',
                      borderRadius: '4px',
                      padding: '2px 6px',
                      fontSize: '11px',
                      cursor: 'pointer'
                    }}
                  >
                    {c}
                  </button>
                ))}
              </div>
            </div>

            {/* Rename Output File */}
            <div className="scooh-formgroup" style={{ margin: 0 }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', fontWeight: 600, color: '#e2e8f0', marginBottom: '6px' }}>
                <span>✏️</span>
                <span>Output Name (Standardized Excel)</span>
              </label>
              <div style={{ position: 'relative' }}>
                <input
                  type="text"
                  value={customFileName}
                  onChange={(e) => setCustomFileName(e.target.value)}
                  placeholder="Converted_MediaBuzz_Master_Format"
                  style={{
                    width: '100%',
                    background: '#0d1522',
                    border: '1px solid #1e293b',
                    borderRadius: '6px',
                    padding: '9px 60px 9px 12px',
                    color: '#fff',
                    fontSize: '13px'
                  }}
                />
                <span style={{ position: 'absolute', right: '12px', top: '50%', transform: 'translateY(-50%)', fontSize: '12px', color: '#64748b', pointerEvents: 'none' }}>
                  .xlsx
                </span>
              </div>
              <span className="scooh-footnote" style={{ marginTop: '4px' }}>
                Custom output title applied to the generated Master Excel workbook.
              </span>
            </div>
          </div>

          {/* Drag & Drop Area */}
          <div
            className={`scooh-ppt-dropzone ${dragActive ? 'active' : ''} ${file ? 'has-file' : ''}`}
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: dragActive ? '2px dashed #22c55e' : '2px dashed #1e293b',
              borderRadius: '12px',
              padding: file ? '20px' : '40px 20px',
              background: dragActive ? 'rgba(34, 197, 94, 0.05)' : '#071224',
              cursor: 'pointer',
              transition: 'all 0.2s ease',
              textAlign: 'center'
            }}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileChange}
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
                  {!inspecting && (
                    <span className="scooh-pill active" style={{ marginLeft: '12px' }}>
                      ✓ Auto-Mapped &amp; Ready to Convert
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
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  <span style={{ fontSize: '13px', fontWeight: 800, color: '#4ade80' }}>
                    ✓ Excel Data Auto-Detected
                  </span>
                  <span style={{ fontSize: '12px', color: '#94a3b8' }}>
                    ({excelInspection.totalSites} sites detected in sheet "{excelInspection.activeSheet}")
                  </span>
                  {excelInspection.sheets && excelInspection.sheets.length > 1 && (
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', background: 'rgba(30, 41, 59, 0.8)', padding: '3px 8px', borderRadius: '6px', border: '1px solid #334155' }}>
                      <span style={{ fontSize: '11.5px', color: '#94a3b8', fontWeight: 700 }}>Sheet:</span>
                      <select
                        value={selectedSheet || excelInspection.activeSheet || ''}
                        onChange={(e) => handleSheetChange(e.target.value)}
                        style={{ background: '#0f172a', color: '#38bdf8', border: '1px solid #0284c7', borderRadius: '4px', padding: '2px 6px', fontSize: '11.5px', fontWeight: 700, outline: 'none', cursor: 'pointer' }}
                      >
                        {excelInspection.sheets.map(s => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </div>
                  )}
                </div>
                <button
                  type="button"
                  className="scooh-btn secondary"
                  onClick={() => setShowMappingModal(true)}
                  style={{ fontSize: '12px', padding: '4px 12px' }}
                >
                  ⚙️ Column Mapping (Optional)
                </button>
              </div>

              {/* Quick Mapped Headers Pill Bar */}
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', fontSize: '11px', color: '#cbd5e1' }}>
                {MASTER_TARGET_FIELDS.map(f => {
                  const numIdx = parseInt(columnMapping[f.key], 10);
                  const colName = (!isNaN(numIdx) && numIdx >= 1 && excelInspection.headers[numIdx - 1]) ? excelInspection.headers[numIdx - 1] : null;
                  return (
                    <span
                      key={f.key}
                      onClick={() => setShowMappingModal(true)}
                      style={{
                        padding: '3px 8px',
                        borderRadius: '5px',
                        background: colName ? 'rgba(34, 197, 94, 0.12)' : 'rgba(255, 255, 255, 0.05)',
                        border: colName ? '1px solid rgba(34, 197, 94, 0.3)' : '1px solid #334155',
                        color: colName ? '#e2e8f0' : '#64748b',
                        cursor: 'pointer'
                      }}
                      title={colName ? `Mapped to Column ${numIdx}: "${colName}" (Click to change)` : 'Not mapped (Click to configure)'}
                    >
                      <strong>{f.masterCol}:</strong> {colName || 'Blank'}
                    </span>
                  );
                })}
              </div>

              <div style={{ marginTop: '12px', padding: '8px 12px', background: 'rgba(15, 23, 42, 0.6)', borderLeft: '3px solid #ffc200', borderRadius: '4px', fontSize: '11.5px', color: '#94a3b8' }}>
                <strong style={{ color: '#ffc200' }}>Note on Terms &amp; Conditions:</strong> Any Terms &amp; Conditions found in the uploaded workbook (e.g. row 18+) are automatically ignored. The official 7-point Media Buzz software terms will be attached to the final standardized Excel workbook.
              </div>

              {/* Pre-Conversion Live Excel Sheet Preview */}
              <div style={{ marginTop: '16px', borderTop: '1px solid #1e293b', paddingTop: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', flexWrap: 'wrap', gap: '8px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <div style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', padding: '4px 6px', borderRadius: '4px', display: 'flex', alignItems: 'center' }}>
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                      </svg>
                    </div>
                    <strong style={{ color: '#fff', fontSize: '13px' }}>
                      Detected Sheet Preview ({preDetectedSites.length} Sites)
                    </strong>
                    <button
                      type="button"
                      onClick={() => setShowPrePreview(prev => !prev)}
                      style={{ background: 'none', border: 'none', color: '#38bdf8', fontSize: '11px', cursor: 'pointer', textDecoration: 'underline', padding: 0 }}
                    >
                      {showPrePreview ? 'Collapse Preview' : 'Expand Preview'}
                    </button>
                  </div>

                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {/* View Switcher: Mapped Master vs Raw Uploaded */}
                    <div style={{ display: 'inline-flex', background: '#0a1628', border: '1px solid #1e293b', borderRadius: '6px', padding: '2px' }}>
                      <button
                        type="button"
                        onClick={() => setPrePreviewTab('mapped')}
                        style={{
                          background: prePreviewTab === 'mapped' ? 'rgba(34, 197, 94, 0.2)' : 'transparent',
                          color: prePreviewTab === 'mapped' ? '#4ade80' : '#94a3b8',
                          border: 'none',
                          borderRadius: '4px',
                          padding: '4px 8px',
                          fontSize: '11px',
                          fontWeight: 700,
                          cursor: 'pointer'
                        }}
                      >
                        Standard Master (11 Cols)
                      </button>
                      <button
                        type="button"
                        onClick={() => setPrePreviewTab('raw')}
                        style={{
                          background: prePreviewTab === 'raw' ? 'rgba(56, 189, 248, 0.2)' : 'transparent',
                          color: prePreviewTab === 'raw' ? '#38bdf8' : '#94a3b8',
                          border: 'none',
                          borderRadius: '4px',
                          padding: '4px 8px',
                          fontSize: '11px',
                          fontWeight: 700,
                          cursor: 'pointer'
                        }}
                      >
                        Original Sheet ({excelInspection.headers?.length || 0} Cols)
                      </button>
                    </div>

                    {/* Search */}
                    <input
                      type="text"
                      value={preFilter}
                      onChange={(e) => setPreFilter(e.target.value)}
                      placeholder="Search preview..."
                      style={{
                        background: '#0d1829',
                        border: '1px solid #1e2f47',
                        borderRadius: '6px',
                        padding: '4px 10px',
                        color: '#fff',
                        fontSize: '11.5px',
                        width: '160px'
                      }}
                    />
                  </div>
                </div>

                {showPrePreview && (
                  <div style={{ overflowX: 'auto', borderRadius: '8px', border: '1px solid #233b5c', maxHeight: '380px', overflowY: 'auto' }}>
                    {prePreviewTab === 'mapped' ? (
                      <table style={{ width: '100%', minWidth: '1000px', borderCollapse: 'collapse', fontFamily: 'Calibri, "Segoe UI", sans-serif' }}>
                        <thead style={{ position: 'sticky', top: 0, zIndex: 2 }}>
                          <tr style={{ background: '#FFFF00', color: '#000000', height: '30px', fontSize: '11.5px', fontWeight: 800 }}>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '55px', textAlign: 'center' }}>SR NO</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '100px', textAlign: 'center' }}>City</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 10px', minWidth: '260px', textAlign: 'left' }}>LOCATION</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '90px', textAlign: 'center' }}>MEDIA</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '65px', textAlign: 'center' }}>LIGHT</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '55px', textAlign: 'center' }}>W</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '55px', textAlign: 'center' }}>H</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '75px', textAlign: 'center' }}>SQ FT</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '100px', textAlign: 'center' }}>AVAILABLITY</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 10px', width: '120px', textAlign: 'right' }}>Selling Amount</th>
                            <th style={{ border: '1px solid #808080', padding: '6px 10px', width: '150px', textAlign: 'left' }}>Latitude Longitude</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredPreSites.length === 0 ? (
                            <tr>
                              <td colSpan={11} style={{ textAlign: 'center', padding: '24px', color: '#94a3b8' }}>
                                No sites matching "{preFilter}"
                              </td>
                            </tr>
                          ) : (
                            filteredPreSites.map((site, sIdx) => {
                              const isEven = sIdx % 2 === 0;
                              return (
                                <tr
                                  key={sIdx}
                                  style={{
                                    background: isEven ? '#0a1525' : '#0d1b2e',
                                    color: '#e2e8f0',
                                    fontSize: '11.5px',
                                    transition: 'background 0.15s ease'
                                  }}
                                  onMouseEnter={(e) => { e.currentTarget.style.background = '#132845'; }}
                                  onMouseLeave={(e) => { e.currentTarget.style.background = isEven ? '#0a1525' : '#0d1b2e'; }}
                                >
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#94a3b8' }}>{site.srNo}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#cbd5e1' }}>{site.city}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 10px', textAlign: 'left', fontWeight: 600, color: '#f8fafc' }}>{site.location}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#cbd5e1' }}>{site.media}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#cbd5e1' }}>{site.light}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#e2e8f0' }}>{site.w}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#e2e8f0' }}>{site.h}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#e2e8f0' }}>{site.sqFt}</td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center' }}>
                                    <span style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', padding: '1px 6px', borderRadius: '4px', fontSize: '10.5px', fontWeight: 700 }}>
                                      {site.availability || 'Available'}
                                    </span>
                                  </td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 10px', textAlign: 'right', fontWeight: 700, color: '#FFC200' }}>
                                    {site.rate || '—'}
                                  </td>
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 10px', textAlign: 'left', color: '#94a3b8', fontFamily: 'monospace', fontSize: '10.5px' }}>
                                    {site.coords || '—'}
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    ) : (
                      /* Raw Sheet Preview */
                      <table style={{ width: '100%', minWidth: '900px', borderCollapse: 'collapse', fontFamily: 'Calibri, "Segoe UI", sans-serif' }}>
                        <thead style={{ position: 'sticky', top: 0, zIndex: 2 }}>
                          <tr style={{ background: '#1e293b', color: '#38bdf8', height: '30px', fontSize: '11px', fontWeight: 700 }}>
                            <th style={{ border: '1px solid #334155', padding: '6px 8px', width: '50px', textAlign: 'center' }}>#</th>
                            {excelInspection.headers.map((h, hIdx) => (
                              <th key={hIdx} style={{ border: '1px solid #334155', padding: '6px 10px', textAlign: 'left', whiteSpace: 'nowrap' }}>
                                Col {hIdx + 1}: {h}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {filteredRawRows.length === 0 ? (
                            <tr>
                              <td colSpan={(excelInspection.headers?.length || 1) + 1} style={{ textAlign: 'center', padding: '24px', color: '#94a3b8' }}>
                                No rows matching search filter
                              </td>
                            </tr>
                          ) : (
                            filteredRawRows.map((r, rIdx) => {
                              const isEven = rIdx % 2 === 0;
                              return (
                                <tr
                                  key={rIdx}
                                  style={{
                                    background: isEven ? '#0a1525' : '#0d1b2e',
                                    color: '#cbd5e1',
                                    fontSize: '11.5px'
                                  }}
                                >
                                  <td style={{ border: '1px solid #1e314d', padding: '6px 8px', textAlign: 'center', color: '#64748b' }}>
                                    {r.row}
                                  </td>
                                  {r.cells.map((cellVal, cIdx) => (
                                    <td key={cIdx} style={{ border: '1px solid #1e314d', padding: '6px 10px', textAlign: 'left', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '240px' }}>
                                      {cellVal || '—'}
                                    </td>
                                  ))}
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    )}
                  </div>
                )}
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
                  <span>1</span> Upload
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 2 ? 'active' : ''} ${progressStep > 2 ? 'done' : ''}`}>
                  <span>2</span> Parse Listings
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 3 ? 'active' : ''} ${progressStep > 3 ? 'done' : ''}`}>
                  <span>3</span> Standardize (11 Cols)
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
              disabled={!file || converting || inspecting}
              id="btn-convert-excel"
              style={{ minWidth: '280px', padding: '12px 24px', fontSize: '14px', fontWeight: 700, background: 'linear-gradient(135deg, #16a34a, #22c55e)' }}
            >
              {converting ? (
                <>
                  <span className="scooh-spin">🔄</span> Standardizing Excel...
                </>
              ) : (
                <>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                  </svg>
                  Convert to Standardized Master Excel (.xlsx)
                </>
              )}
            </button>
          </div>
        </div>
      ) : (
        /* Results Area with Excel Simulation Preview */
        <div className="scooh-ppt-results-view">
          {/* Conversion Success Overview Card */}
          <div className="scooh-panel scooh-ppt-success-card" style={{ borderLeft: '4px solid #22c55e' }}>
            <div className="scooh-ppt-success-left">
              <div className="scooh-ppt-success-badge" style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80' }}>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                </svg>
              </div>
              <div>
                <h3 style={{ margin: '0 0 4px', color: '#fff', fontSize: '17px' }}>
                  {sites.length} Billboard Sites Converted into Master Excel
                </h3>
                <p style={{ margin: 0, color: 'var(--mb-muted)', fontSize: '12px' }}>
                  Clean 11-column Jade Blue Master Standard with Media Buzz Header Banner &amp; Official 7-Point Terms attached.
                  {distinctCities && (
                    <span style={{ color: '#4ade80', marginLeft: '6px' }}>
                      City: <strong>{distinctCities}</strong>.
                    </span>
                  )}
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
                  title="Click to rename file before downloading"
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
                <span style={{ fontSize: '11px', color: '#4ade80', fontWeight: 700 }}>.xlsx</span>
              </div>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleDownloadExcel}
                id="btn-download-excel-conv-xls-main"
                style={{ background: 'linear-gradient(135deg, #16a34a, #22c55e)', color: '#ffffff', fontWeight: 700, padding: '10px 18px', fontSize: '13px' }}
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                Download Master Excel (.xlsx)
              </button>
            </div>
          </div>

          {/* KPI Summary Cards */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '12px', marginBottom: '16px' }}>
            <div className="scooh-panel" style={{ padding: '12px 16px', background: '#091524', border: '1px solid #1e293b', borderRadius: '8px' }}>
              <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px', fontWeight: 600 }}>Total Billboard Sites</span>
              <div style={{ fontSize: '20px', fontWeight: 800, color: '#38bdf8', marginTop: '4px' }}>{sites.length} Sites</div>
            </div>
            <div className="scooh-panel" style={{ padding: '12px 16px', background: '#091524', border: '1px solid #1e293b', borderRadius: '8px' }}>
              <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px', fontWeight: 600 }}>Total Display Area</span>
              <div style={{ fontSize: '20px', fontWeight: 800, color: '#4ade80', marginTop: '4px' }}>{totalSqFt} Sq. Ft.</div>
            </div>
            <div className="scooh-panel" style={{ padding: '12px 16px', background: '#091524', border: '1px solid #1e293b', borderRadius: '8px' }}>
              <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px', fontWeight: 600 }}>Primary Market / City</span>
              <div style={{ fontSize: '18px', fontWeight: 800, color: '#f8fafc', marginTop: '4px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{distinctCities}</div>
            </div>
            <div className="scooh-panel" style={{ padding: '12px 16px', background: '#091524', border: '1px solid #1e293b', borderRadius: '8px' }}>
              <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px', fontWeight: 600 }}>Master Format</span>
              <div style={{ fontSize: '16px', fontWeight: 800, color: '#FFC200', marginTop: '4px' }}>11 Master Columns</div>
            </div>
          </div>

          {/* Interactive Excel Sheet Preview */}
          <div className="scooh-panel" style={{ background: '#081220', border: '1px solid #1e2f47', borderRadius: '12px', padding: '18px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px', flexWrap: 'wrap', gap: '10px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', padding: '6px 8px', borderRadius: '6px' }}>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                  </svg>
                </div>
                <div>
                  <h3 style={{ margin: 0, color: '#fff', fontSize: '15px', fontWeight: 700 }}>
                    Master Excel Spreadsheet Preview ({sites.length} Rows)
                  </h3>
                  <span className="scooh-footnote" style={{ color: '#94a3b8', fontSize: '11.5px' }}>
                    Authentic simulation of the exported 11-column Excel workbook with Media Buzz branding and terms
                  </span>
                </div>
              </div>

              {/* Live search input */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <input
                  type="text"
                  value={tableFilter}
                  onChange={(e) => setTableFilter(e.target.value)}
                  placeholder="Quick search in sites..."
                  style={{
                    background: '#0d1829',
                    border: '1px solid #1e2f47',
                    borderRadius: '6px',
                    padding: '6px 12px',
                    color: '#fff',
                    fontSize: '12px',
                    width: '220px'
                  }}
                />
                <span className="scooh-tag" style={{ background: 'rgba(56, 189, 248, 0.12)', color: '#38bdf8', border: '1px solid rgba(56, 189, 248, 0.25)', fontSize: '11px' }}>
                  Showing {filteredSites.length} of {sites.length}
                </span>
              </div>
            </div>

            {/* Simulated Excel Sheet Container */}
            <div style={{ overflowX: 'auto', borderRadius: '8px', border: '1px solid #233b5c', boxShadow: '0 8px 30px rgba(0,0,0,0.5)' }}>
              <table style={{ width: '100%', minWidth: '1100px', borderCollapse: 'collapse', fontFamily: 'Calibri, "Segoe UI", sans-serif' }}>
                <thead>
                  {/* Row 1: Executive Media Buzz Brand Banner */}
                  <tr style={{ background: '#071C35', height: '52px', borderBottom: '3px solid #FFC200' }}>
                    <th
                      colSpan={11}
                      style={{
                        padding: '0 18px',
                        textAlign: 'center',
                        position: 'relative',
                        color: '#FFC200',
                        fontSize: '14.5px',
                        fontWeight: 800,
                        letterSpacing: '0.6px'
                      }}
                    >
                      <span>{customFileName || 'MEDIA BUZZ — SITES SPECIFICATION'}</span>
                      <div style={{ position: 'absolute', right: '16px', top: '50%', transform: 'translateY(-50%)', display: 'flex', alignItems: 'center' }}>
                        <img
                          src={MEDIA_BUZZ_LOGO_BASE64.startsWith('data:') ? MEDIA_BUZZ_LOGO_BASE64 : `data:image/png;base64,${MEDIA_BUZZ_LOGO_BASE64}`}
                          alt="Media Buzz Logo"
                          style={{ height: '36px', maxWidth: '140px', objectFit: 'contain' }}
                        />
                      </div>
                    </th>
                  </tr>

                  {/* Row 2: 11 Columns in Jade Blue Bright Yellow */}
                  <tr style={{ background: '#FFFF00', color: '#000000', height: '32px', fontSize: '11.5px', fontWeight: 800 }}>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '60px', textAlign: 'center' }}>SR NO</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '110px', textAlign: 'center' }}>City</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 12px', minWidth: '280px', textAlign: 'left' }}>LOCATION</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '100px', textAlign: 'center' }}>MEDIA</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '70px', textAlign: 'center' }}>LIGHT</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '60px', textAlign: 'center' }}>W</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '60px', textAlign: 'center' }}>H</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '80px', textAlign: 'center' }}>SQ FT</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 8px', width: '110px', textAlign: 'center' }}>AVAILABLITY</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 12px', width: '130px', textAlign: 'right' }}>Selling Amount</th>
                    <th style={{ border: '1px solid #808080', padding: '6px 10px', width: '160px', textAlign: 'left' }}>Latitude Longitude</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSites.map((site, sIdx) => {
                    const isEven = sIdx % 2 === 0;
                    return (
                      <tr
                        key={sIdx}
                        style={{
                          background: isEven ? '#0a1525' : '#0d1b2e',
                          color: '#e2e8f0',
                          fontSize: '12px',
                          transition: 'background 0.15s ease'
                        }}
                        onMouseEnter={(e) => { e.currentTarget.style.background = '#132845'; }}
                        onMouseLeave={(e) => { e.currentTarget.style.background = isEven ? '#0a1525' : '#0d1b2e'; }}
                      >
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center', color: '#94a3b8' }}>{site.srNo}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center', color: '#cbd5e1' }}>{site.city}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 12px', textAlign: 'left', fontWeight: 600, color: '#f8fafc' }}>{site.location}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center', color: '#cbd5e1' }}>{site.media}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center', color: '#cbd5e1' }}>{site.light}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center', color: '#e2e8f0' }}>{site.w}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center', color: '#e2e8f0' }}>{site.h}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center', color: '#e2e8f0' }}>{site.sqFt}</td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 8px', textAlign: 'center' }}>
                          <span style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 700 }}>
                            {site.availability || 'Available'}
                          </span>
                        </td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 12px', textAlign: 'right', fontWeight: 700, color: '#FFC200' }}>
                          {site.rate || '—'}
                        </td>
                        <td style={{ border: '1px solid #1e314d', padding: '7px 10px', textAlign: 'left', color: '#94a3b8', fontFamily: 'monospace', fontSize: '11px' }}>
                          {site.coords || '—'}
                        </td>
                      </tr>
                    );
                  })}

                  {/* Spacer Row before Terms */}
                  <tr style={{ height: '18px', background: '#060d17' }}>
                    <td colSpan={11} style={{ border: 'none' }}></td>
                  </tr>

                  {/* Terms Header Row */}
                  <tr style={{ background: '#071C35', height: '34px', borderBottom: '2px solid #FFC200' }}>
                    <td
                      colSpan={11}
                      style={{
                        padding: '6px 12px',
                        textAlign: 'center',
                        color: '#FFC200',
                        fontWeight: 800,
                        fontSize: '12px',
                        letterSpacing: '0.6px',
                        border: '1px solid #234D7D'
                      }}
                    >
                      TERMS &amp; CONDITIONS
                    </td>
                  </tr>

                  {/* 7 Official Software Terms Rows */}
                  {OFFICIAL_SOFTWARE_TERMS.map((term, tIdx) => {
                    const isEvenTerm = tIdx % 2 === 0;
                    return (
                      <tr
                        key={tIdx}
                        style={{
                          background: isEvenTerm ? '#fffff7' : '#ffffff',
                          color: '#071C35',
                          height: '28px',
                          fontSize: '11.5px',
                          border: '1px solid #CBD5E1'
                        }}
                      >
                        <td
                          colSpan={11}
                          style={{
                            padding: '6px 16px',
                            textAlign: 'center',
                            border: '1px solid #CBD5E1',
                            color: '#071C35',
                            fontWeight: 600
                          }}
                        >
                          {term}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Bottom Actions Bar */}
            <div style={{ marginTop: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
              <span style={{ fontSize: '12px', color: '#94a3b8' }}>
                Excel output conforms with <code>Jade Blue 16-10-2026.xlsx</code> standards.
              </span>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleDownloadExcel}
                id="btn-download-master-excel-bottom"
                style={{ background: 'linear-gradient(135deg, #16a34a, #22c55e)', color: '#ffffff', fontWeight: 700, padding: '10px 22px', fontSize: '13px' }}
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                Download Master Excel (.xlsx)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* COLUMN MAPPING MODAL (Mandatory Master Excel Column Alignment) */}
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
                  value={selectedSheet || excelInspection.activeSheet || ''}
                  onChange={(e) => handleSheetChange(e.target.value)}
                  style={{ background: '#111e33', color: '#38bdf8', border: '1px solid #0284c7', borderRadius: '6px', padding: '6px 12px', fontSize: '12.5px', fontWeight: 700 }}
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
                onClick={handleApplyCustomMapping}
                style={{ minHeight: '38px', padding: '0 20px', fontSize: '12.5px', fontWeight: 800 }}
              >
                ✓ Confirm &amp; Apply Mapping
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
                  <span>Official Media Buzz Software Terms &amp; Conditions</span>
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
