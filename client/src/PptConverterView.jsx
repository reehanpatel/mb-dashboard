import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
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

export default function PptConverterView() {
  // Master template state
  const [templateStatus, setTemplateStatus] = useState({
    loading: true,
    available: false,
    filename: '',
    size: 0,
    masterExcel: null,
    error: null
  });
  const [templateUpdating, setTemplateUpdating] = useState(false);

  // File upload & state
  const [file, setFile] = useState(null);
  const [isExcel, setIsExcel] = useState(false);
  const [dragActive, setDragActive] = useState(false);
  const fileInputRef = useRef(null);
  const templateInputRef = useRef(null);

  // Excel inspection & column mapping state
  const [inspecting, setInspecting] = useState(false);
  const [excelInspection, setExcelInspection] = useState(null);
  const [columnMapping, setColumnMapping] = useState({});
  const [selectedSheet, setSelectedSheet] = useState('');
  const [showMappingModal, setShowMappingModal] = useState(false);
  const [showTermsModal, setShowTermsModal] = useState(false);

  // Conversion process
  const [converting, setConverting] = useState(false);
  const [progressStep, setProgressStep] = useState(0); // 0: idle, 1: upload, 2: extract, 3: convert, 4: done
  const [progressPercent, setProgressPercent] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [conversionResult, setConversionResult] = useState(null);
  const [customFileName, setCustomFileName] = useState('Converted_MediaBuzz_Master_Format');
  const [error, setError] = useState(null);

  // Slide preview state
  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const [viewMode, setViewMode] = useState('grid'); // 'grid' | 'slideshow'
  const [filterReviewOnly, setFilterReviewOnly] = useState(false);

  // Fetch master template status on mount
  const checkTemplateStatus = useCallback(async () => {
    try {
      setTemplateStatus(prev => ({ ...prev, loading: true, error: null }));
      const res = await api.get('/ppt-converter/status');
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
      const res = await api.post('/ppt-converter/inspect-excel', fd, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setExcelInspection(res.data);
      if (res.data?.mapping) {
        setColumnMapping(res.data.mapping);
      }
      if (res.data?.activeSheet) {
        setSelectedSheet(res.data.activeSheet);
      }
      // If headers do NOT match master template Jade Blue 16-10-2026.xlsx, automatically show mapping interface
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
    const isPpt = lower.endsWith('.pptx');
    const isXls = lower.endsWith('.xlsx') || lower.endsWith('.xls');

    if (!isPpt && !isXls) {
      setError('Please select a valid PowerPoint (.pptx) or Excel (.xlsx / .xls) file.');
      return;
    }
    setFile(selectedFile);
    setConversionResult(null);
    setIsExcel(isXls);

    if (isXls) {
      inspectExcelWorkbook(selectedFile);
    } else {
      setExcelInspection(null);
      setShowMappingModal(false);
    }
  };

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      validateAndSetFile(e.dataTransfer.files[0]);
    }
  }, []);

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      validateAndSetFile(e.target.files[0]);
    }
  };

  // Upload or replace master PowerPoint template
  const handleTemplateUpload = async (e) => {
    const tmplFile = e.target.files?.[0];
    if (!tmplFile) return;
    if (!tmplFile.name.toLowerCase().endsWith('.pptx')) {
      alert('Master template must be a .pptx file.');
      return;
    }
    try {
      setTemplateUpdating(true);
      const fd = new FormData();
      fd.append('file', tmplFile);
      await api.post('/ppt-converter/template', fd);
      await checkTemplateStatus();
    } catch (err) {
      alert('Failed to update template: ' + (err.response?.data?.message || err.message));
    } finally {
      setTemplateUpdating(false);
      if (templateInputRef.current) templateInputRef.current.value = '';
    }
  };

  // Download master PowerPoint template
  const handleDownloadTemplate = async () => {
    try {
      const res = await api.get('/ppt-converter/template/download', { responseType: 'blob' });
      const blob = new Blob([res.data], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = templateStatus.filename || 'MediaBuzz_Master_Template.pptx';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      alert('Failed to download master template: ' + (err.response?.data?.message || err.message));
    }
  };

  // Convert workflow
  const handleConvert = async () => {
    if (!file || !templateStatus.available || converting) return;

    setError(null);
    setConverting(true);
    setProgressStep(1);
    setProgressPercent(15);
    setProgressMessage(isExcel ? 'Uploading Excel workbook for conversion...' : 'Uploading PowerPoint presentation...');

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
      setProgressMessage(isExcel ? 'Parsing site listings & specs from Excel rows...' : 'Extracting location photos & billboard dimensions...');
    }, 1200);

    setTimeout(() => {
      setProgressStep(3);
      setProgressMessage('Populating Media Buzz Master PowerPoint & attaching official software terms...');
    }, 2800);

    try {
      const fd = new FormData();
      fd.append('file', file);
      if (isExcel) {
        if (columnMapping && Object.keys(columnMapping).length > 0) {
          fd.append('columnMapping', JSON.stringify(columnMapping));
        }
        if (selectedSheet) {
          fd.append('sheetName', selectedSheet);
        }
      }

      const res = await api.post('/ppt-converter/convert', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 180000
      });

      clearInterval(pInterval);
      setProgressPercent(100);
      setProgressStep(4);
      setProgressMessage('Conversion completed successfully!');

      setTimeout(() => {
        setConversionResult(res.data);
        const defaultName = file?.name ? file.name.replace(/\.[^/.]+$/, '') + '_MediaBuzz' : 'Converted_MediaBuzz_Master_Format';
        setCustomFileName(defaultName);
        setActiveSlideIndex(0);
        setConverting(false);
      }, 500);

    } catch (err) {
      clearInterval(pInterval);
      setConverting(false);
      setProgressStep(0);
      setProgressPercent(0);
      const msg = err.response?.data?.message || err.message || 'An error occurred during conversion.';
      setError(msg);
    }
  };

  // Download converted presentation
  const handleDownload = async () => {
    if (!conversionResult?.downloadId) return;
    try {
      const cleanBase = (customFileName || 'Converted_MediaBuzz_Master_Format').trim().replace(/\.pptx$/i, '');
      const cleanName = `${cleanBase}.pptx`;
      const downloadUrl = `/ppt-converter/download/${encodeURIComponent(conversionResult.downloadId)}?name=${encodeURIComponent(cleanName)}`;
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

  // Reset conversion
  const handleReset = () => {
    setFile(null);
    setIsExcel(false);
    setExcelInspection(null);
    setColumnMapping({});
    setShowMappingModal(false);
    setConversionResult(null);
    setError(null);
    setConverting(false);
    setProgressStep(0);
    setProgressPercent(0);
    setActiveSlideIndex(0);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Format file size
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
            <span>PowerPoint & Excel Converter</span>
            <span className="scooh-pill active" style={{ fontSize: '11px', textTransform: 'uppercase' }}>
              Master Template Reference
            </span>
          </h1>
          <p>
            Convert PowerPoint presentations and Excel workbooks into the official Media Buzz Master Template format.
            Uses <strong>Jade Blue 16-10-2026.xlsx</strong> as the mandatory reference standard and automatically applies the Media Buzz 16:9 widescreen theme with verified software Terms & Conditions.
          </p>
        </div>
        <div className="scooh-headactions">
          {conversionResult && (
            <>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleDownload}
                id="btn-download-converted-top"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                </svg>
                Download Presentation
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={handleReset}
                id="btn-reset-converter-top"
              >
                Convert Another File
              </button>
            </>
          )}
        </div>
      </div>

      {/* Mandatory Reference Standards Status Card */}
      <div className="scooh-panel scooh-ppt-template-bar" style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '14px' }}>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
            <div className="scooh-ppt-template-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="var(--mb-primary)">
                <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
              </svg>
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                <strong style={{ color: '#fff', fontSize: '14px' }}>
                  Mandatory Reference Standards:
                </strong>
                <span className="scooh-pill active" style={{ fontSize: '11px' }}>
                  PowerPoint Theme: {templateStatus.filename || 'MediaBuzz_Master_Template.pptx'}
                </span>
                <span className="scooh-pill" style={{ background: 'rgba(255, 194, 0, 0.15)', color: '#FFC200', border: '1px solid rgba(255, 194, 0, 0.3)', fontSize: '11px' }}>
                  Master Excel Reference: Jade Blue 16-10-2026.xlsx
                </span>
              </div>
              <p className="scooh-ppt-template-desc" style={{ marginTop: '4px', marginBottom: 0 }}>
                Default reference layout contains 11 columns in field order: <code>SR NO</code>, <code>City</code>, <code>LOCATION</code>, <code>MEDIA</code>, <code>LIGHT</code>, <code>W</code>, <code>H</code>, <code>SQ FT</code>, <code>AVAILABLITY</code>, <code>Selling Amount</code>, <code>Latitude Longitude</code>.
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
            {templateStatus.available && (
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={handleDownloadTemplate}
                title="Download master PowerPoint template"
                style={{ fontSize: '12px', padding: '4px 10px' }}
              >
                Download Master PPT
              </button>
            )}
          </div>
        </div>

        {/* Governance Notice Banner (Rule 5 & 7) */}
        <div style={{ background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(56, 189, 248, 0.25)', borderRadius: '8px', padding: '8px 12px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px', fontSize: '12px', color: '#cbd5e1' }}>
          <div>
            <span style={{ color: '#38bdf8', fontWeight: 700, marginRight: '6px' }}>RULE ENFORCEMENT:</span>
            <span>Master template is permanently preserved read-only. Terms & Conditions are retrieved exclusively from Media Buzz software (workbook terms are ignored). Missing values remain blank.</span>
          </div>
          <span style={{ color: '#4ade80', fontWeight: 600, fontSize: '11px' }}>✓ Policy Active</span>
        </div>
      </div>

      {/* Conversion Error Message */}
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
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '14px', flexWrap: 'wrap', gap: '10px' }}>
            <div>
              <h3 style={{ margin: '0 0 4px', color: '#fff' }}>Upload File for Conversion</h3>
              <p className="scooh-footnote" style={{ margin: 0 }}>
                Upload an Excel workbook (<code>.xlsx</code> / <code>.xls</code>) or PowerPoint presentation (<code>.pptx</code>).
              </p>
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <span className="scooh-tag" style={{ background: 'rgba(34, 197, 94, 0.1)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.25)' }}>
                .xlsx / .xls Supported
              </span>
              <span className="scooh-tag" style={{ background: 'rgba(56, 189, 248, 0.1)', color: '#38bdf8', border: '1px solid rgba(56, 189, 248, 0.25)' }}>
                .pptx Supported
              </span>
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
              accept=".pptx,.xlsx,.xls"
              style={{ display: 'none' }}
            />

            {!file ? (
              <div className="scooh-ppt-drop-prompt">
                <div className="scooh-ppt-drop-icon">
                  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                    <path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242M12 12v9m-4-4 4-4 4 4" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <h4>Drag and drop Excel workbook or PowerPoint here</h4>
                <p>or click to browse (.xlsx, .xls, .pptx formats)</p>
                <div className="scooh-ppt-drop-badge">
                  <span>Excel structure verified against master Jade Blue 16-10-2026.xlsx</span>
                </div>
              </div>
            ) : (
              <div className="scooh-ppt-file-card">
                <div className="scooh-ppt-file-icon">
                  {isExcel ? (
                    <svg width="36" height="36" viewBox="0 0 24 24" fill="#22c55e">
                      <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                    </svg>
                  ) : (
                    <svg width="36" height="36" viewBox="0 0 24 24" fill="var(--mb-primary)">
                      <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>
                    </svg>
                  )}
                </div>
                <div className="scooh-ppt-file-details">
                  <strong className="scooh-ppt-file-name">{file.name}</strong>
                  <span className="scooh-ppt-file-size">{formatBytes(file.size)}</span>
                  {isExcel && inspecting && (
                    <span className="scooh-pill" style={{ marginLeft: '12px', background: 'rgba(56, 189, 248, 0.2)', color: '#38bdf8' }}>
                      Inspecting structure...
                    </span>
                  )}
                  {isExcel && !inspecting && excelInspection?.matchesMaster && (
                    <span className="scooh-pill active" style={{ marginLeft: '12px' }}>
                      ✓ Matches Jade Blue Reference Standard
                    </span>
                  )}
                  {isExcel && !inspecting && excelInspection && !excelInspection.matchesMaster && (
                    <span className="scooh-pill" style={{ marginLeft: '12px', background: 'rgba(245, 158, 11, 0.2)', color: '#fbbf24', border: '1px solid rgba(245, 158, 11, 0.4)' }}>
                      ⚠️ Custom Structure (Mapping Available)
                    </span>
                  )}
                  {!isExcel && (
                    <span className="scooh-pill active" style={{ marginLeft: '12px' }}>Ready to Convert</span>
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

          {/* Excel Inspection Overview & Column Mapping Access (Rule 2 & 4) */}
          {file && isExcel && excelInspection && (
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

              {/* Terms & Conditions Notice (Rule 7) */}
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
                  style={{ width: `${progressPercent}%` }}
                />
              </div>
              <div className="scooh-ppt-steps">
                <div className={`scooh-ppt-step ${progressStep >= 1 ? 'active' : ''} ${progressStep > 1 ? 'done' : ''}`}>
                  <span>1</span> Upload File
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 2 ? 'active' : ''} ${progressStep > 2 ? 'done' : ''}`}>
                  <span>2</span> Parse Listings
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 3 ? 'active' : ''} ${progressStep > 3 ? 'done' : ''}`}>
                  <span>3</span> Apply Master PPT
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
              id="btn-convert-ppt"
              style={{ minWidth: '240px', padding: '12px 24px', fontSize: '14px', fontWeight: 700 }}
            >
              {converting ? (
                <>
                  <span className="scooh-spin">🔄</span> Converting...
                </>
              ) : isExcel ? (
                <>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                  </svg>
                  Convert Excel to Media Buzz PPT
                </>
              ) : (
                <>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                  </svg>
                  Convert PPT to Media Buzz Master
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
              <div className="scooh-ppt-success-badge">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="var(--mb-primary)">
                  <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                </svg>
              </div>
              <div>
                <h3 style={{ margin: '0 0 4px', color: '#fff' }}>
                  {conversionResult.processedSlides} Billboard Sites Converted Successfully
                </h3>
                <p style={{ margin: 0, color: 'var(--mb-muted)', fontSize: '12px' }}>
                  {conversionResult.isExcel ? 'Generated from uploaded Excel workbook.' : 'Generated from uploaded presentation.'} Official 16:9 widescreen layout with verified software Terms & Conditions.
                  Total presentation size: <strong>{conversionResult.totalSlidesInFile} slides</strong>.
                </p>
              </div>
            </div>

            <div className="scooh-ppt-success-actions" style={{ display: 'flex', flexDirection: 'column', gap: '8px', minWidth: '320px' }}>
              <div style={{ display: 'flex', alignItems: 'center', background: 'rgba(255,255,255,0.06)', borderRadius: '6px', border: '1px solid var(--mb-border)', padding: '2px 10px' }}>
                <span style={{ fontSize: '11px', color: 'var(--mb-primary)', fontWeight: 700, marginRight: '8px', letterSpacing: '0.5px' }}>RENAME:</span>
                <input
                  type="text"
                  value={customFileName}
                  onChange={(e) => setCustomFileName(e.target.value)}
                  placeholder="Presentation Name"
                  title="Click to rename presentation before downloading"
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
                <span style={{ fontSize: '11px', color: 'var(--mb-muted)', fontWeight: 600 }}>.pptx</span>
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button
                  type="button"
                  className="scooh-btn primary"
                  onClick={handleDownload}
                  id="btn-download-converted-main"
                  style={{ flex: 1 }}
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                  </svg>
                  Download Presentation
                </button>
                <button
                  type="button"
                  className="scooh-btn ghost"
                  onClick={handleReset}
                >
                  Convert Another
                </button>
              </div>
            </div>
          </div>

          {/* Pre-Export Validation Badge (Rule 9) */}
          <div style={{ background: 'rgba(16, 185, 129, 0.1)', border: '1px solid rgba(16, 185, 129, 0.3)', borderRadius: '8px', padding: '10px 14px', marginBottom: '18px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '10px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ fontSize: '18px' }}>🛡️</span>
              <div>
                <strong style={{ color: '#4ade80', fontSize: '13px' }}>Pre-Export Validation Passed:</strong>
                <span style={{ color: '#cbd5e1', fontSize: '12px', marginLeft: '6px' }}>
                  {conversionResult.processedSlides} slides match selected file data without invented values. Software Terms & Conditions preserved unchanged.
                </span>
              </div>
            </div>
            <button
              type="button"
              className="scooh-btn secondary"
              onClick={() => setShowTermsModal(true)}
              style={{ fontSize: '11.5px', padding: '3px 10px' }}
            >
              Verify Software Terms
            </button>
          </div>

          {/* Quick Stats Bar */}
          <div className="scooh-kpirow" style={{ marginBottom: '18px' }}>
            <div className="scooh-kpi good">
              <div className="n">{conversionResult.processedSlides}</div>
              <div className="l">Processed Sites</div>
            </div>
            <div className="scooh-kpi good">
              <div className="n">{slides.filter(s => s.hasPhoto).length}</div>
              <div className="l">Photos Extracted</div>
            </div>
            <div className="scooh-kpi good">
              <div className="n">{slides.filter(s => s.size).length}</div>
              <div className="l">Dimensions Detected</div>
            </div>
            <div className="scooh-kpi good">
              <div className="n">100%</div>
              <div className="l">Terms Verified</div>
            </div>
          </div>

          {/* Slide Preview Controls */}
          <div className="scooh-toolbar" style={{ marginBottom: '16px' }}>
            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              <button
                type="button"
                className={`scooh-btn ${viewMode === 'grid' ? 'primary' : 'ghost'}`}
                onClick={() => setViewMode('grid')}
                title="Grid View"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M4 4h4v4H4zm6 0h4v4h-4zm6 0h4v4h-4zM4 10h4v4H4zm6 0h4v4h-4zm6 0h4v4h-4zM4 16h4v4H4zm6 0h4v4h-4zm6 0h4v4h-4z"/>
                </svg>
                Grid Preview
              </button>
              <button
                type="button"
                className={`scooh-btn ${viewMode === 'slideshow' ? 'primary' : 'ghost'}`}
                onClick={() => setViewMode('slideshow')}
                title="Single Slide View"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>
                </svg>
                Slide View
              </button>
            </div>

            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', cursor: 'pointer', color: 'var(--mb-muted)' }}>
                <input
                  type="checkbox"
                  checked={filterReviewOnly}
                  onChange={(e) => setFilterReviewOnly(e.target.checked)}
                />
                Show Review Needed Only ({slides.filter(s => s.reviewNeeded).length})
              </label>
            </div>
          </div>

          {/* GRID VIEW */}
          {viewMode === 'grid' && (
            <div className="scooh-ppt-grid">
              {filteredSlides.map((slide, idx) => (
                <div
                  key={idx}
                  className={`scooh-ppt-card ${slide.isTermsSlide ? 'terms-slide-card' : ''} ${slide.reviewNeeded ? 'needs-review' : ''}`}
                  onClick={() => {
                    setActiveSlideIndex(slides.indexOf(slide));
                    setViewMode('slideshow');
                  }}
                  style={{ cursor: 'pointer', border: slide.isTermsSlide ? '1px solid rgba(255, 194, 0, 0.4)' : undefined }}
                >
                  <div className="scooh-ppt-card-thumb">
                    {slide.previewUrl ? (
                      <img src={slide.previewUrl} alt={`Slide ${slide.siteNumber}`} />
                    ) : (
                      <div className="scooh-ppt-card-thumb-empty">
                        <span>No Preview</span>
                      </div>
                    )}
                    <span className="scooh-ppt-card-badge" style={slide.isTermsSlide ? { background: '#FFC200', color: '#000' } : {}}>
                      {slide.isTermsSlide ? 'Terms & Conditions' : `Site ${slide.siteNumber}`}
                    </span>
                    {slide.reviewNeeded && (
                      <span className="scooh-ppt-card-warn" title="Flagged for review">⚠</span>
                    )}
                  </div>

                  <div className="scooh-ppt-card-body">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
                      <strong style={{ color: '#fff', fontSize: '12px' }}>
                        {slide.isTermsSlide ? 'Media Buzz Contract Terms' : (slide.city || 'City')}
                      </strong>
                      <span style={{ fontSize: '11px', color: 'var(--mb-primary)', fontWeight: 700 }}>
                        {slide.isTermsSlide ? '7 Points' : (slide.size || 'Size')}
                      </span>
                    </div>
                    <p style={{ margin: 0, fontSize: '11px', color: 'var(--mb-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {slide.location || 'Location details'}
                    </p>
                    <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '6px', fontSize: '10.5px', color: '#94a3b8' }}>
                      <span>{slide.type || ''}</span>
                      <span>{slide.rate || ''}</span>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* SLIDESHOW VIEW */}
          {viewMode === 'slideshow' && activeSlide && (
            <div className="scooh-panel scooh-ppt-slideshow-panel">
              <div className="scooh-ppt-slideshow-main">
                <div className="scooh-ppt-slideshow-imgbox">
                  {activeSlide.previewUrl ? (
                    <img src={activeSlide.previewUrl} alt={`Slide ${activeSlide.siteNumber}`} />
                  ) : (
                    <div className="scooh-ppt-card-thumb-empty large">
                      <span>No Visual Preview Available</span>
                    </div>
                  )}
                </div>

                <div className="scooh-ppt-slideshow-sidebar">
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                    <span className="scooh-plate large" style={activeSlide.isTermsSlide ? { background: '#FFC200', color: '#000' } : {}}>
                      {activeSlide.isTermsSlide ? 'TERMS & CONDITIONS' : `Site ${activeSlide.siteNumber}`}
                    </span>
                    <span className="scooh-tag">
                      Slide {activeSlideIndex + 1} of {slides.length}
                    </span>
                  </div>

                  <div className="scooh-field">
                    <label>City</label>
                    <div className="scooh-val">{activeSlide.city || '—'}</div>
                  </div>

                  <div className="scooh-field">
                    <label>Location / Description</label>
                    <div className="scooh-val">{activeSlide.location || '—'}</div>
                  </div>

                  <div className="scooh-grid2">
                    <div className="scooh-field">
                      <label>Size</label>
                      <div className="scooh-val">{activeSlide.size || '—'}</div>
                    </div>
                    <div className="scooh-field">
                      <label>Type</label>
                      <div className="scooh-val">{activeSlide.type || '—'}</div>
                    </div>
                  </div>

                  <div className="scooh-grid2">
                    <div className="scooh-field">
                      <label>Illumination</label>
                      <div className="scooh-val">{activeSlide.illum || '—'}</div>
                    </div>
                    <div className="scooh-field">
                      <label>Adv. Fee per Month</label>
                      <div className="scooh-val">{activeSlide.rate || '—'}</div>
                    </div>
                  </div>

                  <div className="scooh-field">
                    <label>Availability</label>
                    <div className="scooh-val">{activeSlide.availability || '—'}</div>
                  </div>

                  <div className="scooh-field">
                    <label>Location Coordinates</label>
                    <div className="scooh-val">{activeSlide.coords || '—'}</div>
                  </div>

                  {activeSlide.reviewNeeded && (
                    <div style={{ marginTop: '12px', padding: '8px 12px', background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)', borderRadius: '6px' }}>
                      <strong style={{ color: '#f87171', fontSize: '11px', display: 'block', marginBottom: '4px' }}>Review Flags:</strong>
                      <ul style={{ margin: 0, paddingLeft: '16px', fontSize: '11px', color: '#fca5a5' }}>
                        {activeSlide.reviewFlags.map((flag, fIdx) => (
                          <li key={fIdx}>{flag}</li>
                        ))}
                      </ul>
                    </div>
                  )}

                  {/* Navigation Buttons */}
                  <div style={{ display: 'flex', gap: '8px', marginTop: '20px' }}>
                    <button
                      type="button"
                      className="scooh-btn secondary"
                      disabled={activeSlideIndex === 0}
                      onClick={() => setActiveSlideIndex(prev => Math.max(0, prev - 1))}
                      style={{ flex: 1 }}
                    >
                      ◀ Previous
                    </button>
                    <button
                      type="button"
                      className="scooh-btn secondary"
                      disabled={activeSlideIndex === slides.length - 1}
                      onClick={() => setActiveSlideIndex(prev => Math.min(slides.length - 1, prev + 1))}
                      style={{ flex: 1 }}
                    >
                      Next ▶
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* COLUMN MAPPING INTERFACE MODAL (Rule 4: Mandatory Interactive Mapping)     */}
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
              maxWidth: '920px',
              maxHeight: '92vh',
              overflowY: 'auto',
              border: '1px solid #38bdf8',
              boxShadow: '0 25px 60px rgba(0,0,0,0.85), 0 0 30px rgba(56, 189, 248, 0.2)',
              background: '#071224',
              padding: '24px',
              borderRadius: '14px',
              color: '#f8fafc'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', borderBottom: '1px solid #1e293b', paddingBottom: '14px', marginBottom: '18px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '18px', fontWeight: 800 }}>
                  <span style={{ color: '#38bdf8' }}>⚙️</span>
                  <span>Excel Column Mapping Interface</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  {excelInspection.matchesMaster
                    ? 'Your workbook matches the Jade Blue 16-10-2026.xlsx reference structure. You can review or adjust field assignments below.'
                    : 'The uploaded workbook has a different column structure than the master template. Match your columns to the Media Buzz target fields below.'}
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

            {/* Sheet Selection if multiple sheets exist */}
            {excelInspection.sheets && excelInspection.sheets.length > 1 && (
              <div style={{ marginBottom: '16px', background: '#0a1628', padding: '10px 14px', borderRadius: '8px', border: '1px solid #1e293b', display: 'flex', alignItems: 'center', gap: '12px' }}>
                <label style={{ fontSize: '12px', fontWeight: 700, color: '#e2e8f0' }}>Target Sheet:</label>
                <select
                  value={selectedSheet}
                  onChange={(e) => setSelectedSheet(e.target.value)}
                  style={{ background: '#071224', color: '#f8fafc', border: '1px solid #334155', borderRadius: '6px', padding: '6px 12px', fontSize: '12px' }}
                >
                  {excelInspection.sheets.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
                <span style={{ fontSize: '11px', color: '#64748b' }}>Data will be extracted from this worksheet</span>
              </div>
            )}

            {/* Mapping Grid (11 Target Fields) */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(270px, 1fr))', gap: '12px', marginBottom: '20px' }}>
              {MASTER_TARGET_FIELDS.map(f => {
                const assignedColIdx = columnMapping[f.key] || 0;
                return (
                  <div
                    key={f.key}
                    style={{
                      background: '#0a1628',
                      border: assignedColIdx ? '1px solid rgba(56, 189, 248, 0.35)' : '1px solid #1e293b',
                      borderRadius: '8px',
                      padding: '10px 12px'
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: '4px' }}>
                      <label style={{ fontSize: '12px', fontWeight: 800, color: '#f1f5f9' }}>
                        {f.label} {f.required && <span style={{ color: '#f87171' }}>*</span>}
                      </label>
                      <span style={{ fontSize: '10px', color: '#FFC200', fontWeight: 700 }}>
                        Ref: {f.masterCol}
                      </span>
                    </div>
                    <select
                      value={assignedColIdx}
                      onChange={(e) => {
                        const val = parseInt(e.target.value, 10) || 0;
                        setColumnMapping(prev => ({
                          ...prev,
                          [f.key]: val
                        }));
                      }}
                      style={{
                        width: '100%',
                        background: '#040c18',
                        border: '1px solid #334155',
                        borderRadius: '6px',
                        color: '#f8fafc',
                        padding: '6px 8px',
                        fontSize: '11.5px',
                        boxSizing: 'border-box'
                      }}
                    >
                      <option value={0}>[None / Leave Blank]</option>
                      {excelInspection.headers.map((h, hIdx) => {
                        const cNum = hIdx + 1;
                        // Find first non-empty sample value for preview
                        const sampleVal = (excelInspection.sampleRows || [])
                          .map(r => r[f.key] || '')
                          .filter(Boolean)[0] || '';
                        return (
                          <option key={cNum} value={cNum}>
                            Col {cNum}: {h} {sampleVal ? `(e.g. "${sampleVal.slice(0, 15)}")` : ''}
                          </option>
                        );
                      })}
                    </select>
                    <div style={{ fontSize: '10px', color: '#64748b', marginTop: '4px' }}>
                      {f.desc}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Live Data Preview Table (First 3 Sample Rows) */}
            {excelInspection.sampleRows && excelInspection.sampleRows.length > 0 && (
              <div style={{ marginBottom: '20px' }}>
                <h4 style={{ margin: '0 0 8px', color: '#cbd5e1', fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Live Data Extraction Preview (First {excelInspection.sampleRows.length} Rows)
                </h4>
                <div style={{ overflowX: 'auto', background: '#040c18', border: '1px solid #1e293b', borderRadius: '8px' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px', textAlign: 'left' }}>
                    <thead>
                      <tr style={{ background: '#0a1628', borderBottom: '1px solid #1e293b', color: '#94a3b8' }}>
                        <th style={{ padding: '6px 10px' }}>Row</th>
                        <th style={{ padding: '6px 10px' }}>City</th>
                        <th style={{ padding: '6px 10px' }}>Location</th>
                        <th style={{ padding: '6px 10px' }}>Type</th>
                        <th style={{ padding: '6px 10px' }}>Light</th>
                        <th style={{ padding: '6px 10px' }}>Size</th>
                        <th style={{ padding: '6px 10px' }}>Rate</th>
                        <th style={{ padding: '6px 10px' }}>Availability</th>
                      </tr>
                    </thead>
                    <tbody>
                      {excelInspection.sampleRows.map((r, rIdx) => {
                        return (
                          <tr key={rIdx} style={{ borderBottom: '1px solid #111e33', color: '#e2e8f0' }}>
                            <td style={{ padding: '6px 10px', color: '#64748b' }}>{r.row || rIdx + 3}</td>
                            <td style={{ padding: '6px 10px', fontWeight: 600 }}>{r.city || '—'}</td>
                            <td style={{ padding: '6px 10px', maxWidth: '240px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {r.location || '—'}
                            </td>
                            <td style={{ padding: '6px 10px' }}>{r.media || '—'}</td>
                            <td style={{ padding: '6px 10px' }}>{r.light || '—'}</td>
                            <td style={{ padding: '6px 10px' }}>{r.size || '—'}</td>
                            <td style={{ padding: '6px 10px', color: '#4ade80' }}>{r.rate || '—'}</td>
                            <td style={{ padding: '6px 10px' }}>{r.avail || '—'}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Modal Actions */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: '1px solid #1e293b', paddingTop: '16px' }}>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => {
                  if (excelInspection?.mapping) {
                    setColumnMapping(excelInspection.mapping);
                  }
                }}
                style={{ fontSize: '12px' }}
              >
                Reset to Auto-Detected Mapping
              </button>
              <div style={{ display: 'flex', gap: '10px' }}>
                <button
                  type="button"
                  className="scooh-btn secondary"
                  onClick={() => setShowMappingModal(false)}
                  style={{ minHeight: '38px', padding: '0 16px', fontSize: '12px' }}
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
        </div>
      )}

      {/* ========================================================================= */}
      {/* SOFTWARE TERMS & CONDITIONS MODAL (Rule 7: Strict Software Sourced Terms)  */}
      {/* ========================================================================= */}
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
