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

export default function PptConverterView() {
  // Master template state
  const [templateStatus, setTemplateStatus] = useState({
    loading: true,
    available: false,
    filename: '',
    size: 0,
    error: null
  });
  const [templateUpdating, setTemplateUpdating] = useState(false);

  // Primary PowerPoint File
  const [pptFile, setPptFile] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const pptInputRef = useRef(null);
  const templateInputRef = useRef(null);

  // Companion Excel File & Mapping State
  const [companionExcelFile, setCompanionExcelFile] = useState(null);
  const [companionInspection, setCompanionInspection] = useState(null);
  const [companionColumnMapping, setCompanionColumnMapping] = useState({});
  const [companionSelectedSheet, setCompanionSelectedSheet] = useState('');
  const [inspectingCompanion, setInspectingCompanion] = useState(false);
  const [showCompanionMappingModal, setShowCompanionMappingModal] = useState(false);
  const companionInputRef = useRef(null);

  // Manual inputs
  const [manualCity, setManualCity] = useState('');
  const [customFileName, setCustomFileName] = useState('');

  // Conversion process
  const [converting, setConverting] = useState(false);
  const [progressStep, setProgressStep] = useState(0); // 0: idle, 1: upload, 2: extract, 3: convert, 4: done
  const [progressPercent, setProgressPercent] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [conversionResult, setConversionResult] = useState(null);
  const [error, setError] = useState(null);

  // Slide deck inspection
  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const [filterReviewOnly, setFilterReviewOnly] = useState(false);
  const [showTermsModal, setShowTermsModal] = useState(false);

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
        error: null
      });
    } catch (err) {
      setTemplateStatus({
        loading: false,
        available: false,
        filename: 'MediaBuzz_Master_Template.pptx',
        size: 0,
        error: err.response?.data?.error || err.message || 'Could not verify master template'
      });
    }
  }, []);

  useEffect(() => {
    checkTemplateStatus();
  }, [checkTemplateStatus]);

  // Handle Drag & Drop for PPT
  const handleDrag = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  }, []);

  const validateAndSetPpt = (selectedFile) => {
    setError(null);
    if (!selectedFile) return;
    const name = selectedFile.name || '';
    const lower = name.toLowerCase();
    if (!lower.endsWith('.pptx') && !lower.endsWith('.xlsx') && !lower.endsWith('.xls')) {
      setError('Please select a valid PowerPoint presentation (.pptx) or Excel workbook (.xlsx / .xls).');
      return;
    }
    if ((lower.endsWith('.xlsx') || lower.endsWith('.xls')) && pptFile) {
      setCompanionExcelFile(selectedFile);
      inspectCompanionExcel(selectedFile);
      return;
    }
    setPptFile(selectedFile);
    setConversionResult(null);
    const defaultName = name.replace(/\.[^/.]+$/, '') + '_MediaBuzz';
    if (!customFileName) setCustomFileName(defaultName);
  };

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      validateAndSetPpt(e.dataTransfer.files[0]);
    }
  }, [customFileName]);

  const handlePptFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      validateAndSetPpt(e.target.files[0]);
    }
  };

  const inspectCompanionExcel = async (selectedFile, targetSheet = null, customMap = null) => {
    try {
      setInspectingCompanion(true);
      const fd = new FormData();
      fd.append('file', selectedFile);
      const sheetToUse = targetSheet !== null ? targetSheet : companionSelectedSheet;
      if (sheetToUse) {
        fd.append('sheetName', sheetToUse);
      }
      const mapToUse = customMap !== null ? customMap : companionColumnMapping;
      if (mapToUse && Object.keys(mapToUse).length > 0) {
        fd.append('columnMapping', JSON.stringify(mapToUse));
      }
      const res = await api.post('/ppt-converter/inspect-excel', fd, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setCompanionInspection(res.data);
      if (res.data?.mapping) {
        setCompanionColumnMapping(res.data.mapping);
      }
      if (res.data?.activeSheet) {
        setCompanionSelectedSheet(res.data.activeSheet);
      }
    } catch (err) {
      console.warn('Companion Excel inspection note:', err.message);
    } finally {
      setInspectingCompanion(false);
    }
  };

  const handleCompanionSheetChange = (newSheet) => {
    setCompanionSelectedSheet(newSheet);
    if (companionExcelFile) {
      inspectCompanionExcel(companionExcelFile, newSheet, {});
    }
  };

  const handleCompanionApplyCustomMapping = () => {
    setShowCompanionMappingModal(false);
    if (companionExcelFile) {
      inspectCompanionExcel(companionExcelFile, companionSelectedSheet, companionColumnMapping);
    }
  };

  const handleCompanionExcelChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      const selected = e.target.files[0];
      const lower = (selected.name || '').toLowerCase();
      if (!lower.endsWith('.xlsx') && !lower.endsWith('.xls')) {
        alert('Companion file must be an Excel spreadsheet (.xlsx or .xls).');
        return;
      }
      setCompanionExcelFile(selected);
      inspectCompanionExcel(selected);
    }
  };

  // Convert workflow
  const handleConvert = async () => {
    if (!pptFile || !templateStatus.available || converting) return;

    setError(null);
    setConverting(true);
    setProgressStep(1);
    setProgressPercent(15);
    setProgressMessage('Uploading PowerPoint presentation...');

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
      setProgressMessage(
        companionExcelFile
          ? 'Extracting billboard photos & matching companion Excel availability & coordinates...'
          : 'Extracting billboard photos, specs & location descriptions...'
      );
    }, 1200);

    setTimeout(() => {
      setProgressStep(3);
      setProgressMessage('Populating Media Buzz Master PowerPoint & generating standardized Excel...');
    }, 2800);

    try {
      const fd = new FormData();
      fd.append('file', pptFile);
      if (companionExcelFile) {
        fd.append('excelFile', companionExcelFile);
        if (companionColumnMapping && Object.keys(companionColumnMapping).length > 0) {
          fd.append('columnMapping', JSON.stringify(companionColumnMapping));
        }
        if (companionSelectedSheet) {
          fd.append('sheetName', companionSelectedSheet);
        }
      }
      if (manualCity.trim()) {
        fd.append('manualCity', manualCity.trim());
      }
      if (customFileName.trim()) {
        fd.append('customName', customFileName.trim());
      }

      const res = await api.post('/ppt-converter/convert', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 180000
      });

      clearInterval(pInterval);
      setProgressPercent(100);
      setProgressStep(4);
      setProgressMessage('PowerPoint conversion completed successfully!');

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
      const msg = err.response?.data?.message || err.message || 'An error occurred during PowerPoint conversion.';
      setError(msg);
    }
  };

  // Download converted presentation (.pptx)
  const handleDownloadPpt = async () => {
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

  // Download extracted & standardized Excel (.xlsx)
  const handleDownloadExcel = async () => {
    if (!conversionResult?.downloadId && !conversionResult?.excelDownloadId) return;
    try {
      const cleanBase = (customFileName || 'Converted_MediaBuzz_Master_Format').trim().replace(/\.xlsx$/i, '').replace(/\.pptx$/i, '');
      const cleanName = `${cleanBase}.xlsx`;
      const downloadId = conversionResult.excelDownloadId || conversionResult.downloadId.replace(/\.pptx$/i, '.xlsx');
      const downloadUrl = `/ppt-converter/download-excel/${encodeURIComponent(downloadId)}?name=${encodeURIComponent(cleanName)}`;
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

  // Download master template
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

  // Reset conversion
  const handleReset = () => {
    setPptFile(null);
    setCompanionExcelFile(null);
    setConversionResult(null);
    setError(null);
    setConverting(false);
    setProgressStep(0);
    setProgressPercent(0);
    setActiveSlideIndex(0);
    if (pptInputRef.current) pptInputRef.current.value = '';
    if (companionInputRef.current) companionInputRef.current.value = '';
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

  return (
    <div className="scooh-ppt-converter-module">
      {/* Page Header */}
      <div className="scooh-pagehead">
        <div>
          <h1 style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <span>PPT Converter</span>
            <span className="scooh-pill active" style={{ fontSize: '11px', textTransform: 'uppercase' }}>
              Master Presentation Reference
            </span>
          </h1>
          <p>
            Convert PowerPoint presentations (<code>.pptx</code>) into the official Media Buzz Master Presentation format.
            Optionally upload an Excel sheet to extract and merge Availability, Latitude & Longitude coordinates, and rates into the presentation and exported Excel.
          </p>
        </div>
        <div className="scooh-headactions">
          {conversionResult && (
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleDownloadPpt}
                id="btn-download-ppt-conv-ppt-top"
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
                id="btn-download-ppt-conv-xls-top"
                style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.35)' }}
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                Download Extracted Excel (.xlsx)
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={handleReset}
                id="btn-reset-ppt-converter-top"
              >
                Convert Another Presentation
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Mandatory Reference Standards Status Card */}
      <div className="scooh-panel scooh-ppt-template-bar" style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '14px' }}>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
            <div className="scooh-ppt-template-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="var(--mb-primary)">
                <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>
              </svg>
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                <strong style={{ color: '#fff', fontSize: '14px' }}>
                  PowerPoint Master Reference Standards:
                </strong>
                <span className="scooh-pill active" style={{ fontSize: '11px' }}>
                  Theme: {templateStatus.filename || 'MediaBuzz_Master_Template.pptx'}
                </span>
                <span className="scooh-pill" style={{ background: 'rgba(255, 194, 0, 0.15)', color: '#FFC200', border: '1px solid rgba(255, 194, 0, 0.3)', fontSize: '11px' }}>
                  16:9 Widescreen Standard
                </span>
              </div>
              <p className="scooh-ppt-template-desc" style={{ marginTop: '4px', marginBottom: 0 }}>
                Automatically standardizes billboard presentation slides into the 16:9 widescreen Media Buzz theme with verified 7-point software Terms & Conditions.
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

        {/* Governance Notice */}
        <div style={{ background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(56, 189, 248, 0.25)', borderRadius: '8px', padding: '8px 12px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px', fontSize: '12px', color: '#cbd5e1' }}>
          <div>
            <span style={{ color: '#38bdf8', fontWeight: 700, marginRight: '6px' }}>RULE ENFORCEMENT:</span>
            <span>Master template is permanently preserved read-only. Official 7-point Media Buzz software terms will be attached to the final presentation.</span>
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
              <h3 style={{ margin: '0 0 4px', color: '#fff' }}>1. Upload PowerPoint Presentation (.pptx)</h3>
              <p className="scooh-footnote" style={{ margin: 0 }}>
                Upload any vendor or client PowerPoint presentation to convert into Media Buzz Master format.
              </p>
            </div>
            <span className="scooh-tag" style={{ background: 'rgba(56, 189, 248, 0.1)', color: '#38bdf8', border: '1px solid rgba(56, 189, 248, 0.25)' }}>
              .pptx Supported
            </span>
          </div>

          {/* Primary Drag & Drop Upload Zone for PPT */}
          <div
            className={`scooh-ppt-dropzone ${dragActive ? 'drag-over' : ''} ${pptFile ? 'has-file' : ''}`}
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
            onClick={() => !pptFile && pptInputRef.current?.click()}
          >
            <input
              type="file"
              ref={pptInputRef}
              onChange={handlePptFileChange}
              accept=".pptx,.xlsx,.xls"
              style={{ display: 'none' }}
            />

            {!pptFile ? (
              <div className="scooh-ppt-drop-prompt">
                <div className="scooh-ppt-drop-icon">
                  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                    <path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242M12 12v9m-4-4 4-4 4 4" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <h4>Drag and drop presentation or Excel workbook here</h4>
                <p>or click to browse (.pptx, .xlsx, .xls formats)</p>
              </div>
            ) : (
              <div className="scooh-ppt-file-card">
                <div className="scooh-ppt-file-icon">
                  <svg width="36" height="36" viewBox="0 0 24 24" fill="var(--mb-primary)">
                    <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>
                  </svg>
                </div>
                <div className="scooh-ppt-file-details">
                  <strong className="scooh-ppt-file-name">{pptFile.name}</strong>
                  <span className="scooh-ppt-file-size">{formatBytes(pptFile.size)}</span>
                  <span className="scooh-pill active" style={{ marginLeft: '12px' }}>PowerPoint Ready</span>
                </div>
                <button
                  type="button"
                  className="scooh-iconbtn danger-icon"
                  onClick={(e) => {
                    e.stopPropagation();
                    setPptFile(null);
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

          {/* Secondary: Companion Excel Data Upload Box */}
          <div style={{ marginTop: '20px', background: '#0a1424', border: '1px solid #1e293b', borderRadius: '12px', padding: '18px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px', flexWrap: 'wrap', gap: '8px' }}>
              <div>
                <h4 style={{ margin: '0 0 4px', color: '#fff', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '14px' }}>
                  <span style={{ color: '#22c55e' }}>📊</span>
                  <span>2. Upload Companion / Data Excel (Optional)</span>
                  <span className="scooh-pill" style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.3)', fontSize: '10px' }}>
                    Auto-Extracts Availability & Coordinates
                  </span>
                </h4>
                <p style={{ margin: 0, fontSize: '12px', color: '#94a3b8' }}>
                  If your PPT misses Availability, Latitude / Longitude coordinates, or Rates, upload your Excel spreadsheet here.
                  We will automatically extract the data from Excel and put it into your PPT slides and exported Excel!
                </p>
              </div>
              <input
                type="file"
                ref={companionInputRef}
                onChange={handleCompanionExcelChange}
                accept=".xlsx,.xls"
                style={{ display: 'none' }}
              />
              {!companionExcelFile && (
                <button
                  type="button"
                  className="scooh-btn secondary"
                  onClick={() => companionInputRef.current?.click()}
                  style={{ background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.35)', fontSize: '12px', padding: '6px 14px' }}
                >
                  + Attach Data Excel (.xlsx / .xls)
                </button>
              )}
            </div>

            {companionExcelFile ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', background: '#111e33', border: '1px solid rgba(34, 197, 94, 0.3)', borderRadius: '8px', padding: '12px 14px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="#22c55e">
                      <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                    </svg>
                    <div>
                      <strong style={{ color: '#fff', fontSize: '13px' }}>{companionExcelFile.name}</strong>
                      <div style={{ fontSize: '11px', color: '#4ade80', marginTop: '2px' }}>
                        ✓ Attached ({formatBytes(companionExcelFile.size)})
                        {inspectingCompanion && ' — Inspecting columns...'}
                        {!inspectingCompanion && companionInspection && ` — ${companionInspection.totalSites} sites detected in sheet "${companionInspection.activeSheet}"`}
                      </div>
                    </div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {companionInspection && (
                      <button
                        type="button"
                        className="scooh-btn secondary"
                        onClick={() => setShowCompanionMappingModal(true)}
                        style={{ fontSize: '11.5px', padding: '3px 10px' }}
                      >
                        ⚙️ Column Mapping (Optional)
                      </button>
                    )}
                    <button
                      type="button"
                      className="scooh-iconbtn danger-icon"
                      onClick={() => {
                        setCompanionExcelFile(null);
                        setCompanionInspection(null);
                        setCompanionColumnMapping({});
                        setCompanionSelectedSheet('');
                      }}
                      title="Remove companion Excel"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
                      </svg>
                    </button>
                  </div>
                </div>

                {companionInspection && (
                  <div style={{ borderTop: '1px solid rgba(255, 255, 255, 0.08)', paddingTop: '10px' }}>
                    {companionInspection.sheets && companionInspection.sheets.length > 1 && (
                      <div style={{ marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', fontWeight: 600 }}>Worksheet:</span>
                        <select
                          value={companionSelectedSheet || companionInspection.activeSheet || ''}
                          onChange={(e) => handleCompanionSheetChange(e.target.value)}
                          style={{ background: '#0f172a', color: '#38bdf8', border: '1px solid #0284c7', borderRadius: '4px', padding: '2px 6px', fontSize: '11px', fontWeight: 700, outline: 'none', cursor: 'pointer' }}
                        >
                          {companionInspection.sheets.map(s => (
                            <option key={s} value={s}>{s}</option>
                          ))}
                        </select>
                      </div>
                    )}
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', fontSize: '10.5px', color: '#cbd5e1' }}>
                      {MASTER_TARGET_FIELDS.map(f => {
                        const numIdx = parseInt(companionColumnMapping[f.key], 10);
                        const colName = (!isNaN(numIdx) && numIdx >= 1 && companionInspection.headers[numIdx - 1]) ? companionInspection.headers[numIdx - 1] : null;
                        return (
                          <span
                            key={f.key}
                            onClick={() => setShowCompanionMappingModal(true)}
                            style={{
                              padding: '2px 7px',
                              borderRadius: '4px',
                              background: colName ? 'rgba(34, 197, 94, 0.12)' : 'rgba(255, 255, 255, 0.04)',
                              border: colName ? '1px solid rgba(34, 197, 94, 0.3)' : '1px solid #334155',
                              color: colName ? '#e2e8f0' : '#64748b',
                              cursor: 'pointer'
                            }}
                            title={colName ? `Mapped to Column ${numIdx}: "${colName}" (Click to customize)` : 'Not mapped (Click to configure)'}
                          >
                            <strong>{f.masterCol}:</strong> {colName || 'Blank'}
                          </span>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <div style={{ padding: '8px 12px', background: 'rgba(255, 255, 255, 0.02)', border: '1px dashed #243447', borderRadius: '8px', textAlign: 'center', fontSize: '12px', color: '#64748b' }}>
                No companion Excel attached. (PPT will be converted using slide content alone, or attach an Excel above to enrich with availability and GPS coordinates).
              </div>
            )}
          </div>

          {/* Configuration Inputs: Manual City & Rename File */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px', marginTop: '20px' }}>
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
                      background: manualCity === c ? 'rgba(242, 201, 76, 0.25)' : 'rgba(255, 255, 255, 0.05)',
                      color: manualCity === c ? 'var(--mb-primary)' : '#94a3b8',
                      border: manualCity === c ? '1px solid var(--mb-primary)' : '1px solid #334155',
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
                  placeholder="e.g. MediaBuzz_Proposal_Campaign"
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
                You can customize the filename here before conversion or on the download screen.
              </p>
            </div>
          </div>

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
                  <span>1</span> Upload Presentation
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 2 ? 'active' : ''} ${progressStep > 2 ? 'done' : ''}`}>
                  <span>2</span> Parse & Match Excel
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
              disabled={!pptFile || !templateStatus.available || converting}
              id="btn-convert-ppt"
              style={{ minWidth: '260px', padding: '12px 24px', fontSize: '14px', fontWeight: 700 }}
            >
              {converting ? (
                <>
                  <span className="scooh-spin">🔄</span> Converting Presentation...
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
                  {conversionResult.processedSlides} Billboard Slides Converted into Master Format
                </h3>
                <p style={{ margin: 0, color: 'var(--mb-muted)', fontSize: '12px' }}>
                  Official 16:9 widescreen layout without unnecessary terms slide.
                  {conversionResult.companionMerged && (
                    <span style={{ color: '#4ade80', marginLeft: '6px', fontWeight: 700 }}>
                      ✓ Availability Date & GPS Coordinates successfully merged from companion Excel into slides.
                    </span>
                  )}
                  {conversionResult.manualCity && (
                    <span style={{ color: '#FFC200', marginLeft: '6px' }}>
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
                  id="btn-download-ppt-conv-ppt-main"
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
                  id="btn-download-ppt-conv-xls-main"
                  style={{ flex: 1, background: 'rgba(34, 197, 94, 0.15)', color: '#4ade80', border: '1px solid rgba(34, 197, 94, 0.35)' }}
                >
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                  </svg>
                  Download Extracted Excel
                </button>
              </div>
            </div>
          </div>

          {/* Slide Deck Grid / Cards */}
          <div className="scooh-panel scooh-ppt-gallery-panel">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '10px' }}>
              <div>
                <h3 style={{ margin: '0 0 2px', color: '#fff' }}>Converted Presentation Slides ({slides.length})</h3>
                <span className="scooh-footnote">Includes Title Covers, Billboard Sites, and Closing Slide (Clean Presentation without Terms Slide)</span>
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
                        alt={`Slide ${slide.siteNumber}`}
                        style={{ width: '100%', height: '100%', objectFit: 'cover' }}
                      />
                    ) : (
                      <div style={{ color: '#64748b', fontSize: '12px', textAlign: 'center', padding: '10px' }}>
                        <span>Site {slide.siteNumber}</span>
                        <div style={{ fontSize: '11px', marginTop: '4px' }}>Extracted from Presentation</div>
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
                    {slide.companionMatched && (
                      <span
                        style={{
                          position: 'absolute',
                          top: '8px',
                          right: '8px',
                          background: 'rgba(34, 197, 94, 0.9)',
                          color: '#fff',
                          fontWeight: 700,
                          fontSize: '10px',
                          padding: '2px 6px',
                          borderRadius: '4px'
                        }}
                      >
                        ✓ Excel Enriched
                      </span>
                    )}
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

      {/* Companion Excel Column Mapping Modal */}
      {showCompanionMappingModal && companionInspection && (
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
          onClick={() => setShowCompanionMappingModal(false)}
        >
          <div
            className="scooh-panel"
            style={{
              width: '100%',
              maxWidth: '920px',
              maxHeight: '90vh',
              overflowY: 'auto',
              border: '1px solid #22c55e',
              boxShadow: '0 25px 60px rgba(0,0,0,0.85), 0 0 25px rgba(34, 197, 94, 0.2)',
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
                  <span style={{ color: '#22c55e' }}>📊</span>
                  <span>Companion Excel Column Mapping</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  Map which columns in your companion Excel spreadsheet contain Availability, Latitude / Longitude GPS coordinates, Rates, etc.
                </p>
              </div>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => setShowCompanionMappingModal(false)}
                style={{ padding: '4px 10px', fontSize: '14px', minHeight: '30px' }}
              >
                ✕
              </button>
            </div>

            {/* Sheet Selector */}
            {companionInspection.sheets && companionInspection.sheets.length > 1 && (
              <div style={{ marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '12px', background: '#0a1628', padding: '10px 14px', borderRadius: '8px' }}>
                <span style={{ fontSize: '12.5px', fontWeight: 700, color: '#e2e8f0' }}>Target Worksheet:</span>
                <select
                  value={companionSelectedSheet || companionInspection.activeSheet || ''}
                  onChange={(e) => handleCompanionSheetChange(e.target.value)}
                  style={{ background: '#111e33', color: '#38bdf8', border: '1px solid #0284c7', borderRadius: '6px', padding: '6px 12px', fontSize: '12.5px', fontWeight: 700 }}
                >
                  {companionInspection.sheets.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              </div>
            )}

            {/* Mapping Grid */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(420px, 1fr))', gap: '12px', marginBottom: '20px' }}>
              {MASTER_TARGET_FIELDS.map((field) => {
                const currentIdx = companionColumnMapping[field.key] || '';
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
                        setCompanionColumnMapping(prev => ({
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
                      {companionInspection.headers.map((h, i) => (
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
                onClick={() => setShowCompanionMappingModal(false)}
              >
                Cancel
              </button>
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleCompanionApplyCustomMapping}
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
