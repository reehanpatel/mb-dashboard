import React, { useState, useEffect, useRef, useCallback } from 'react';
import api from './api';

export default function PptConverterView() {
  // Master template state
  const [templateStatus, setTemplateStatus] = useState({ loading: true, available: false, filename: '', size: 0, error: null });
  const [templateUpdating, setTemplateUpdating] = useState(false);

  // File upload & state
  const [file, setFile] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const fileInputRef = useRef(null);
  const templateInputRef = useRef(null);

  // Conversion process
  const [converting, setConverting] = useState(false);
  const [progressStep, setProgressStep] = useState(0); // 0: idle, 1: upload, 2: extract, 3: convert, 4: done
  const [progressPercent, setProgressPercent] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [conversionResult, setConversionResult] = useState(null);
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

  const validateAndSetFile = (selectedFile) => {
    setError(null);
    if (!selectedFile) return;
    const name = selectedFile.name || '';
    if (!name.toLowerCase().endsWith('.pptx')) {
      setError('Please select a valid PowerPoint (.pptx) file.');
      return;
    }
    setFile(selectedFile);
    setConversionResult(null);
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

  // Upload or replace master template
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

  // Convert workflow
  const handleConvert = async () => {
    if (!file || !templateStatus.available || converting) return;

    setError(null);
    setConverting(true);
    setProgressStep(1);
    setProgressPercent(15);
    setProgressMessage('Uploading PowerPoint presentation...');

    // Progress animation milestones
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
      setProgressMessage('Extracting location photos & billboard dimensions...');
    }, 1200);

    setTimeout(() => {
      setProgressStep(3);
      setProgressMessage('Populating Testing PPT template & generating previews...');
    }, 2800);

    try {
      const fd = new FormData();
      fd.append('file', file);

      const res = await api.post('/ppt-converter/convert', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 180000 // 3 minutes timeout for large presentations
      });

      clearInterval(pInterval);
      setProgressPercent(100);
      setProgressStep(4);
      setProgressMessage('Conversion completed successfully!');

      setTimeout(() => {
        setConversionResult(res.data);
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
      const downloadUrl = `/api/ppt-converter/download/${conversionResult.downloadId}`;
      const res = await api.get(downloadUrl, { responseType: 'blob' });
      const blob = new Blob([res.data], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = conversionResult.downloadName || 'Converted_MediaBuzz_Master_Format.pptx';
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
          <h1>PPT Converter</h1>
          <p>
            Convert PowerPoint presentations into the official Media Buzz Master Template format used in Automated PPT.
            Extracts billboard photographs, site codes, location descriptions, and dimensions while preserving the exact 16:9 widescreen brand master design.
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
                Download Media Buzz PPT
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

      {/* Master Template Status Bar */}
      <div className={`scooh-panel scooh-ppt-template-status ${templateStatus.available ? 'template-ready' : 'template-warning'}`}>
        <div className="scooh-ppt-template-info">
          <div className="scooh-ppt-template-badge">
            {templateStatus.loading ? (
              <span className="scooh-spinner-inline" />
            ) : templateStatus.available ? (
              <span className="scooh-status-icon ready">✓</span>
            ) : (
              <span className="scooh-status-icon missing">!</span>
            )}
          </div>
          <div className="scooh-ppt-template-text">
            <div className="scooh-ppt-template-title">
              <strong>Master Template:</strong> <code>{templateStatus.filename}</code>
              {templateStatus.available && (
                <span className="scooh-pill active" style={{ marginLeft: '10px' }}>
                  Available ({formatBytes(templateStatus.size)})
                </span>
              )}
              {!templateStatus.available && !templateStatus.loading && (
                <span className="scooh-tag danger" style={{ marginLeft: '10px' }}>
                  Template Unavailable
                </span>
              )}
            </div>
            <p className="scooh-ppt-template-desc">
              {templateStatus.available
                ? 'Official 16:9 widescreen Automated PPT theme with pitch black background, yellow code badge, specifications panel, and brand closing slide is active.'
                : 'The master template MediaBuzz_Master_Template.pptx is required to generate presentations.'}
            </p>
          </div>
        </div>

        <div className="scooh-ppt-template-actions">
          {templateStatus.available && (
            <button
              type="button"
              className="scooh-btn secondary"
              onClick={handleDownloadTemplate}
              title="Download master template"
              id="btn-download-master-template"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
              </svg>
              Download Template
            </button>
          )}
          <input
            type="file"
            ref={templateInputRef}
            onChange={handleTemplateUpload}
            accept=".pptx"
            style={{ display: 'none' }}
          />
          <button
            type="button"
            className="scooh-btn ghost"
            onClick={() => templateInputRef.current?.click()}
            disabled={templateUpdating}
            title="Upload or update master template"
          >
            {templateUpdating ? 'Updating...' : 'Replace Template'}
          </button>
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
          <h3>Upload Presentation</h3>
          <p className="scooh-footnote" style={{ marginBottom: '16px' }}>
            Upload any source PowerPoint presentation containing billboard or outdoor site photography.
          </p>

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
              accept=".pptx"
              style={{ display: 'none' }}
            />

            {!file ? (
              <div className="scooh-ppt-drop-prompt">
                <div className="scooh-ppt-drop-icon">
                  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                    <path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242M12 12v9m-4-4 4-4 4 4" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <h4>Drag and drop PowerPoint file here</h4>
                <p>or click to browse from your computer (.pptx format)</p>
                <div className="scooh-ppt-drop-badge">
                  <span>Supports all OOH / Billboard presentation formats</span>
                </div>
              </div>
            ) : (
              <div className="scooh-ppt-file-card">
                <div className="scooh-ppt-file-icon">
                  <svg width="36" height="36" viewBox="0 0 24 24" fill="var(--mb-primary)">
                    <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>
                  </svg>
                </div>
                <div className="scooh-ppt-file-details">
                  <strong className="scooh-ppt-file-name">{file.name}</strong>
                  <span className="scooh-ppt-file-size">{formatBytes(file.size)}</span>
                  <span className="scooh-pill active" style={{ marginLeft: '12px' }}>Ready to Convert</span>
                </div>
                <button
                  type="button"
                  className="scooh-iconbtn danger-icon"
                  onClick={(e) => {
                    e.stopPropagation();
                    setFile(null);
                    if (fileInputRef.current) fileInputRef.current.value = '';
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

          {/* Conversion Progress Indicator */}
          {converting && (
            <div className="scooh-ppt-progress-container">
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
                <div className={`scooh-ppt-step ${progressStep >= 1 ? 'active' : ''}`}>
                  <span className="scooh-ppt-step-dot" />
                  <span>1. Upload</span>
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 2 ? 'active' : ''}`}>
                  <span className="scooh-ppt-step-dot" />
                  <span>2. Extraction</span>
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 3 ? 'active' : ''}`}>
                  <span className="scooh-ppt-step-dot" />
                  <span>3. Conversion</span>
                </div>
                <div className={`scooh-ppt-step ${progressStep >= 4 ? 'active' : ''}`}>
                  <span className="scooh-ppt-step-dot" />
                  <span>4. Ready</span>
                </div>
              </div>
            </div>
          )}

          {/* Action Row */}
          <div className="scooh-ppt-action-row" style={{ marginTop: '20px', display: 'flex', gap: '12px', alignItems: 'center' }}>
            <button
              type="button"
              className={`scooh-btn primary ${converting ? 'is-loading' : ''}`}
              disabled={!file || !templateStatus.available || converting}
              onClick={handleConvert}
              id="btn-convert-ppt"
              style={{ minWidth: '180px' }}
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                <path d="M12 4V1L8 5l4 4V6c3.31 0 6 2.69 6 6 0 1.01-.25 1.97-.7 2.8l1.46 1.46A7.93 7.93 0 0 0 20 12c0-4.42-3.58-8-8-8zm0 14c-3.31 0-6-2.69-6-6 0-1.01.25-1.97.7-2.8L5.24 7.74A7.93 7.93 0 0 0 4 12c0 4.42 3.58 8 8 8v3l4-4-4-4v3z"/>
              </svg>
              Convert Presentation
            </button>

            {file && !converting && (
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={handleReset}
              >
                Clear Selection
              </button>
            )}

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
                  Official Media Buzz Master Template format generated using <code>{conversionResult.templateUsed}</code>.
                  Total presentation size: {conversionResult.totalSlidesInFile} slides.
                </p>
              </div>
            </div>

            <div className="scooh-ppt-success-actions">
              <button
                type="button"
                className="scooh-btn primary"
                onClick={handleDownload}
                id="btn-download-converted-main"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                </svg>
                Download Converted_MediaBuzz_Master_Format.pptx
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
            <div className={`scooh-kpi ${slides.some(s => s.reviewNeeded) ? 'alert' : 'good'}`}>
              <div className="n">{slides.filter(s => s.reviewNeeded).length}</div>
              <div className="l">Review Flags</div>
              {slides.some(s => s.reviewNeeded) && (
                <div className="scooh-kpi-note">Items need verification</div>
              )}
            </div>
          </div>

          {/* Slide Preview Controls */}
          <div className="scooh-toolbar" style={{ marginBottom: '16px' }}>
            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              <span style={{ fontSize: '12px', fontWeight: 700, color: '#eef2f6' }}>
                Converted Slide Previews ({slides.length})
              </span>
              {slides.some(s => s.reviewNeeded) && (
                <button
                  type="button"
                  className={`scooh-pill ${filterReviewOnly ? 'watch' : ''}`}
                  onClick={() => setFilterReviewOnly(!filterReviewOnly)}
                  style={{ cursor: 'pointer' }}
                >
                  {filterReviewOnly ? 'Showing Review Items' : `Filter Review Items (${slides.filter(s => s.reviewNeeded).length})`}
                </button>
              )}
            </div>

            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                type="button"
                className={`scooh-btn ghost ${viewMode === 'grid' ? 'active' : ''}`}
                onClick={() => setViewMode('grid')}
                style={{
                  border: viewMode === 'grid' ? '1px solid var(--mb-primary)' : '1px solid var(--mb-border)',
                  color: viewMode === 'grid' ? 'var(--mb-primary)' : '#8793a0'
                }}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M4 11h5V5H4v6zm0 7h5v-6H4v6zm6 0h5v-6h-5v6zm6 0h5v-6h-5v6zm-6-7h5V5h-5v6zm6-6v6h5V5h-5z"/>
                </svg>
                Grid View
              </button>
              <button
                type="button"
                className={`scooh-btn ghost ${viewMode === 'slideshow' ? 'active' : ''}`}
                onClick={() => setViewMode('slideshow')}
                style={{
                  border: viewMode === 'slideshow' ? '1px solid var(--mb-primary)' : '1px solid var(--mb-border)',
                  color: viewMode === 'slideshow' ? 'var(--mb-primary)' : '#8793a0'
                }}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M21 3H3c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h18c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H3V5h18v14zM10 8v8l6-4z"/>
                </svg>
                Slide View
              </button>
            </div>
          </div>

          {/* Preview Modes */}
          {viewMode === 'grid' ? (
            /* Responsive Grid of Slide Previews */
            <div className="scooh-ppt-grid-previews">
              {filteredSlides.map((slide, idx) => (
                <div
                  key={slide.siteNumber}
                  className={`scooh-panel scooh-ppt-slide-card ${slide.reviewNeeded ? 'card-review' : ''}`}
                  onClick={() => {
                    setActiveSlideIndex(slides.findIndex(s => s.siteNumber === slide.siteNumber));
                    setViewMode('slideshow');
                  }}
                >
                  <div className="scooh-ppt-card-thumb">
                    {slide.previewUrl ? (
                      <img src={slide.previewUrl} alt={`Site ${slide.siteNumber}`} />
                    ) : (
                      <div className="scooh-ppt-no-thumb">No Preview</div>
                    )}
                    <span className="scooh-ppt-slide-num">Site {slide.siteNumber}</span>
                    {slide.size && (
                      <span className="scooh-ppt-size-tag">{slide.size}</span>
                    )}
                  </div>

                  <div className="scooh-ppt-card-body">
                    <div className="scooh-ppt-card-loc" title={slide.location}>
                      {slide.location || 'Location not specified'}
                    </div>
                    {slide.city && (
                      <span className="scooh-badgechip" style={{ marginTop: '6px' }}>{slide.city}</span>
                    )}

                    {slide.reviewNeeded && (
                      <div className="scooh-ppt-review-flags">
                        {slide.reviewFlags.map((flag, fIdx) => (
                          <span key={fIdx} className="scooh-tag danger">⚠ {flag}</span>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            /* Slideshow / Inspector View */
            activeSlide && (
              <div className="scooh-panel scooh-ppt-inspector">
                <div className="scooh-ppt-inspector-header">
                  <div>
                    <span className="scooh-topbar-eyebrow">SLIDE {activeSlide.siteNumber} OF {slides.length}</span>
                    <h3 style={{ margin: '4px 0 0' }}>Site {activeSlide.siteNumber}: {activeSlide.location || 'Location'}</h3>
                  </div>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button
                      type="button"
                      className="scooh-btn ghost"
                      disabled={activeSlideIndex === 0}
                      onClick={() => setActiveSlideIndex(prev => Math.max(0, prev - 1))}
                    >
                      ← Previous Slide
                    </button>
                    <button
                      type="button"
                      className="scooh-btn ghost"
                      disabled={activeSlideIndex === slides.length - 1}
                      onClick={() => setActiveSlideIndex(prev => Math.min(slides.length - 1, prev + 1))}
                    >
                      Next Slide →
                    </button>
                  </div>
                </div>

                <div className="scooh-ppt-inspector-body">
                  {/* Big Slide Visual Presentation */}
                  <div className="scooh-ppt-stage">
                    {activeSlide.previewUrl ? (
                      <img
                        src={activeSlide.previewUrl}
                        alt={`Slide ${activeSlide.siteNumber}`}
                        className="scooh-ppt-stage-img"
                      />
                    ) : (
                      <div className="scooh-ppt-stage-placeholder">Slide Preview Unavailable</div>
                    )}
                  </div>

                  {/* Extracted Details Sidebar */}
                  <div className="scooh-ppt-meta-panel">
                    <h4>Populated Media Buzz Master Theme Data</h4>

                    <div className="scooh-statline">
                      <span>Site Identifier:</span>
                      <b>Site {activeSlide.siteNumber}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Site Code:</span>
                      <b style={{ color: 'var(--mb-primary)' }}>{activeSlide.code || `MB - ${String(activeSlide.siteNumber).padStart(2, '0')}`}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Dimensions (Size):</span>
                      <b style={{ color: '#ffffff' }}>{activeSlide.size || 'Not detected'}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>City:</span>
                      <b>{activeSlide.city || '—'}</b>
                    </div>

                    <div className="scooh-statline" style={{ alignItems: 'flex-start' }}>
                      <span>Location:</span>
                      <b style={{ textAlign: 'right', maxWidth: '240px' }}>{activeSlide.location || '—'}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Media Type:</span>
                      <b>{activeSlide.type || 'Hoarding'}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Illumination:</span>
                      <b>{activeSlide.illum || 'BL'}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Adv.fee per month:</span>
                      <b>{activeSlide.rate || 'On Request'}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Availability:</span>
                      <b>{activeSlide.availability || 'Available'}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Coordinates:</span>
                      <b>{activeSlide.coords || '23.012345, 72.567890'}</b>
                    </div>

                    <div className="scooh-statline">
                      <span>Billboard Photo:</span>
                      <b>{activeSlide.hasPhoto ? '✓ Extracted' : '⚠ Missing (Clean Card Used)'}</b>
                    </div>

                    {activeSlide.reviewNeeded && (
                      <div style={{ marginTop: '16px' }}>
                        <span className="scooh-tag danger" style={{ display: 'block', padding: '6px 10px', height: 'auto' }}>
                          <strong>Review Notice:</strong>
                          <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                            {activeSlide.reviewFlags.map((f, i) => (
                              <li key={i}>{f}</li>
                            ))}
                          </ul>
                        </span>
                      </div>
                    )}

                    <div style={{ marginTop: '24px' }}>
                      <button
                        type="button"
                        className="scooh-btn primary"
                        style={{ width: '100%' }}
                        onClick={handleDownload}
                      >
                        Download Media Buzz PPT
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )
          )}
        </div>
      )}
    </div>
  );
}
