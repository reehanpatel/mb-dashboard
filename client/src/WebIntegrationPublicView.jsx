import React, { useState, useEffect, useMemo, useRef } from 'react';

export default function WebIntegrationPublicView() {
  const [sites, setSites] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [search, setSearch] = useState('');
  const [cityFilter, setCityFilter] = useState('');
  const [mediaFilter, setMediaFilter] = useState('');
  const [lightingFilter, setLightingFilter] = useState('');
  const [availFilter, setAvailFilter] = useState('');
  const [sortKey, setSortKey] = useState('site_code');
  const [selectedPhotoSite, setSelectedPhotoSite] = useState(null);
  const [activePhotoIdx, setActivePhotoIdx] = useState(0);
  const [inquirySite, setInquirySite] = useState(null);
  const [liveSyncPulse, setLiveSyncPulse] = useState(false);
  const [lastUpdatedTime, setLastUpdatedTime] = useState(null);
  const [companySettings, setCompanySettings] = useState({
    company_name: 'Media Buzz',
    phone: '+91 98250 12345',
    email: 'info@mediabuzz.in'
  });

  // URL query options for customized embedding
  const urlParams = useMemo(() => new URLSearchParams(window.location.search), []);
  const hideHeader = urlParams.get('header') === '0' || urlParams.get('hide_header') === '1';
  const hideRates = urlParams.get('hide_rates') === '1';
  const initialCity = urlParams.get('city') || '';
  const initialMedia = urlParams.get('media_type') || '';

  useEffect(() => {
    if (initialCity) setCityFilter(initialCity);
    if (initialMedia) setMediaFilter(initialMedia);
  }, [initialCity, initialMedia]);

  // Fetch sites from public unauthenticated endpoint
  const fetchSites = async (isBackground = false) => {
    if (!isBackground) setLoading(true);
    try {
      const res = await fetch('/api/public/sites');
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      const rawList = Array.isArray(json?.data) ? json.data : (Array.isArray(json?.sites) ? json.sites : []);
      const normalized = rawList.map(s => ({
        ...s,
        images: Array.isArray(s.images) && s.images.length > 0 ? s.images : (Array.isArray(s.photos) ? s.photos : []),
        availability: s.availability || s.status || (s.is_available ? 'Available' : 'Booked')
      }));
      setSites(normalized);
      if (json?.company) {
        setCompanySettings(prev => ({ ...prev, ...json.company }));
      }
      setLastUpdatedTime(new Date());
      if (isBackground) {
        triggerSyncAnimation();
      }
    } catch (err) {
      console.warn('Public sites fetch fallback:', err);
      if (!isBackground) setError(err.message);
    } finally {
      if (!isBackground) setLoading(false);
    }
  };

  const triggerSyncAnimation = () => {
    setLiveSyncPulse(true);
    setTimeout(() => setLiveSyncPulse(false), 2500);
  };

  useEffect(() => {
    fetchSites();

    // 1. Server-Sent Events (SSE) for Real-Time Auto-Sync
    let sse = null;
    try {
      sse = new EventSource('/api/public/sites/stream');
      sse.onmessage = (event) => {
        try {
          const data = JSON.parse(event.data);
          if (data && data.type === 'SITES_UPDATED') {
            console.log('[WebIntegration] Received real-time site update event:', data);
            fetchSites(true);
          }
        } catch (e) {}
      };
      sse.onerror = () => {
        // SSE disconnected, will retry automatically
      };
    } catch (e) {
      console.warn('SSE not supported or blocked, relying on broadcast/poll', e);
    }

    // 2. BroadcastChannel for instant same-browser multi-tab synchronization
    let channel = null;
    try {
      if (typeof BroadcastChannel !== 'undefined') {
        channel = new BroadcastChannel('mb_sites_sync');
        channel.onmessage = (msg) => {
          if (msg && msg.data && msg.data.type === 'SITES_UPDATED') {
            console.log('[WebIntegration] BroadcastChannel site update received');
            fetchSites(true);
          }
        };
      }
    } catch (e) {}

    // 3. Fallback background poll (every 15 seconds)
    const interval = setInterval(() => {
      fetchSites(true);
    }, 15000);

    return () => {
      if (sse) sse.close();
      if (channel) channel.close();
      clearInterval(interval);
    };
  }, []);

  const cities = useMemo(() => Array.from(new Set(sites.map(s => s.city).filter(Boolean))).sort(), [sites]);
  const mediaTypes = useMemo(() => Array.from(new Set(sites.map(s => s.media_type).filter(Boolean))).sort(), [sites]);
  const lightings = useMemo(() => Array.from(new Set(sites.map(s => s.lighting).filter(Boolean))).sort(), [sites]);

  const filteredSites = useMemo(() => {
    return sites.filter(s => {
      if (search) {
        const q = search.toLowerCase().trim();
        const fullText = [
          s.site_code,
          s.city,
          s.area,
          s.address,
          s.media_type,
          s.size,
          s.lighting,
          s.notes
        ].join(' ').toLowerCase();
        if (!fullText.includes(q)) return false;
      }
      if (cityFilter && s.city !== cityFilter) return false;
      if (mediaFilter && s.media_type !== mediaFilter) return false;
      if (lightingFilter && s.lighting !== lightingFilter) return false;
      if (availFilter) {
        const isAvail = String(s.availability || '').toLowerCase().includes('avail');
        if (availFilter === 'Available' && !isAvail) return false;
        if (availFilter === 'Booked' && isAvail) return false;
      }
      return true;
    }).sort((a, b) => {
      if (sortKey === 'site_code') return String(a.site_code).localeCompare(String(b.site_code), undefined, { numeric: true });
      if (sortKey === 'city') return String(a.city || '').localeCompare(String(b.city || ''));
      if (sortKey === 'size') return String(a.size || '').localeCompare(String(b.size || ''));
      return 0;
    });
  }, [sites, search, cityFilter, mediaFilter, lightingFilter, availFilter, sortKey]);

  return (
    <div className="mb-web-integration-page">
      {/* Top Banner & Header */}
      {!hideHeader && (
        <header className="mb-web-header">
          <div className="mb-web-header-inner">
            <div className="mb-web-brand">
              <img src="/assets/media-buzz-logo.png" alt="Media Buzz" className="mb-web-logo" onError={(e) => { e.target.style.display = 'none'; }} />
              <div>
                <h1 className="mb-web-title">Prime Outdoor Advertising Inventory</h1>
                <p className="mb-web-subtitle">Explore premium billboards, hoardings, and OOH media locations</p>
              </div>
            </div>

            <div className="mb-web-status-pill-wrap">
              <div className={`mb-web-status-pill ${liveSyncPulse ? 'pulse' : ''}`} title="Connected to Media Buzz live inventory server">
                <span className="mb-status-dot"></span>
                <span>Live Real-Time Sync</span>
              </div>
              {lastUpdatedTime && (
                <span className="mb-last-sync-label">
                  Synced: {lastUpdatedTime.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' })}
                </span>
              )}
            </div>
          </div>
        </header>
      )}

      {/* Floating Live Update Toast */}
      {liveSyncPulse && (
        <div className="mb-live-toast">
          <span className="toast-icon">⚡</span>
          <span>Inventory updated in real-time with latest site changes!</span>
        </div>
      )}

      {/* Controls & Filter Bar */}
      <section className="mb-web-controls">
        <div className="mb-web-controls-row">
          {/* Search Box */}
          <div className="mb-search-input-wrap">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="11" cy="11" r="8"></circle>
              <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
            </svg>
            <input
              type="text"
              placeholder="Search by site code, location, area, city..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="mb-search-input"
            />
            {search && (
              <button type="button" className="mb-clear-btn" onClick={() => setSearch('')}>×</button>
            )}
          </div>

          {/* Filters */}
          <div className="mb-filters-group">
            <select value={cityFilter} onChange={e => setCityFilter(e.target.value)} className="mb-select">
              <option value="">All Cities ({cities.length})</option>
              {cities.map(c => <option key={c} value={c}>{c}</option>)}
            </select>

            <select value={mediaFilter} onChange={e => setMediaFilter(e.target.value)} className="mb-select">
              <option value="">All Media Types ({mediaTypes.length})</option>
              {mediaTypes.map(m => <option key={m} value={m}>{m}</option>)}
            </select>

            <select value={lightingFilter} onChange={e => setLightingFilter(e.target.value)} className="mb-select">
              <option value="">All Lighting ({lightings.length})</option>
              {lightings.map(l => <option key={l} value={l}>{l === 'FL' ? 'Front Lit (FL)' : l === 'BL' ? 'Back Lit (BL)' : l}</option>)}
            </select>

            <select value={availFilter} onChange={e => setAvailFilter(e.target.value)} className="mb-select">
              <option value="">All Status</option>
              <option value="Available">Available Only</option>
              <option value="Booked">Booked Only</option>
            </select>

            <select value={sortKey} onChange={e => setSortKey(e.target.value)} className="mb-select">
              <option value="site_code">Sort: Site Code</option>
              <option value="city">Sort: City</option>
              <option value="size">Sort: Size</option>
            </select>
          </div>
        </div>

        {/* Results summary bar */}
        <div className="mb-results-summary">
          <span>Showing <strong>{filteredSites.length}</strong> of {sites.length} outdoor advertising sites</span>
          {(search || cityFilter || mediaFilter || lightingFilter || availFilter) && (
            <button
              type="button"
              className="mb-reset-filters-btn"
              onClick={() => {
                setSearch('');
                setCityFilter('');
                setMediaFilter('');
                setLightingFilter('');
                setAvailFilter('');
              }}
            >
              Reset Filters
            </button>
          )}
        </div>
      </section>

      {/* Main Grid View */}
      {loading ? (
        <div className="mb-loading-state">
          <div className="mb-spinner"></div>
          <p>Loading live outdoor inventory...</p>
        </div>
      ) : error ? (
        <div className="mb-error-state">
          <p>Unable to load site catalog: {error}</p>
          <button type="button" onClick={() => fetchSites()} className="mb-btn mb-btn-primary">Retry</button>
        </div>
      ) : filteredSites.length === 0 ? (
        <div className="mb-empty-state">
          <div className="mb-empty-icon">📍</div>
          <h3>No matching outdoor sites found</h3>
          <p>Try clearing your search query or adjusting your filters to see more locations.</p>
        </div>
      ) : (
        <div className="mb-sites-grid">
          {filteredSites.map(s => {
            const hasImages = Array.isArray(s.images) && s.images.length > 0;
            const primaryImg = hasImages ? s.images[0] : null;
            const isAvail = String(s.availability || '').toLowerCase().includes('avail');
            const googleMapsUrl = s.maps_url || (s.latitude && s.longitude ? `https://www.google.com/maps?q=${s.latitude},${s.longitude}` : `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent([s.address, s.area, s.city].filter(Boolean).join(', '))}`);

            return (
              <article key={s.id || s.site_code} className="mb-site-card">
                {/* Photo Header */}
                <div
                  className="mb-card-media"
                  onClick={() => {
                    if (hasImages) {
                      setSelectedPhotoSite(s);
                      setActivePhotoIdx(0);
                    }
                  }}
                  style={{ cursor: hasImages ? 'pointer' : 'default' }}
                >
                  {primaryImg ? (
                    <img src={primaryImg} alt={s.site_code} loading="lazy" className="mb-card-img" />
                  ) : (
                    <div className="mb-card-no-img">
                      <span>🖼️ No Photo Available</span>
                    </div>
                  )}

                  {/* Badges on Image */}
                  <div className="mb-card-badge-top">
                    <span className="mb-badge mb-badge-code">{s.site_code}</span>
                    <span className={`mb-badge ${isAvail ? 'mb-badge-avail' : 'mb-badge-booked'}`}>
                      {s.availability || 'Available'}
                    </span>
                  </div>

                  {hasImages && (
                    <div className="mb-photo-count-badge">
                      📷 {s.images.length} {s.images.length === 1 ? 'Photo' : 'Photos'}
                    </div>
                  )}
                </div>

                {/* Card Content */}
                <div className="mb-card-body">
                  <div className="mb-card-location">
                    <h3 className="mb-card-title">{s.address || s.area || s.site_code}</h3>
                    <p className="mb-card-sublocation">
                      📍 {[s.area, s.city].filter(Boolean).join(', ')}
                    </p>
                  </div>

                  {/* Specs Grid */}
                  <div className="mb-specs-grid">
                    <div className="mb-spec-item">
                      <span className="spec-label">DIMENSIONS</span>
                      <span className="spec-val">{s.size || (s.width && s.height ? `${s.width} x ${s.height}` : '—')}</span>
                    </div>
                    <div className="mb-spec-item">
                      <span className="spec-label">MEDIA TYPE</span>
                      <span className="spec-val">{s.media_type || 'Hoarding'}</span>
                    </div>
                    <div className="mb-spec-item">
                      <span className="spec-label">LIGHTING</span>
                      <span className="spec-val">{s.lighting === 'FL' ? 'Front Lit' : s.lighting === 'BL' ? 'Back Lit' : (s.lighting || 'Non-Lit')}</span>
                    </div>
                    <div className="mb-spec-item">
                      <span className="spec-label">FACING</span>
                      <span className="spec-val">{s.facing || 'Upper'}</span>
                    </div>
                  </div>

                  {/* Card Actions */}
                  <div className="mb-card-footer">
                    <a
                      href={googleMapsUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="mb-btn mb-btn-outline"
                      title="View on Google Maps"
                    >
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <polygon points="3 11 22 2 13 21 11 13 3 11"></polygon>
                      </svg>
                      Directions
                    </a>

                    <button
                      type="button"
                      className="mb-btn mb-btn-primary"
                      onClick={() => setInquirySite(s)}
                    >
                      Inquire Now
                    </button>
                  </div>
                </div>
              </article>
            );
          })}
        </div>
      )}

      {/* Lightbox / Photo Gallery Modal */}
      {selectedPhotoSite && (
        <div className="mb-modal-overlay" onClick={() => setSelectedPhotoSite(null)}>
          <div className="mb-modal-content mb-gallery-modal" onClick={e => e.stopPropagation()}>
            <div className="mb-modal-header">
              <div>
                <h2>{selectedPhotoSite.site_code} Photo Gallery</h2>
                <p>{[selectedPhotoSite.address, selectedPhotoSite.area, selectedPhotoSite.city].filter(Boolean).join(', ')}</p>
              </div>
              <button type="button" className="mb-modal-close" onClick={() => setSelectedPhotoSite(null)}>×</button>
            </div>

            <div className="mb-modal-body mb-gallery-body">
              <div className="mb-gallery-main">
                <img
                  src={selectedPhotoSite.images[activePhotoIdx]}
                  alt={`${selectedPhotoSite.site_code} photo ${activePhotoIdx + 1}`}
                  className="mb-gallery-active-img"
                />
                {selectedPhotoSite.images.length > 1 && (
                  <>
                    <button
                      type="button"
                      className="mb-gallery-nav mb-prev"
                      onClick={() => setActivePhotoIdx((activePhotoIdx - 1 + selectedPhotoSite.images.length) % selectedPhotoSite.images.length)}
                    >
                      ‹
                    </button>
                    <button
                      type="button"
                      className="mb-gallery-nav mb-next"
                      onClick={() => setActivePhotoIdx((activePhotoIdx + 1) % selectedPhotoSite.images.length)}
                    >
                      ›
                    </button>
                  </>
                )}
              </div>

              {selectedPhotoSite.images.length > 1 && (
                <div className="mb-gallery-thumbs">
                  {selectedPhotoSite.images.map((img, idx) => (
                    <img
                      key={idx}
                      src={img}
                      alt={`Thumb ${idx + 1}`}
                      className={`mb-gallery-thumb ${idx === activePhotoIdx ? 'active' : ''}`}
                      onClick={() => setActivePhotoIdx(idx)}
                    />
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Inquiry Modal */}
      {inquirySite && (
        <div className="mb-modal-overlay" onClick={() => setInquirySite(null)}>
          <div className="mb-modal-content mb-inquiry-modal" onClick={e => e.stopPropagation()}>
            <div className="mb-modal-header">
              <div>
                <h2>Inquire About {inquirySite.site_code}</h2>
                <p>{[inquirySite.address, inquirySite.area, inquirySite.city].filter(Boolean).join(', ')}</p>
              </div>
              <button type="button" className="mb-modal-close" onClick={() => setInquirySite(null)}>×</button>
            </div>

            <div className="mb-modal-body">
              <div className="mb-inquiry-details">
                <div className="mb-inquiry-spec">
                  <strong>Site Code:</strong> {inquirySite.site_code}
                </div>
                <div className="mb-inquiry-spec">
                  <strong>Dimensions:</strong> {inquirySite.size || 'Standard'}
                </div>
                <div className="mb-inquiry-spec">
                  <strong>Status:</strong> {inquirySite.availability || 'Available'}
                </div>
              </div>

              <div className="mb-inquiry-actions">
                <a
                  href={`https://wa.me/919825012345?text=${encodeURIComponent(`Hi Media Buzz, I am interested in booking site ${inquirySite.site_code} (${inquirySite.size}) located at ${inquirySite.address || inquirySite.area || inquirySite.city}. Please share pricing and availability dates.`)}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="mb-btn mb-btn-whatsapp"
                >
                  <span>💬 Chat on WhatsApp</span>
                </a>

                <a
                  href={`mailto:info@mediabuzz.in?subject=${encodeURIComponent(`Booking Inquiry: Site ${inquirySite.site_code}`)}&body=${encodeURIComponent(`Hello Media Buzz team,\n\nI am interested in advertising on Site Code: ${inquirySite.site_code}.\nLocation: ${[inquirySite.address, inquirySite.area, inquirySite.city].filter(Boolean).join(', ')}\nDimensions: ${inquirySite.size}\n\nPlease get in touch with rate cards and available booking periods.\n\nThank you!`)}`}
                  className="mb-btn mb-btn-email"
                >
                  <span>✉️ Email Inquiry</span>
                </a>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Embedded footer */}
      {!hideHeader && (
        <footer className="mb-web-footer">
          <p>© {new Date().getFullYear()} Media Buzz OOH • Live Integrated Inventory Portal</p>
        </footer>
      )}
    </div>
  );
}
