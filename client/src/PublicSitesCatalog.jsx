import React, { useState, useEffect, useMemo } from 'react';
import axios from 'axios';
import { defaultSites } from './defaultSites.js';
import './PublicSitesCatalog.css';

export default function PublicSitesCatalog({ embed = false }) {
  // Check if embed mode from props or query string (?embed=true) or iframe
  const isEmbed = useMemo(() => {
    if (embed) return true;
    try {
      const params = new URLSearchParams(window.location.search);
      if (params.get('embed') === 'true' || params.get('embed') === '1') return true;
      if (window.self !== window.top) return true;
    } catch (_) {}
    return false;
  }, [embed]);

  const [sites, setSites] = useState([]);
  const [loading, setLoading] = useState(true);
  const [company, setCompany] = useState({
    name: 'Media Buzz',
    phone: '+91 98250 00000',
    whatsapp: '919825000000',
    email: 'info@mediabuzz.com'
  });
  const [stats, setStats] = useState({ total: 0, available: 0, occupied: 0 });

  // Filters
  const [search, setSearch] = useState('');
  const [availFilter, setAvailFilter] = useState('all'); // all | available | occupied
  const [cityFilter, setCityFilter] = useState('all');
  const [typeFilter, setTypeFilter] = useState('all');
  const [lightingFilter, setLightingFilter] = useState('all');
  const [targetDate, setTargetDate] = useState('');
  const [viewMode, setViewMode] = useState('grid'); // grid | list

  // Modals
  const [lightboxImg, setLightboxImg] = useState(null);
  const [inquireSite, setInquireSite] = useState(null);
  const [showEmbedModal, setShowEmbedModal] = useState(false);
  const [embedTab, setEmbedTab] = useState('iframe');
  const [copied, setCopied] = useState(false);

  // Fetch sites and runtime availability
  const fetchSites = async (dateParam = '') => {
    setLoading(true);
    try {
      const params = new URLSearchParams(window.location.search);
      const keyParam = params.get('key') || params.get('api_key') || params.get('apiKey') || '';

      const qParams = new URLSearchParams();
      if (dateParam) qParams.set('date', dateParam);
      if (keyParam) qParams.set('api_key', keyParam);
      const qs = qParams.toString() ? `?${qParams.toString()}` : '';

      const headers = {};
      if (keyParam) {
        headers['x-api-key'] = keyParam;
      } else {
        const token = localStorage.getItem('token') || localStorage.getItem('auth_token');
        if (token) headers['Authorization'] = `Bearer ${token}`;
      }

      const res = await axios.get(`/api/public/sites${qs}`, { headers, timeout: 8000 });
      if (res.data && res.data.sites && res.data.sites.length > 0) {
        setSites(res.data.sites);
        if (res.data.company) setCompany(res.data.company);
        if (res.data.stats) setStats(res.data.stats);
        setLoading(false);
        return;
      }
    } catch (err) {
      console.warn('API error, falling back to local dataset:', err.message);
    }

    // Fallback if API is unavailable or offline
    const fallback = defaultSites.map((s, idx) => {
      let flags = {};
      try { flags = typeof s.flags === 'string' ? JSON.parse(s.flags) : (s.flags || {}); } catch (_) {}
      const isAvail = String(s.availability || 'Available').toLowerCase().startsWith('avail');
      return {
        id: s.id || idx + 1,
        site_code: s.site_code,
        city: s.city || 'Ahmedabad',
        area: s.area || '',
        address: s.address || '',
        media_type: s.media_type || 'Billboard',
        lighting: s.lighting || 'FL',
        facing: s.facing || '',
        size: s.size || `${s.width || ''}x${s.height || ''}`,
        width: s.width,
        height: s.height,
        orientation: s.orientation || '',
        monthly_rate: s.monthly_rate || 0,
        latitude: s.latitude,
        longitude: s.longitude,
        maps_url: s.maps_url || (s.latitude && s.longitude ? `https://www.google.com/maps?q=${s.latitude},${s.longitude}` : ''),
        photos: Array.isArray(flags.ppt_images) ? flags.ppt_images : [],
        is_available: isAvail,
        status: isAvail ? 'Available' : 'Occupied',
        available_from: null
      };
    });

    setSites(fallback);
    const availCount = fallback.filter(s => s.is_available).length;
    setStats({ total: fallback.length, available: availCount, occupied: fallback.length - availCount });
    setLoading(false);
  };

  useEffect(() => {
    fetchSites(targetDate);
  }, [targetDate]);

  // Unique options for filters
  const cities = useMemo(() => {
    const set = new Set();
    sites.forEach(s => { if (s.city) set.add(s.city); });
    return Array.from(set).sort();
  }, [sites]);

  const mediaTypes = useMemo(() => {
    const set = new Set();
    sites.forEach(s => { if (s.media_type) set.add(s.media_type); });
    return Array.from(set).sort();
  }, [sites]);

  const lightingTypes = useMemo(() => {
    const set = new Set();
    sites.forEach(s => { if (s.lighting) set.add(s.lighting); });
    return Array.from(set).sort();
  }, [sites]);

  // Filtered sites
  const filteredSites = useMemo(() => {
    return sites.filter(s => {
      // Search term
      if (search) {
        const q = search.toLowerCase();
        const code = String(s.site_code || '').toLowerCase();
        const loc = String(s.address || '').toLowerCase();
        const area = String(s.area || '').toLowerCase();
        const city = String(s.city || '').toLowerCase();
        if (!code.includes(q) && !loc.includes(q) && !area.includes(q) && !city.includes(q)) {
          return false;
        }
      }

      // Availability filter
      if (availFilter === 'available' && !s.is_available) return false;
      if (availFilter === 'occupied' && s.is_available) return false;

      // City filter
      if (cityFilter !== 'all' && s.city !== cityFilter) return false;

      // Media type filter
      if (typeFilter !== 'all' && s.media_type !== typeFilter) return false;

      // Lighting filter
      if (lightingFilter !== 'all' && s.lighting !== lightingFilter) return false;

      return true;
    });
  }, [sites, search, availFilter, cityFilter, typeFilter, lightingFilter]);

  // WhatsApp click-to-chat builder
  const getWhatsAppUrl = (site) => {
    const cleanPhone = String(company.whatsapp || '919825000000').replace(/[^0-9]/g, '');
    const msg = `Hello ${company.name || 'Media Buzz'},\n\nI am inquiring about Site *${site.site_code}*:\n📍 Location: ${site.address || site.area || 'Prime Location'} (${site.city})\n📐 Size: ${site.size || '-'}\n💡 Type: ${site.media_type} (${site.lighting || 'FL'})\n📅 Status: ${site.status}${site.available_from ? ' (Available from ' + site.available_from + ')' : ''}\n\nPlease share availability details, rate card and proposal pitch deck.\n\nThank you!`;
    return `https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`;
  };

  const copyToClipboard = (text) => {
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2200);
  };

  const currentHost = typeof window !== 'undefined' ? window.location.origin : 'https://mb.webtrionix.com';

  return (
    <div className={`psc-container ${isEmbed ? 'psc-embed-mode' : ''}`}>
      {/* Top Header */}
      <header className="psc-header">
        <div className="psc-header-left">
          <h1>
            <span>{company.name || 'Media Buzz'}</span> Outdoor Sites Portfolio
            <span className="psc-live-pulse">
              <span className="psc-pulse-dot" /> Live Runtime
            </span>
          </h1>
          <p>Real-time hoarding availability, specifications & verified geo-locations</p>
        </div>

        <div className="psc-header-stats">
          <div className="psc-stat-pill">
            <span>Total Sites:</span>
            <span className="psc-stat-num">{stats.total}</span>
          </div>
          <div className="psc-stat-pill available">
            <span>🟢 Available:</span>
            <span className="psc-stat-num">{stats.available}</span>
          </div>
          <div className="psc-stat-pill occupied">
            <span>🔴 Booked:</span>
            <span className="psc-stat-num">{stats.occupied}</span>
          </div>
          <button 
            type="button" 
            className="psc-btn-embed"
            onClick={() => setShowEmbedModal(true)}
            title="Get WordPress integration code"
          >
            <span>WordPress Embed</span>
          </button>
        </div>
      </header>

      {/* Toolbar & Filters */}
      <section className="psc-toolbar">
        <div className="psc-search-row">
          <div className="psc-search-box">
            <span className="psc-search-icon">🔍</span>
            <input
              type="text"
              className="psc-search-input"
              placeholder="Search by Site Code (e.g. MB-01), Landmark, Area or City..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            {search && (
              <button 
                type="button" 
                className="psc-clear-btn" 
                onClick={() => setSearch('')}
                aria-label="Clear search"
              >
                ✕
              </button>
            )}
          </div>

          {/* Quick Availability Chips */}
          <div className="psc-filter-chips">
            <button
              type="button"
              className={`psc-chip ${availFilter === 'all' ? 'active' : ''}`}
              onClick={() => setAvailFilter('all')}
            >
              All ({sites.length})
            </button>
            <button
              type="button"
              className={`psc-chip ${availFilter === 'available' ? 'active' : ''}`}
              onClick={() => setAvailFilter('available')}
            >
              🟢 Available Now ({stats.available})
            </button>
            <button
              type="button"
              className={`psc-chip ${availFilter === 'occupied' ? 'active' : ''}`}
              onClick={() => setAvailFilter('occupied')}
            >
              🔴 Booked ({stats.occupied})
            </button>
          </div>
        </div>

        {/* Secondary Filters */}
        <div className="psc-filter-dropdowns">
          {cities.length > 1 && (
            <select
              className="psc-select"
              value={cityFilter}
              onChange={(e) => setCityFilter(e.target.value)}
              aria-label="Filter by City"
            >
              <option value="all">🏙️ All Cities</option>
              {cities.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          )}

          {mediaTypes.length > 1 && (
            <select
              className="psc-select"
              value={typeFilter}
              onChange={(e) => setTypeFilter(e.target.value)}
              aria-label="Filter by Media Type"
            >
              <option value="all">📐 All Media Types</option>
              {mediaTypes.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          )}

          {lightingTypes.length > 1 && (
            <select
              className="psc-select"
              value={lightingFilter}
              onChange={(e) => setLightingFilter(e.target.value)}
              aria-label="Filter by Lighting"
            >
              <option value="all">💡 All Lighting</option>
              {lightingTypes.map(l => (
                <option key={l} value={l}>
                  {l === 'BL' ? 'Back Lit (BL)' : l === 'FL' ? 'Front Lit (FL)' : l}
                </option>
              ))}
            </select>
          )}

          {/* Target Date Picker */}
          <div className="psc-date-box">
            <span>📅 Check date:</span>
            <input
              type="date"
              className="psc-date-input"
              value={targetDate}
              onChange={(e) => setTargetDate(e.target.value)}
              title="Check site availability for a future campaign date"
            />
            {targetDate && (
              <button 
                type="button" 
                className="psc-chip" 
                onClick={() => setTargetDate('')}
                style={{ padding: '4px 8px', fontSize: '11px' }}
              >
                Reset
              </button>
            )}
          </div>

          {/* Grid vs List View */}
          <div className="psc-view-toggles">
            <button
              type="button"
              className={`psc-view-btn ${viewMode === 'grid' ? 'active' : ''}`}
              onClick={() => setViewMode('grid')}
              title="Cards Grid View"
            >
              Cards
            </button>
            <button
              type="button"
              className={`psc-view-btn ${viewMode === 'list' ? 'active' : ''}`}
              onClick={() => setViewMode('list')}
              title="List View"
            >
              List
            </button>
          </div>
        </div>
      </section>

      {/* Main Content Area */}
      {loading ? (
        <div className="psc-empty">
          <div className="psc-empty-icon">⏳</div>
          <h3>Loading real-time site portfolio...</h3>
          <p>Syncing availability with active campaigns</p>
        </div>
      ) : filteredSites.length === 0 ? (
        <div className="psc-empty">
          <div className="psc-empty-icon">🔍</div>
          <h3>No matching sites found</h3>
          <p>Try clearing your filters or search keywords</p>
          <button 
            type="button" 
            className="psc-btn-embed"
            style={{ marginTop: '12px' }}
            onClick={() => {
              setSearch('');
              setAvailFilter('all');
              setCityFilter('all');
              setTypeFilter('all');
              setLightingFilter('all');
              setTargetDate('');
            }}
          >
            Reset All Filters
          </button>
        </div>
      ) : viewMode === 'grid' ? (
        <div className="psc-grid">
          {filteredSites.map((site) => {
            const hasPhoto = Array.isArray(site.photos) && site.photos.length > 0;
            const photoUrl = hasPhoto ? site.photos[0] : null;

            return (
              <article key={site.site_code || site.id} className="psc-card">
                {/* Visual Header */}
                <div 
                  className="psc-card-visual"
                  onClick={() => hasPhoto && setLightboxImg(photoUrl)}
                  title={hasPhoto ? 'Click to view photo' : site.site_code}
                >
                  {hasPhoto ? (
                    <img 
                      src={photoUrl} 
                      alt={`Site ${site.site_code} - ${site.address || site.area}`} 
                      className="psc-card-img" 
                      loading="lazy"
                    />
                  ) : (
                    <div className="psc-card-placeholder">
                      <span className="psc-placeholder-icon">🏙️</span>
                      <span className="psc-placeholder-text">{site.media_type || 'Billboard'} • {site.size || 'Prime Site'}</span>
                    </div>
                  )}

                  {/* Top Badges */}
                  <span className="psc-badge-type">{site.media_type || 'OOH Media'}</span>
                  
                  <span className={`psc-badge-status ${site.status.toLowerCase()}`}>
                    {site.is_available ? '● Available' : '● Booked'}
                  </span>

                  {hasPhoto && site.photos.length > 1 && (
                    <span className="psc-photo-count">📷 {site.photos.length}</span>
                  )}
                </div>

                {/* Card Body */}
                <div className="psc-card-body">
                  <div className="psc-card-title-row">
                    <span className="psc-site-code">{site.site_code}</span>
                    {site.city && <span className="psc-city-tag">{site.city}</span>}
                  </div>

                  <p className="psc-location-text" title={site.address || site.area}>
                    {site.address || site.area || 'Prime Outdoor Location'}
                  </p>

                  {/* Specs Grid */}
                  <div className="psc-specs-grid">
                    <div className="psc-spec-item">
                      <span className="psc-spec-label">Dimensions</span>
                      <span className="psc-spec-val">{site.size ? `${site.size} ft` : 'Custom'}</span>
                    </div>
                    <div className="psc-spec-item">
                      <span className="psc-spec-label">Illumination</span>
                      <span className="psc-spec-val">
                        {site.lighting === 'BL' ? 'Back Lit' : site.lighting === 'FL' ? 'Front Lit' : site.lighting || 'Standard'}
                      </span>
                    </div>
                    <div className="psc-spec-item">
                      <span className="psc-spec-label">Facing</span>
                      <span className="psc-spec-val">{site.facing || 'Prime Traffic'}</span>
                    </div>
                    <div className="psc-spec-item">
                      <span className="psc-spec-label">Area</span>
                      <span className="psc-spec-val">{site.area || site.city || 'Central'}</span>
                    </div>
                  </div>

                  {/* Runtime Availability Banner */}
                  <div className={`psc-avail-banner ${site.status.toLowerCase()}`}>
                    {site.is_available ? (
                      <span>🟢 <strong>Available Immediately</strong> for booking</span>
                    ) : (
                      <span>
                        🔴 <strong>Occupied</strong>
                        {site.available_from ? ` • Available from ${new Date(site.available_from).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' })}` : ' (Currently Running Campaign)'}
                      </span>
                    )}
                  </div>

                  {/* Action Buttons */}
                  <div className="psc-card-actions">
                    <a
                      href={getWhatsAppUrl(site)}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="psc-btn-inquire"
                    >
                      <span>💬 Inquire on WhatsApp</span>
                    </a>

                    {site.maps_url ? (
                      <a
                        href={site.maps_url}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="psc-btn-map"
                        title="View Location on Google Maps"
                      >
                        📍 Map
                      </a>
                    ) : (
                      <button
                        type="button"
                        className="psc-btn-map"
                        onClick={() => setInquireSite(site)}
                        title="View Site Specifications"
                      >
                        ℹ️ Info
                      </button>
                    )}
                  </div>
                </div>
              </article>
            );
          })}
        </div>
      ) : (
        /* Table / List View */
        <div className="psc-list-table-wrap">
          <table className="psc-table">
            <thead>
              <tr>
                <th>Site Code</th>
                <th>Location / Landmark</th>
                <th>City</th>
                <th>Type</th>
                <th>Size (W x H)</th>
                <th>Lighting</th>
                <th>Status</th>
                <th>Availability</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {filteredSites.map(site => (
                <tr key={site.site_code || site.id}>
                  <td><strong>{site.site_code}</strong></td>
                  <td>{site.address || site.area}</td>
                  <td>{site.city}</td>
                  <td>{site.media_type}</td>
                  <td>{site.size}</td>
                  <td>{site.lighting === 'BL' ? 'Back Lit' : site.lighting === 'FL' ? 'Front Lit' : site.lighting}</td>
                  <td>
                    <span className={`psc-badge-status ${site.status.toLowerCase()}`} style={{ position: 'static', display: 'inline-flex' }}>
                      {site.status}
                    </span>
                  </td>
                  <td>
                    {site.is_available ? (
                      <span style={{ color: '#34d399', fontWeight: 600 }}>Available Now</span>
                    ) : site.available_from ? (
                      <span style={{ color: '#f87171' }}>Free from {site.available_from}</span>
                    ) : (
                      <span style={{ color: '#f87171' }}>Booked</span>
                    )}
                  </td>
                  <td>
                    <a
                      href={getWhatsAppUrl(site)}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="psc-btn-inquire"
                      style={{ padding: '6px 12px', fontSize: '12px' }}
                    >
                      WhatsApp
                    </a>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* WordPress Embed Modal */}
      {showEmbedModal && (
        <div className="psc-modal-overlay" onClick={() => setShowEmbedModal(false)}>
          <div className="psc-modal" onClick={e => e.stopPropagation()}>
            <div className="psc-modal-header">
              <h2>🌐 Embed Live Sites on WordPress</h2>
              <button 
                type="button" 
                className="psc-modal-close" 
                onClick={() => setShowEmbedModal(false)}
              >
                ✕
              </button>
            </div>

            <div className="psc-modal-body">
              <p style={{ margin: '0 0 16px 0', fontSize: '14px', color: '#94a3b8' }}>
                Integrate this real-time cards gallery into any WordPress page, Elementor widget, Gutenberg block, or PHP template.
              </p>

              <div className="psc-tab-nav">
                <button
                  type="button"
                  className={`psc-tab-btn ${embedTab === 'iframe' ? 'active' : ''}`}
                  onClick={() => setEmbedTab('iframe')}
                >
                  1. Iframe Code (Easiest)
                </button>
                <button
                  type="button"
                  className={`psc-tab-btn ${embedTab === 'shortcode' ? 'active' : ''}`}
                  onClick={() => setEmbedTab('shortcode')}
                >
                  2. WordPress Shortcode
                </button>
                <button
                  type="button"
                  className={`psc-tab-btn ${embedTab === 'link' ? 'active' : ''}`}
                  onClick={() => setEmbedTab('link')}
                >
                  3. Direct URL
                </button>
                <button
                  type="button"
                  className={`psc-tab-btn ${embedTab === 'api' ? 'active' : ''}`}
                  onClick={() => setEmbedTab('api')}
                >
                  4. JSON API
                </button>
              </div>

              {embedTab === 'iframe' && (
                <div>
                  <p style={{ fontSize: '13px', margin: '0 0 8px 0', color: '#cbd5e1' }}>
                    <strong>Step 1:</strong> In WordPress, edit your page (e.g. <em>Our Sites</em> or <em>Inventory</em>).<br />
                    <strong>Step 2:</strong> Add a <strong>Custom HTML</strong> block (in Gutenberg) or <strong>HTML Widget</strong> (in Elementor).<br />
                    <strong>Step 3:</strong> Paste the embed code below:
                  </p>
                  <pre className="psc-code-block">{`<iframe src="${currentHost}/sites-catalog?embed=true" width="100%" height="950" frameborder="0" style="border:none; width:100%; min-height:950px; border-radius:12px; overflow:hidden;" loading="lazy" allowfullscreen></iframe>`}</pre>
                  <button
                    type="button"
                    className="psc-btn-copy"
                    onClick={() => copyToClipboard(`<iframe src="${currentHost}/sites-catalog?embed=true" width="100%" height="950" frameborder="0" style="border:none; width:100%; min-height:950px; border-radius:12px; overflow:hidden;" loading="lazy" allowfullscreen></iframe>`)}
                  >
                    {copied ? '✓ Copied to Clipboard!' : '📋 Copy Iframe Code'}
                  </button>
                </div>
              )}

              {embedTab === 'shortcode' && (
                <div>
                  <p style={{ fontSize: '13px', margin: '0 0 8px 0', color: '#cbd5e1' }}>
                    Add this snippet to your WordPress theme's <code>functions.php</code> file. Once added, you can simply type <code>[mediabuzz_sites]</code> anywhere on your website!
                  </p>
                  <pre className="psc-code-block">{`// Media Buzz Sites Catalog Shortcode
function mb_display_live_sites_catalog() {
    return '<iframe src="${currentHost}/sites-catalog?embed=true" width="100%" height="950" frameborder="0" style="border:none; width:100%; min-height:950px; border-radius:12px;" loading="lazy"></iframe>';
}
add_shortcode('mediabuzz_sites', 'mb_display_live_sites_catalog');`}</pre>
                  <button
                    type="button"
                    className="psc-btn-copy"
                    onClick={() => copyToClipboard(`function mb_display_live_sites_catalog() {\n    return '<iframe src="${currentHost}/sites-catalog?embed=true" width="100%" height="950" frameborder="0" style="border:none; width:100%; min-height:950px; border-radius:12px;" loading="lazy"></iframe>';\n}\nadd_shortcode('mediabuzz_sites', 'mb_display_live_sites_catalog');`)}
                  >
                    {copied ? '✓ Copied to Clipboard!' : '📋 Copy Shortcode Snippet'}
                  </button>
                </div>
              )}

              {embedTab === 'link' && (
                <div>
                  <p style={{ fontSize: '13px', margin: '0 0 8px 0', color: '#cbd5e1' }}>
                    Share this direct link with clients or link it from your WordPress menu:
                  </p>
                  <pre className="psc-code-block">{`${currentHost}/sites-catalog`}</pre>
                  <button
                    type="button"
                    className="psc-btn-copy"
                    onClick={() => copyToClipboard(`${currentHost}/sites-catalog`)}
                  >
                    {copied ? '✓ Copied to Clipboard!' : '📋 Copy Direct URL'}
                  </button>
                </div>
              )}

              {embedTab === 'api' && (
                <div>
                  <p style={{ fontSize: '13px', margin: '0 0 8px 0', color: '#cbd5e1' }}>
                    Fetch live runtime availability directly in custom WordPress PHP or JavaScript (all photo URLs are full & absolute):
                  </p>
                  <pre className="psc-code-block">{`${currentHost}/api/public/sites\n${currentHost}/api/public/sites?date=2026-11-01`}</pre>
                  <div style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
                    <button
                      type="button"
                      className="psc-btn-copy"
                      onClick={() => copyToClipboard(`${currentHost}/api/public/sites`)}
                    >
                      {copied ? '✓ Copied to Clipboard!' : '📋 Copy Full JSON URL'}
                    </button>
                    <button
                      type="button"
                      className="psc-btn-copy"
                      style={{ background: '#334155' }}
                      onClick={() => window.open(`${currentHost}/api/public/sites`, '_blank')}
                    >
                      Open JSON ↗
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Lightbox Photo Preview */}
      {lightboxImg && (
        <div className="psc-lightbox-modal" onClick={() => setLightboxImg(null)}>
          <button 
            type="button" 
            className="psc-lightbox-close" 
            onClick={() => setLightboxImg(null)}
          >
            ✕ Close
          </button>
          <img src={lightboxImg} alt="Site preview" className="psc-lightbox-img" />
        </div>
      )}

      {/* Site Info Modal */}
      {inquireSite && (
        <div className="psc-modal-overlay" onClick={() => setInquireSite(null)}>
          <div className="psc-modal" onClick={e => e.stopPropagation()}>
            <div className="psc-modal-header">
              <h2>Site Details: {inquireSite.site_code}</h2>
              <button 
                type="button" 
                className="psc-modal-close" 
                onClick={() => setInquireSite(null)}
              >
                ✕
              </button>
            </div>
            <div className="psc-modal-body">
              <p><strong>Location:</strong> {inquireSite.address || inquireSite.area}</p>
              <p><strong>City:</strong> {inquireSite.city}</p>
              <p><strong>Dimensions:</strong> {inquireSite.size} ft</p>
              <p><strong>Illumination:</strong> {inquireSite.lighting}</p>
              <p><strong>Status:</strong> {inquireSite.status} {inquireSite.available_from ? `(Free from ${inquireSite.available_from})` : ''}</p>
              <div style={{ marginTop: '20px' }}>
                <a
                  href={getWhatsAppUrl(inquireSite)}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="psc-btn-inquire"
                  style={{ display: 'inline-flex' }}
                >
                  💬 Inquire on WhatsApp
                </a>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
