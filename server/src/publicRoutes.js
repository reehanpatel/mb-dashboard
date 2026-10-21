import crypto from 'crypto';
import jwt from 'jsonwebtoken';
import { q } from './db.js';
import { canonicalSiteCode, computeAllOccupied } from './siteHierarchy.js';

let cachedApiKey = null;
let lastKeyFetchTime = 0;

/**
 * Retrieve or automatically generate the unique integration API key
 */
export async function getOrGenerateIntegrationApiKey(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && cachedApiKey && (now - lastKeyFetchTime < 60000)) {
    return cachedApiKey;
  }
  try {
    const rows = await q("SELECT setting_value FROM settings WHERE setting_key = 'integration_api_key' LIMIT 1");
    if (rows && rows.length > 0 && rows[0].setting_value && String(rows[0].setting_value).trim()) {
      cachedApiKey = String(rows[0].setting_value).trim();
      lastKeyFetchTime = now;
      return cachedApiKey;
    }
    // Generate new secure key (32 bytes hex prefixed with mb_live_)
    const generated = 'mb_live_' + crypto.randomBytes(16).toString('hex');
    await q(`
      INSERT INTO settings (setting_key, setting_value, updated_at)
      VALUES ('integration_api_key', ?, NOW())
      ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value), updated_at = NOW()
    `, [generated]);
    cachedApiKey = generated;
    lastKeyFetchTime = now;
    return cachedApiKey;
  } catch (err) {
    console.error('[IntegrationAuth] Failed to load/generate API key:', err);
    return cachedApiKey || 'mb_live_mediabuzz_default_key';
  }
}

/**
 * Regenerate a new API key and persist to settings table
 */
export async function regenerateIntegrationApiKey() {
  const generated = 'mb_live_' + crypto.randomBytes(16).toString('hex');
  await q(`
    INSERT INTO settings (setting_key, setting_value, updated_at)
    VALUES ('integration_api_key', ?, NOW())
    ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value), updated_at = NOW()
  `, [generated]);
  cachedApiKey = generated;
  lastKeyFetchTime = Date.now();
  return cachedApiKey;
}

/**
 * Middleware to require a valid API key for integration endpoints.
 * Supports:
 *  - Header: x-api-key: <key>
 *  - Query param: ?api_key=<key> or ?key=<key>
 *  - Header: Authorization: Bearer <key>
 *  - Authenticated dashboard JWT token (for seamless previewing)
 */
export async function requireIntegrationApiKey(req, res, next) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, x-api-key, X-Api-Key');

  if (req.method === 'OPTIONS') {
    return res.sendStatus(204);
  }

  const headerKey = req.headers['x-api-key'] || req.headers['X-Api-Key'];
  const authHeader = req.headers['authorization'] || '';
  const bearerKey = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';
  const queryKey = req.query.api_key || req.query.key || req.query.apiKey;

  const candidateKey = String(headerKey || queryKey || (bearerKey.startsWith('mb_live_') ? bearerKey : '') || '').trim();
  const activeKey = await getOrGenerateIntegrationApiKey();

  if (candidateKey && candidateKey === activeKey) {
    req.isApiKeyAuthenticated = true;
    return next();
  }

  // Also allow authenticated dashboard users (admin, manager, staff)
  if (bearerKey && !bearerKey.startsWith('mb_live_')) {
    try {
      const decoded = jwt.verify(bearerKey, process.env.JWT_SECRET);
      if (decoded && (decoded.id || decoded.email)) {
        req.user = decoded;
        return next();
      }
    } catch (_) {}
  }

  return res.status(401).json({
    success: false,
    error: 'Unauthorized',
    code: 'API_KEY_REQUIRED',
    message: 'Access restricted: A valid Media Buzz integration API key is required to access the sites inventory. Provide your key via the x-api-key header or ?api_key= query parameter.'
  });
}

export function registerPublicRoutes(app, options = {}) {
  const { auth, admin, managerOrAdmin, notifySitesUpdated } = options;

  // Key management endpoints (authenticated dashboard users)
  if (auth) {
    const roleManager = managerOrAdmin || ((req, res, next) => next());
    const roleAdmin = admin || ((req, res, next) => next());

    app.get('/api/integration/key', auth, roleManager, async (req, res) => {
      try {
        const apiKey = await getOrGenerateIntegrationApiKey();
        res.json({ success: true, apiKey });
      } catch (err) {
        res.status(500).json({ success: false, message: 'Failed to retrieve API key: ' + err.message });
      }
    });

    app.post('/api/integration/key/regenerate', auth, roleAdmin, async (req, res) => {
      try {
        const newKey = await regenerateIntegrationApiKey();
        if (typeof notifySitesUpdated === 'function') {
          notifySitesUpdated({ action: 'api_key_regenerated', timestamp: Date.now() });
        }
        res.json({ success: true, apiKey: newKey, message: 'New integration API key generated successfully' });
      } catch (err) {
        res.status(500).json({ success: false, message: 'Failed to regenerate API key: ' + err.message });
      }
    });
  }

  // Key-protected Sites Catalog & Real-time Availability API for Web Integration
  app.get('/api/public/sites', requireIntegrationApiKey, async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, x-api-key, X-Api-Key');
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    try {
      const activeKey = await getOrGenerateIntegrationApiKey();

      // 1. Fetch active sites
      const sites = await q(`
        SELECT id, site_code, city, area, address, media_type, lighting, facing,
               size, width, height, orientation, ownership, availability,
               monthly_rate, latitude, longitude, maps_url, gps, flags
        FROM sites
        WHERE record_status='active'
        ORDER BY id ASC
      `);

      // Target date for availability check (defaults to today)
      const targetDateStr = (typeof req.query.date === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(req.query.date))
        ? req.query.date
        : new Date().toISOString().slice(0, 10);

      // 2. Fetch active and upcoming campaigns covering target date or later
      let campaigns = [];
      try {
        campaigns = await q(`
          SELECT site_code, client, brand, display, start_date, end_date
          FROM campaigns
          WHERE record_status='active'
            AND site_code != ''
            AND (end_date IS NULL OR end_date >= ?)
            AND LOWER(TRIM(COALESCE(client, ''))) NOT IN ('blank', 'vacant', 'unassigned', '', '-')
            AND LOWER(TRIM(COALESCE(client, ''))) NOT LIKE 'blank%'
            AND LOWER(TRIM(COALESCE(client, ''))) NOT LIKE 'vacant%'
          ORDER BY start_date ASC
        `, [targetDateStr]);
      } catch (cErr) {
        console.warn('[PublicSites] Campaign query notice:', cErr.message);
      }

      // Track direct bookings on target date and upcoming bookings
      const directBookedOnDate = [];
      const campaignMapByCode = {};

      for (const c of campaigns) {
        const raw = String(c.site_code || '').trim();
        const sDate = c.start_date ? new Date(c.start_date).toISOString().slice(0, 10) : null;
        const eDate = c.end_date ? new Date(c.end_date).toISOString().slice(0, 10) : null;

        const matches = raw.match(/MB[-_ ]?\d+/gi);
        const codes = matches ? matches.map(m => canonicalSiteCode(m)) : [canonicalSiteCode(raw)];

        for (const code of codes) {
          if (!code) continue;
          if (!campaignMapByCode[code]) campaignMapByCode[code] = { current: null, upcoming: null };

          const isRunningOnTargetDate = (!sDate || sDate <= targetDateStr) && (!eDate || eDate >= targetDateStr);
          if (isRunningOnTargetDate) {
            if (!directBookedOnDate.includes(code)) directBookedOnDate.push(code);
            if (!campaignMapByCode[code].current || (eDate && (!campaignMapByCode[code].current.end_date || eDate > campaignMapByCode[code].current.end_date))) {
              campaignMapByCode[code].current = { ...c, start_date: sDate, end_date: eDate };
            }
          } else if (sDate && sDate > targetDateStr) {
            if (!campaignMapByCode[code].upcoming || sDate < campaignMapByCode[code].upcoming.start_date) {
              campaignMapByCode[code].upcoming = { ...c, start_date: sDate, end_date: eDate };
            }
          }
        }
      }

      // Physical panel conflict resolution
      const occupiedMap = computeAllOccupied(directBookedOnDate);

      // Fetch branding / contact settings
      let settings = {};
      try {
        const settingsRows = await q("SELECT setting_key, setting_value FROM settings WHERE setting_key IN ('company_name','currency','contact_phone','contact_whatsapp','contact_email')");
        settings = Object.fromEntries(settingsRows.map(r => [r.setting_key, r.setting_value]));
      } catch (_) {}

      // Base URL resolution for full absolute URLs in external integrations
      const protocol = (req.headers['x-forwarded-proto'] || req.protocol || 'http').split(',')[0].trim();
      const host = (req.headers['x-forwarded-host'] || req.get('host') || 'localhost:3000').split(',')[0].trim();
      const baseUrl = `${protocol}://${host}`;

      const makeFullUrl = (url) => {
        if (!url || typeof url !== 'string') return '';
        const trimmed = url.trim();
        if (trimmed.startsWith('http://') || trimmed.startsWith('https://') || trimmed.startsWith('data:')) {
          return trimmed;
        }
        const cleanPath = trimmed.startsWith('/') ? trimmed : `/${trimmed}`;
        return `${baseUrl}${cleanPath}`;
      };

      const publicSites = sites.map(s => {
        const code = canonicalSiteCode(s.site_code);
        const conflictReason = occupiedMap.get(code);

        let flags = {};
        try { flags = typeof s.flags === 'string' ? JSON.parse(s.flags) : (s.flags || {}); } catch (_) {}
        if (Array.isArray(flags)) flags = { ppt_images: flags };

        const rawImages = Array.isArray(flags.ppt_images) ? flags.ppt_images : [];
        const fullPhotos = rawImages.map(makeFullUrl).filter(Boolean);
        const primaryPhoto = fullPhotos[0] || null;

        const currentCamp = campaignMapByCode[code]?.current;
        const upcomingCamp = campaignMapByCode[code]?.upcoming;

        const formatFmtDate = (dStr) => {
          if (!dStr) return '';
          try {
            const dt = new Date(dStr);
            if (isNaN(dt.getTime())) return dStr;
            return dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
          } catch (_) {
            return dStr;
          }
        };

        const rawAvail = String(flags.ppt_availability || s.availability || '').trim();
        const rawAvailLower = rawAvail.toLowerCase();

        const isMaintenance = rawAvailLower === 'maintenance';
        const hasCampaignBooking = Boolean(currentCamp);
        const hasConflictBooking = Boolean(conflictReason);
        const hasManualOccupied = !isMaintenance && (rawAvailLower.startsWith('occupied') || rawAvailLower.startsWith('booked'));

        const isOccupied = !isMaintenance && (hasCampaignBooking || hasConflictBooking || hasManualOccupied);

        let availableFrom = null;
        let bookedTill = null;
        let displayAvailability = 'Available';

        if (isMaintenance) {
          displayAvailability = 'Maintenance';
        } else if (hasCampaignBooking && currentCamp.end_date) {
          bookedTill = currentCamp.end_date;
          try {
            const d = new Date(currentCamp.end_date);
            d.setDate(d.getDate() + 1);
            availableFrom = d.toISOString().slice(0, 10);
          } catch (_) {}
          displayAvailability = `Booked till ${formatFmtDate(currentCamp.end_date)}`;
        } else if (hasConflictBooking) {
          if (currentCamp?.end_date) {
            bookedTill = currentCamp.end_date;
            displayAvailability = `Booked till ${formatFmtDate(currentCamp.end_date)}`;
          } else {
            displayAvailability = conflictReason !== 'Direct' ? `Occupied (${conflictReason})` : 'Occupied';
          }
        } else if (hasManualOccupied) {
          displayAvailability = rawAvail;
        } else if (rawAvail && rawAvailLower !== 'available') {
          displayAvailability = rawAvail;
        } else {
          displayAvailability = 'Available';
        }

        const status = isMaintenance ? 'Maintenance' : (isOccupied ? 'Occupied' : 'Available');
        const isAvailable = status === 'Available';

        return {
          id: s.id,
          site_code: s.site_code,
          city: s.city || '',
          area: s.area || '',
          address: s.address || '',
          media_type: s.media_type || 'Billboard',
          lighting: s.lighting || 'FL',
          facing: s.facing || '',
          size: s.size || (s.width && s.height ? `${s.width}x${s.height}` : ''),
          width: s.width,
          height: s.height,
          orientation: s.orientation || '',
          monthly_rate: s.monthly_rate || 0,
          latitude: s.latitude,
          longitude: s.longitude,
          maps_url: s.maps_url || (s.latitude && s.longitude ? `https://www.google.com/maps?q=${s.latitude},${s.longitude}` : ''),
          photos: fullPhotos,
          primary_photo: primaryPhoto,
          photo_url: primaryPhoto,
          raw_photos: rawImages,
          availability: displayAvailability,
          status,
          is_available: isAvailable,
          available_from: availableFrom,
          booked_till: bookedTill,
          active_client: currentCamp?.client || null,
          active_brand: currentCamp?.brand || null,
          occupied_reason: conflictReason && conflictReason !== 'Direct' ? `Linked panel (${conflictReason})` : null,
          next_booking_start: upcomingCamp?.start_date || null
        };
      });

      const availableCount = publicSites.filter(s => s.status === 'Available').length;
      const occupiedCount = publicSites.filter(s => s.status === 'Occupied').length;
      const maintenanceCount = publicSites.filter(s => s.status === 'Maintenance').length;

      const fullJsonUrl = `${baseUrl}/api/public/sites?api_key=${activeKey}`;
      const fullStreamUrl = `${baseUrl}/api/public/sites/stream?api_key=${activeKey}`;
      const fullEmbedUrl = `${baseUrl}/embed/sites?key=${activeKey}`;
      const fullCatalogUrl = `${baseUrl}/sites-catalog?key=${activeKey}`;

      res.json({
        success: true,
        authenticated: true,
        timestamp: new Date().toISOString(),
        target_date: targetDateStr,
        base_url: baseUrl,
        full_json_url: fullJsonUrl,
        full_stream_url: fullStreamUrl,
        full_embed_url: fullEmbedUrl,
        full_catalog_url: fullCatalogUrl,
        api_key: activeKey,
        company: {
          name: settings.company_name || 'Media Buzz',
          phone: settings.contact_phone || '+91 98250 00000',
          whatsapp: settings.contact_whatsapp || '919825000000',
          email: settings.contact_email || 'info@mediabuzz.com'
        },
        stats: {
          total: publicSites.length,
          available: availableCount,
          occupied: occupiedCount,
          maintenance: maintenanceCount
        },
        sites: publicSites,
        data: publicSites
      });
    } catch (err) {
      console.error('Error fetching public sites:', err);
      res.status(500).json({ success: false, message: 'Failed to load public sites: ' + err.message });
    }
  });

  // Protected quick stats endpoint for embedding a mini live counter
  app.get('/api/public/stats', requireIntegrationApiKey, async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cache-Control', 'no-cache');
    try {
      const sites = await q("SELECT site_code, availability FROM sites WHERE record_status='active'");
      const available = sites.filter(s => String(s.availability || '').toLowerCase().startsWith('avail')).length;
      res.json({
        total: sites.length,
        available,
        occupied: sites.length - available
      });
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  });
}

