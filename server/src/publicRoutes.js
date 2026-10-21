import { q } from './db.js';
import { canonicalSiteCode, computeAllOccupied } from './siteHierarchy.js';

export function registerPublicRoutes(app) {
  // Public Sites Catalog & Real-time Availability API for WordPress / Web embedding
  app.get('/api/public/sites', async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    try {
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

        const dbAvail = String(s.availability || '').toLowerCase();
        const isOccupied = Boolean(conflictReason) || dbAvail.startsWith('occupied');
        const isMaintenance = dbAvail === 'maintenance';

        let availableFrom = null;
        if (isOccupied && currentCamp?.end_date) {
          try {
            const d = new Date(currentCamp.end_date);
            d.setDate(d.getDate() + 1);
            availableFrom = d.toISOString().slice(0, 10);
          } catch (_) {}
        }

        let status = 'Available';
        if (isMaintenance) status = 'Maintenance';
        else if (isOccupied) status = 'Occupied';

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
          is_available: status === 'Available',
          status,
          available_from: availableFrom,
          occupied_reason: conflictReason && conflictReason !== 'Direct' ? `Linked panel (${conflictReason})` : null,
          next_booking_start: upcomingCamp?.start_date || null
        };
      });

      const availableCount = publicSites.filter(s => s.status === 'Available').length;
      const occupiedCount = publicSites.filter(s => s.status === 'Occupied').length;
      const maintenanceCount = publicSites.filter(s => s.status === 'Maintenance').length;

      const fullJsonUrl = `${baseUrl}/api/public/sites`;
      const fullStreamUrl = `${baseUrl}/api/public/sites/stream`;
      const fullEmbedUrl = `${baseUrl}/embed/sites`;
      const fullCatalogUrl = `${baseUrl}/sites-catalog`;

      res.json({
        success: true,
        timestamp: new Date().toISOString(),
        target_date: targetDateStr,
        base_url: baseUrl,
        full_json_url: fullJsonUrl,
        full_stream_url: fullStreamUrl,
        full_embed_url: fullEmbedUrl,
        full_catalog_url: fullCatalogUrl,
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

  // Public quick stats endpoint for embedding a mini live counter in WordPress headers/widgets
  app.get('/api/public/stats', async (req, res) => {
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
