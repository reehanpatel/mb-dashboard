import 'dotenv/config';
import express from 'express';import cors from 'cors';import helmet from 'helmet';import path from 'path';import fs from 'fs';import multer from 'multer';import bcrypt from 'bcryptjs';import {fileURLToPath} from 'url';import * as xlsxModule from 'xlsx';
const XLSX = xlsxModule.default || xlsxModule;
try {
  if (XLSX.set_fs) XLSX.set_fs(fs);
} catch (e) {}
import {pool,q} from './db.js';import {auth,admin,sign,managerOrAdmin,notViewer,requireRole} from './auth.js';
import {computeAllOccupied, canonicalSiteCode, getConflictSummary, syncLinkedCampaigns, deleteLinkedCampaigns, batchDeleteLinkedCampaigns, syncAllLinkedCampaigns} from './siteHierarchy.js';
process.on('uncaughtException', err => {
  console.error('[Fatal] Uncaught Exception:', err.message);
});
process.on('unhandledRejection', (reason, promise) => {
  console.error('[Fatal] Unhandled Rejection at:', promise, 'reason:', reason);
});

const __dirname=path.dirname(fileURLToPath(import.meta.url));
const app=express();
const rawPort = process.env.PORT || 3000;
const isNumericPort = /^\d+$/.test(String(rawPort).trim());
const PORT = isNumericPort ? parseInt(rawPort, 10) : rawPort;
const uploadDir=path.resolve(__dirname,'..',process.env.UPLOAD_DIR||'uploads');
fs.mkdirSync(uploadDir,{recursive:true});
app.use(helmet({
  crossOriginResourcePolicy: { policy: "cross-origin" },
  contentSecurityPolicy: false
}));
app.use(cors({origin:true,credentials:true}));
app.use(express.json({limit:'10mb'}));
app.use((req, res, next) => {
  res.setTimeout(25000, () => {
    if (!res.headersSent) {
      console.warn(`[Timeout] Request timed out: ${req.method} ${req.originalUrl || req.url}`);
      res.status(504).json({ message: 'Request timeout' });
    }
  });
  next();
});
const serveUploads = express.static(uploadDir, {
  setHeaders: (res, filePath) => {
    res.setHeader('Cross-Origin-Resource-Policy', 'cross-origin');
    res.setHeader('Access-Control-Allow-Origin', '*');
    const ext = path.extname(filePath).toLowerCase();
    if (!ext || ext === '') {
      res.setHeader('Content-Type', 'image/jpeg');
    }
  }
});

app.use(['/uploads', '/api/uploads'], (req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cross-Origin-Resource-Policy', 'cross-origin');
  next();
}, serveUploads);

// Explicit route for uploads with smart extension fallback & 404 guard (never return index.html for assets)
app.get(['/uploads/:filename', '/api/uploads/:filename'], (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cross-Origin-Resource-Policy', 'cross-origin');
  const rawFilename = path.basename(req.params.filename || '');
  if (!rawFilename) return res.status(404).type('text/plain').send('File not found');

  const exact = path.join(uploadDir, rawFilename);
  if (fs.existsSync(exact) && fs.statSync(exact).isFile()) {
    const ext = path.extname(exact).toLowerCase();
    if (!ext) res.setHeader('Content-Type', 'image/jpeg');
    return res.sendFile(exact);
  }

  const baseNoExt = rawFilename.replace(/\.[^.]+$/, '');
  const extensions = ['.jpg', '.jpeg', '.png', '.webp', '.avif', '.gif', '.svg'];
  for (const ext of extensions) {
    const cand = path.join(uploadDir, baseNoExt + ext);
    if (fs.existsSync(cand) && fs.statSync(cand).isFile()) {
      return res.sendFile(cand);
    }
  }

  try {
    const files = fs.readdirSync(uploadDir);
    const matched = files.find(f => f.toLowerCase().startsWith(baseNoExt.toLowerCase()));
    if (matched) {
      const matchPath = path.join(uploadDir, matched);
      if (fs.statSync(matchPath).isFile()) {
        const ext = path.extname(matchPath).toLowerCase();
        if (!ext) res.setHeader('Content-Type', 'image/jpeg');
        return res.sendFile(matchPath);
      }
    }
  } catch {}

  return res.status(404).type('text/plain').send('File not found');
});

const uploadStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    let ext = path.extname(file.originalname || '').toLowerCase();
    if (!ext || ext === '.') {
      const mime = (file.mimetype || '').toLowerCase();
      if (mime.includes('png')) ext = '.png';
      else if (mime.includes('webp')) ext = '.webp';
      else if (mime.includes('gif')) ext = '.gif';
      else if (mime.includes('svg')) ext = '.svg';
      else if (mime.includes('sheet') || mime.includes('excel')) ext = '.xlsx';
      else if (mime.includes('json')) ext = '.json';
      else ext = '.jpg';
    }
    const safeBase = path.basename(file.originalname || 'photo', ext)
      .replace(/[^a-zA-Z0-9_-]/g, '_')
      .slice(0, 40) || 'upload';
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, `${safeBase}-${uniqueSuffix}${ext}`);
  }
});
const upload = multer({
  storage: uploadStorage,
  limits: { fileSize: Number(process.env.MAX_UPLOAD_MB || 50) * 1024 * 1024 }
});
const entities={
 sites:{table:'sites',order:'LENGTH(site_code) ASC, site_code ASC, id ASC'},clients:{table:'clients',order:'id DESC'},campaigns:{table:'campaigns',order:'LENGTH(site_code) ASC, site_code ASC, id DESC'},validations:{table:'validations',order:'id DESC'},electricity:{table:'electricity',order:'LENGTH(site_code) ASC, site_code ASC, id DESC'},vendors:{table:'vendors',order:'id DESC'},'vendor-jobs':{table:'vendor_jobs',order:'id DESC'},proposals:{table:'proposals',order:'id DESC'},invoices:{table:'invoices',order:'id DESC'},occupancy:{table:'occupancy_records',order:'LENGTH(site_code) ASC, site_code ASC, id DESC'}
};
const blocked=new Set(['id','created_at','updated_at']);
function safeEntity(req,res){const e=entities[req.params.entity];if(!e){res.status(404).json({message:'Unknown module'});return null}return e}
async function tableColumns(table){const rows=await q(`SHOW COLUMNS FROM \`${table}\``);return new Set(rows.map(r=>r.Field))}
async function cleanData(table,body){const cols=await tableColumns(table),out={};for(const [k,v] of Object.entries(body||{})){if(cols.has(k)&&!blocked.has(k)){if(typeof v==='object'&&v!==null)out[k]=JSON.stringify(v);else out[k]=v===''?null:v}}return out}
async function audit(user,action,type,id,before=null,after=null){try{await q('INSERT INTO activity_log (user_id,user_name,action,object_type,object_id,old_value,new_value,created_at) VALUES (?,?,?,?,?,?,?,NOW())',[user?.id||null,user?.name||'',action,type,id||null,before?JSON.stringify(before):null,after?JSON.stringify(after):null])}catch{}}

export async function ensureCampaignColumns() {
  try {
    const existingCols = await tableColumns('campaigns');
    const colsToAdd = [
      { name: 'status', type: 'VARCHAR(40) NOT NULL DEFAULT "active"' },
      { name: 'month', type: 'VARCHAR(60) NOT NULL DEFAULT ""' },
      { name: 'display', type: 'VARCHAR(190) NOT NULL DEFAULT ""' },
      { name: 'vendor_name', type: 'VARCHAR(190) NOT NULL DEFAULT ""' },
      { name: 'location', type: 'VARCHAR(255) NOT NULL DEFAULT ""' },
      { name: 'width', type: 'DECIMAL(10,2) NULL' },
      { name: 'height', type: 'DECIMAL(10,2) NULL' },
      { name: 'size', type: 'VARCHAR(60) NOT NULL DEFAULT ""' },
      { name: 'type', type: 'VARCHAR(80) NOT NULL DEFAULT "Hoarding"' },
      { name: 'days', type: 'INT NOT NULL DEFAULT 30' },
      { name: 'advt_fees', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'printing_mounting_cost', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'total_amount', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'po', type: 'VARCHAR(100) NOT NULL DEFAULT ""' },
      { name: 'bill', type: 'VARCHAR(100) NOT NULL DEFAULT ""' },
      { name: 'pending', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'occupancy', type: 'VARCHAR(100) NOT NULL DEFAULT ""' }
    ];
    for (const c of colsToAdd) {
      if (!existingCols.has(c.name)) {
        try {
          await q(`ALTER TABLE campaigns ADD COLUMN \`${c.name}\` ${c.type}`);
          console.log(`[Schema Migration] Added column \`${c.name}\` to campaigns table.`);
        } catch (err) {
          console.warn(`[Schema Migration] Notice adding \`${c.name}\`:`, err.message);
        }
      }
    }

    // Ensure booking_code is NOT a unique constraint (different clients or bookings can repeat codes)
    try {
      const indexes = await q('SHOW INDEX FROM campaigns');
      const uniqueBookingCodeIdx = indexes.find(idx => idx.Key_name === 'booking_code' && (idx.Non_unique === 0 || idx.Non_unique === '0'));
      if (uniqueBookingCodeIdx) {
        await q('ALTER TABLE campaigns DROP INDEX booking_code');
        await q('ALTER TABLE campaigns ADD KEY booking_code (booking_code)');
        console.log('[Schema Migration] Relaxed campaigns.booking_code from UNIQUE constraint to non-unique KEY.');
      }
    } catch (idxErr) {
      console.warn('[Schema Migration] booking_code index relaxation notice:', idxErr.message);
    }
  } catch (err) {
    console.warn('[Schema Migration] ensureCampaignColumns notice:', err.message);
  }
}

export async function ensureElectricityColumns() {
  try {
    const existingCols = await tableColumns('electricity');
    const colsToAdd = [
      { name: 'ecs', type: 'VARCHAR(100) NOT NULL DEFAULT ""' },
      { name: 'payment_date', type: 'DATE NULL DEFAULT NULL' }
    ];
    for (const c of colsToAdd) {
      if (!existingCols.has(c.name)) {
        try {
          await q(`ALTER TABLE electricity ADD COLUMN \`${c.name}\` ${c.type}`);
          console.log(`[Schema Migration] Added column \`${c.name}\` to electricity table.`);
        } catch (err) {
          console.warn(`[Schema Migration] Notice adding electricity.\`${c.name}\`:`, err.message);
        }
      }
    }
    // Sync payment_date with paid_date if missing
    try {
      await q("UPDATE electricity SET payment_date = paid_date WHERE (payment_date IS NULL OR payment_date = '') AND paid_date IS NOT NULL");
    } catch (_) {}
  } catch (err) {
    console.warn('[Schema Migration] ensureElectricityColumns notice:', err.message);
  }
}

/**
 * syncSiteAvailability() – Recalculates and persists availability for ALL sites
 * based on currently active campaigns. Uses the physical panel conflict map
 * so booking a combined site automatically marks all overlapping split sites
 * as Occupied, and vice-versa.
 */
async function syncSiteAvailability() {
  try {
    // 1. Get all active campaign site codes running today
    const activeCampaigns = await q(
      `SELECT site_code FROM campaigns 
       WHERE record_status='active' 
         AND site_code != '' 
         AND (start_date IS NULL OR start_date <= CURDATE())
         AND (end_date IS NULL OR end_date >= CURDATE())
         AND LOWER(TRIM(COALESCE(client, ''))) NOT IN ('blank', 'vacant', 'unassigned', '', '-')
         AND LOWER(TRIM(COALESCE(client, ''))) NOT LIKE 'blank%'
         AND LOWER(TRIM(COALESCE(client, ''))) NOT LIKE 'vacant%'`
    );
    const directBooked = [];
    for (const c of activeCampaigns) {
      const raw = String(c.site_code || '').trim();
      const matches = raw.match(/MB[-_ ]?\d+/gi);
      if (matches) {
        for (const m of matches) directBooked.push(canonicalSiteCode(m));
      } else if (raw) {
        directBooked.push(canonicalSiteCode(raw));
      }
    }

    // 2. Compute all occupied site codes (direct + linked via panel conflicts)
    const occupiedMap = computeAllOccupied(directBooked); // Map<siteCode, reason>

    // 3. Get all active sites
    const allSites = await q(`SELECT id, site_code, availability FROM sites WHERE record_status='active'`);

    // 4. Update each site's availability based on computed map
    for (const site of allSites) {
      const code = canonicalSiteCode(site.site_code);
      const occupiedEntry = occupiedMap.get(code);
      const currentAvail = String(site.availability || '').toLowerCase().trim();

      if (occupiedEntry) {
        // Site should be Occupied
        const reason = occupiedEntry === 'Direct' ? 'Occupied' : `Occupied (${occupiedEntry})`;
        if (currentAvail !== reason.toLowerCase() && currentAvail !== 'maintenance') {
          await q(`UPDATE sites SET availability=?, updated_at=NOW() WHERE id=?`, [reason, site.id]);
        }
      } else {
        // Site should be Available (revert only if previously marked occupied)
        if (currentAvail.startsWith('occupied')) {
          await q(`UPDATE sites SET availability='Available', updated_at=NOW() WHERE id=?`, [site.id]);
        }
      }
    }
  } catch (err) {
    console.error('[syncSiteAvailability] Error:', err.message);
  }
}
app.get('/api/health', async (req, res) => {
  try {
    await q('SELECT 1');
    res.json({ ok: true, version: '2026.09.09-no-nullif', timestamp: new Date().toISOString() });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});
app.post('/api/auth/login', async (req, res) => {
  try {
    const {email, password} = req.body || {};
    const rows = await q('SELECT * FROM users WHERE email=? AND status="active" LIMIT 1', [email]);
    const u = rows[0];
    if (!u || !await bcrypt.compare(password || '', u.password_hash)) {
      return res.status(401).json({message: 'Invalid email or password'});
    }
    res.json({token: sign(u), user: {id: u.id, email: u.email, name: u.name, role: u.role}});
  } catch (err) {
    console.error('Login error:', err.message);
    res.status(500).json({message: 'Database error: ' + err.message});
  }
});
app.get('/api/auth/me',auth,(req,res)=>res.json(req.user));
app.get('/api/dashboard', auth, async (req, res) => {
  try {
    const [
      [s],
      [cLive],
      [cAll],
      [e],
      [i],
      [p],
      campaignsRows,
      sitesRows,
      electricityRows
    ] = await Promise.all([
      // Sites count and status
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN LOWER(availability) = 'available' THEN 1 ELSE 0 END), 0) as available,
          COALESCE(SUM(CASE WHEN LOWER(availability) IN ('occupied', 'booked', 'active') THEN 1 ELSE 0 END), 0) as occupied
         FROM sites WHERE record_status = 'active'`),
      // Live campaigns running right now
      q(`SELECT 
          COUNT(*) as live, 
          COALESCE(SUM(revenue), 0) as revenue, 
          COALESCE(SUM(revenue - COALESCE(vendor_cost,0) - COALESCE(printing_cost,0) - COALESCE(mounting_cost,0) - COALESCE(electricity_cost,0) - COALESCE(other_cost,0)), 0) as margin 
         FROM campaigns 
         WHERE record_status = 'active' 
           AND (
             (start_date IS NOT NULL AND end_date IS NOT NULL AND CURDATE() BETWEEN start_date AND end_date)
             OR (mounting_status = 'Mounted' AND (end_date IS NULL OR end_date >= CURDATE()))
           )`),
      // Action triggers from campaigns
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN mounting_status NOT IN ('Mounted', 'Completed') AND start_date <= CURDATE() THEN 1 ELSE 0 END), 0) as mounting_overdue, 
          COALESCE(SUM(CASE WHEN validation_15_date IS NOT NULL AND validation_15_date <= CURDATE() THEN 1 ELSE 0 END), 0) as val15, 
          COALESCE(SUM(CASE WHEN final_validation_date IS NOT NULL AND final_validation_date <= CURDATE() THEN 1 ELSE 0 END), 0) as valFinal 
         FROM campaigns WHERE record_status = 'active'`),
      // Electricity bills
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN payment_status != 'Paid' AND due_date < CURDATE() THEN 1 ELSE 0 END), 0) as overdue, 
          COALESCE(SUM(CASE WHEN payment_status != 'Paid' THEN 1 ELSE 0 END), 0) as pending_bills,
          COALESCE(SUM(CASE WHEN payment_status != 'Paid' THEN amount ELSE 0 END), 0) as unpaid 
         FROM electricity WHERE record_status = 'active'`),
      // Invoices
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN invoice_status IN ('Pending', 'Draft') OR hard_copy_status = 'Pending' OR payment_status = 'Pending' THEN 1 ELSE 0 END), 0) as pending 
         FROM invoices WHERE record_status = 'active'`),
      // Proposals
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN status IN ('Draft', 'Sent', 'Open', 'Pending') THEN 1 ELSE 0 END), 0) as open_proposals 
         FROM proposals`),
      // Active campaigns rows for live alerts
      q(`SELECT id, booking_code, site_code, client, campaign_name, start_date, end_date, mounting_status, validation_15_date, final_validation_date, invoice_status, hard_copy_status 
         FROM campaigns WHERE record_status = 'active' ORDER BY id DESC`),
      // Sites rows for classification and flags
      q(`SELECT id, site_code, area, city, availability, monthly_rate, flags FROM sites WHERE record_status = 'active'`),
      // Electricity rows for live overdue alerts
      q(`SELECT id, site_code, bill_type, amount, due_date, payment_status FROM electricity WHERE record_status = 'active' AND payment_status != 'Paid'`)
    ]);

    const totalSites = Number(s?.total || 0);

    // Active sites: sites that either have availability='occupied' OR are currently linked to a live active campaign
    const liveCampaignSiteCodes = new Set(
      (campaignsRows || [])
        .filter(c => {
          const now = new Date();
          const start = c.start_date ? new Date(c.start_date) : null;
          const end = c.end_date ? new Date(c.end_date) : null;
          return (!start || start <= now) && (!end || end >= now);
        })
        .map(c => c.site_code)
        .filter(Boolean)
    );

    let activeSites = Number(s?.occupied || 0);
    if (liveCampaignSiteCodes.size > activeSites) {
      activeSites = Math.min(totalSites, liveCampaignSiteCodes.size);
    }
    const nonActiveSites = Math.max(0, totalSites - activeSites);
    const activeCampaigns = Number(cLive?.live || 0);

    // Count flagged and premium sites
    let flaggedCount = 0;
    let primeCount = 0;
    for (const st of (sitesRows || [])) {
      try {
        const fl = typeof st.flags === 'string' ? JSON.parse(st.flags) : (st.flags || {});
        if (Array.isArray(fl) ? fl.length > 0 : (Array.isArray(fl.tags) && fl.tags.length > 0)) {
          flaggedCount++;
        }
      } catch {}
      if (Number(st.monthly_rate || 0) >= 100000) {
        primeCount++;
      }
    }

    // Real portfolio occupancy breakdown
    const occupancyBuckets = {
      'Occupied': activeSites,
      'Available': nonActiveSites,
      'Prime Sites': primeCount,
      'Needs review': flaggedCount
    };

    const avgOccupancy = totalSites > 0 ? Math.round((activeSites / totalSites) * 100) : 0;

    // Real alerts from actual active database records only
    const alerts = [];
    const now = new Date();
    for (const cmp of (campaignsRows || [])) {
      const start = cmp.start_date ? new Date(cmp.start_date) : null;
      const end = cmp.end_date ? new Date(cmp.end_date) : null;
      const endIn7 = end && (end - now) / 86400000 <= 7 && (end - now) >= 0;

      if (cmp.mounting_status !== 'Mounted' && start && start <= now) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Mounting overdue',
          class: 'danger',
          link: '/campaigns'
        });
      }
      if (endIn7) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Campaign ending soon',
          class: 'watch',
          link: '/campaigns'
        });
      }
      if (cmp.validation_15_date && new Date(cmp.validation_15_date) <= now) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: '15-Day validation due',
          class: 'danger',
          link: '/validations'
        });
      }
      if (cmp.final_validation_date && new Date(cmp.final_validation_date) <= now) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Final validation due',
          class: 'danger',
          link: '/validations'
        });
      }
      if (cmp.invoice_status === 'Pending' || cmp.hard_copy_status === 'Pending') {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Invoice action required',
          class: 'watch',
          link: '/invoices'
        });
      }
    }

    // Real overdue electricity alerts
    for (const el of (electricityRows || [])) {
      if (el.due_date && new Date(el.due_date) < now) {
        alerts.push({
          id: el.id,
          site_code: el.site_code || 'Power',
          client: el.bill_type || 'Power Bill',
          campaign: `Due ₹${Number(el.amount || 0).toLocaleString('en-IN')}`,
          tag: 'Power bill overdue',
          class: 'danger',
          link: '/electricity'
        });
      }
    }

    res.json({
      kpis: {
        total_sites: totalSites,
        flagged_count: flaggedCount,
        active_sites: activeSites,
        nonactive_sites: nonActiveSites,
        active_campaigns: activeCampaigns,
        mounting_overdue: Number(cAll?.mounting_overdue || 0),
        validation_15_due: Number(cAll?.val15 || 0),
        final_validation_due: Number(cAll?.valFinal || 0),
        invoice_actions: Number(i?.pending || 0),
        open_proposals: Number(p?.open_proposals || 0),
        pending_bills: Number(e?.pending_bills || 0),
        electricity_overdue: Number(e?.overdue || 0),
        unpaid_electricity: Number(e?.unpaid || 0),
        average_occupancy: avgOccupancy,
        campaign_revenue: Number(cLive?.revenue || 0),
        gross_margin: Number(cLive?.margin || 0)
      },
      occupancy_buckets: occupancyBuckets,
      alerts: alerts.slice(0, 20)
    });
  } catch (err) {
    console.error('Dashboard error:', err.message);
    res.status(500).json({ message: 'Dashboard data error: ' + err.message });
  }
});
// Dedicated API Endpoints (must be registered BEFORE /api/:entity)

// Site Photos Upload & Delete
app.post('/api/sites/:id/images', auth, upload.array('files', 100), async (req, res) => {
  try {
    const site = (await q('SELECT * FROM sites WHERE id=?', [req.params.id]))[0];
    if (!site) return res.status(404).json({ message: 'Site not found' });
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    const shouldReplace = req.query.replace === 'true' || req.query.replace === '1' || req.body?.replace === 'true' || req.body?.replace === true;
    const existing = shouldReplace ? [] : (Array.isArray(flags.ppt_images) ? flags.ppt_images : []);
    const urls = (req.files || []).map(f => `/uploads/${f.filename}`);
    flags.ppt_images = shouldReplace ? urls : [...existing, ...urls];
    await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), req.params.id]);
    res.json({ site_id: site.id, site_code: site.site_code, images: flags.ppt_images, replaced: shouldReplace });
  } catch (err) {
    console.error('Image upload error:', err);
    res.status(500).json({ message: 'Upload error: ' + err.message });
  }
});

// Upload images by Site Code (e.g. MB-01)
app.post('/api/sites/code/:site_code/images', auth, upload.array('files', 100), async (req, res) => {
  try {
    const code = req.params.site_code;
    let site = (await q('SELECT * FROM sites WHERE site_code=? OR REPLACE(LOWER(site_code),"-","")=REPLACE(LOWER(?),"-","") LIMIT 1', [code, code]))[0];
    if (!site) {
      // Auto-create site record if not found so photos are never rejected
      const defaultFlags = JSON.stringify({ ppt_images: [] });
      const ins = await q(
        `INSERT INTO sites (site_code, city, area, address, size, media_type, lighting, ownership, availability, monthly_cost, monthly_rate, flags, record_status, created_at, updated_at)
         VALUES (?, 'Ahmedabad', '', '', '30x10', 'Hoarding', 'BL', 'Owned', 'Available', 0, 0, ?, 'active', NOW(), NOW())`,
        [code, defaultFlags]
      );
      site = { id: ins.insertId, site_code: code, flags: defaultFlags };
    }
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    const shouldReplace = req.query.replace === 'true' || req.query.replace === '1' || req.body?.replace === 'true' || req.body?.replace === true;
    const existing = shouldReplace ? [] : (Array.isArray(flags.ppt_images) ? flags.ppt_images : []);
    const urls = (req.files || []).map(f => `/uploads/${f.filename}`);
    flags.ppt_images = shouldReplace ? urls : [...existing, ...urls];
    await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), site.id]);
    res.json({ site_id: site.id, site_code: site.site_code, images: flags.ppt_images, replaced: shouldReplace });
  } catch (err) {
    console.error('Image upload by code error:', err);
    res.status(500).json({ message: 'Upload error: ' + err.message });
  }
});

app.delete('/api/sites/:id/images/:index', auth, async (req, res) => {
  try {
    const site = (await q('SELECT * FROM sites WHERE id=?', [req.params.id]))[0];
    if (!site) return res.status(404).json({ message: 'Site not found' });
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    let existing = Array.isArray(flags.ppt_images) ? flags.ppt_images : [];
    const idx = Number(req.params.index);
    if (!isNaN(idx) && idx >= 0 && idx < existing.length) {
      existing.splice(idx, 1);
      flags.ppt_images = existing;
      await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), req.params.id]);
    }
    res.json({ images: flags.ppt_images });
  } catch (err) {
    res.status(500).json({ message: 'Delete error: ' + err.message });
  }
});

// Delete image by site code and index
app.delete('/api/sites/code/:site_code/images/:index', auth, async (req, res) => {
  try {
    const code = req.params.site_code;
    const site = (await q('SELECT * FROM sites WHERE site_code=? OR REPLACE(LOWER(site_code),"-","")=REPLACE(LOWER(?),"-","") LIMIT 1', [code, code]))[0];
    if (!site) return res.status(404).json({ message: 'Site not found' });
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    let existing = Array.isArray(flags.ppt_images) ? flags.ppt_images : [];
    const idx = Number(req.params.index);
    if (!isNaN(idx) && idx >= 0 && idx < existing.length) {
      existing.splice(idx, 1);
      flags.ppt_images = existing;
      await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), site.id]);
    }
    res.json({ images: flags.ppt_images });
  } catch (err) {
    res.status(500).json({ message: 'Delete error: ' + err.message });
  }
});

// Delete all images for a specific site
app.delete('/api/sites/:id/images', auth, async (req, res) => {
  try {
    const site = (await q('SELECT * FROM sites WHERE id=?', [req.params.id]))[0];
    if (!site) return res.status(404).json({ message: 'Site not found' });
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    flags.ppt_images = [];
    await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), req.params.id]);
    res.json({ site_id: site.id, site_code: site.site_code, images: [] });
  } catch (err) {
    res.status(500).json({ message: 'Delete all images error: ' + err.message });
  }
});

// Delete all images for a specific site by code
app.delete('/api/sites/code/:site_code/images', auth, async (req, res) => {
  try {
    const code = req.params.site_code;
    const site = (await q('SELECT * FROM sites WHERE site_code=? OR REPLACE(LOWER(site_code),"-","")=REPLACE(LOWER(?),"-","") LIMIT 1', [code, code]))[0];
    if (!site) return res.status(404).json({ message: 'Site not found' });
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    flags.ppt_images = [];
    await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), site.id]);
    res.json({ site_id: site.id, site_code: site.site_code, images: [] });
  } catch (err) {
    res.status(500).json({ message: 'Delete all images error: ' + err.message });
  }
});

// Sync site photos from client cache to database if server missing any
app.post('/api/sites/sync-photos', auth, async (req, res) => {
  try {
    const { photo_map } = req.body || {};
    if (!photo_map || typeof photo_map !== 'object') return res.json({ updated: 0 });
    let updated = 0;
    for (const [code, urls] of Object.entries(photo_map)) {
      if (!Array.isArray(urls) || urls.length === 0) continue;
      const site = (await q('SELECT * FROM sites WHERE site_code=? OR REPLACE(LOWER(site_code),"-","")=REPLACE(LOWER(?),"-","") LIMIT 1', [code, code]))[0];
      if (site) {
        let flags = {};
        try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
        if (Array.isArray(flags)) flags = { tags: flags };
        const existing = Array.isArray(flags.ppt_images) ? flags.ppt_images : [];
        if (existing.length === 0) {
          flags.ppt_images = urls;
          await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), site.id]);
          updated++;
        }
      }
    }
    res.json({ success: true, updated });
  } catch (err) {
    res.status(500).json({ message: 'Sync error: ' + err.message });
  }
});

// Clear all sites from database (for fresh import)
app.post('/api/sites/clear-all-sites', auth, async (req, res) => {
  try {
    await q('DELETE FROM sites');
    await q('ALTER TABLE sites AUTO_INCREMENT = 1').catch(() => {});
    res.json({ success: true, message: 'All sites have been deleted from the database' });
  } catch (err) {
    res.status(500).json({ message: 'Failed to clear sites: ' + err.message });
  }
});

// Clear all photos together across all or selected sites
app.post('/api/sites/clear-all-images', auth, async (req, res) => {
  try {
    const { site_ids } = req.body || {};
    if (Array.isArray(site_ids) && site_ids.length > 0) {
      for (const id of site_ids) {
        const site = (await q('SELECT id, flags FROM sites WHERE id=?', [id]))[0];
        if (site) {
          let flags = {};
          try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
          if (Array.isArray(flags)) flags = { tags: flags };
          flags.ppt_images = [];
          await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), id]);
        }
      }
    } else {
      const allSites = await q('SELECT id, flags FROM sites WHERE record_status="active"');
      for (const site of allSites) {
        let flags = {};
        try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
        if (Array.isArray(flags)) flags = { tags: flags };
        if (flags.ppt_images && flags.ppt_images.length > 0) {
          flags.ppt_images = [];
          await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), site.id]);
        }
      }
    }
    res.json({ success: true, message: 'All photos removed successfully' });
  } catch (err) {
    res.status(500).json({ message: 'Clear all images error: ' + err.message });
  }
});

// Auto-match Site Code by Location and Size
app.post('/api/sites/match', auth, async (req, res) => {
  try {
    const { location, size, width, height } = req.body || {};
    const existingSites = await q('SELECT id, site_code, address, area, city, width, height, size, media_type FROM sites WHERE record_status="active"');
    const matched = matchSiteByLocationAndSize(location, size, width, height, existingSites);
    if (matched) {
      res.json({ matched: true, site: matched });
    } else {
      res.json({ matched: false, site: null });
    }
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.get('/api/sites/match', auth, async (req, res) => {
  try {
    const { location, size, width, height } = req.query || {};
    const existingSites = await q('SELECT id, site_code, address, area, city, width, height, size, media_type FROM sites WHERE record_status="active"');
    const matched = matchSiteByLocationAndSize(location, size, width, height, existingSites);
    if (matched) {
      res.json({ matched: true, site: matched });
    } else {
      res.json({ matched: false, site: null });
    }
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Fixed PPT Pages
app.get('/api/ppt-pages', auth, async (req, res) => {
  try {
    const rows = await q("SELECT setting_key, setting_value FROM settings WHERE setting_key IN ('ppt_first_page','ppt_second_last_page','ppt_last_page')");
    const m = Object.fromEntries(rows.map(r => [r.setting_key, r.setting_value]));
    res.json({
      first: m.ppt_first_page || '',
      second_last: m.ppt_second_last_page || '',
      last: m.ppt_last_page || ''
    });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Optional public endpoint for PPT pages in case token expires during generation
app.get('/api/ppt-pages/public', async (req, res) => {
  try {
    const rows = await q("SELECT setting_key, setting_value FROM settings WHERE setting_key IN ('ppt_first_page','ppt_second_last_page','ppt_last_page')");
    const m = Object.fromEntries(rows.map(r => [r.setting_key, r.setting_value]));
    res.json({
      first: m.ppt_first_page || '',
      second_last: m.ppt_second_last_page || '',
      last: m.ppt_last_page || ''
    });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/ppt-pages/:key', auth, upload.single('file'), async (req, res) => {
  try {
    const map = { first: 'ppt_first_page', second_last: 'ppt_second_last_page', last: 'ppt_last_page' };
    const sk = map[req.params.key];
    if (!sk || !req.file) return res.status(400).json({ message: 'Invalid page upload' });
    const url = `/uploads/${req.file.filename}`;
    await q('INSERT INTO settings(setting_key,setting_value,updated_at) VALUES(?,?,NOW()) ON DUPLICATE KEY UPDATE setting_value=VALUES(setting_value),updated_at=NOW()', [sk, url]);
    res.json({ key: req.params.key, url, filename: req.file.filename });
  } catch (err) {
    res.status(500).json({ message: 'Page upload error: ' + err.message });
  }
});

app.delete('/api/ppt-pages/:key', auth, async (req, res) => {
  try {
    const map = { first: 'ppt_first_page', second_last: 'ppt_second_last_page', last: 'ppt_last_page' };
    const sk = map[req.params.key];
    if (sk) {
      await q('UPDATE settings SET setting_value="", updated_at=NOW() WHERE setting_key=?', [sk]);
    }
    res.json({ success: true, key: req.params.key });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// User Management
app.get('/api/users', auth, async (req, res) => {
  try {
    const rows = await q('SELECT id, name, email, role, status, created_at, updated_at FROM users ORDER BY id ASC');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/users', auth, admin, async (req, res) => {
  try {
    const { name, email, password, role = 'staff', status = 'active' } = req.body || {};
    if (!name || !email || !password) {
      return res.status(400).json({ message: 'Name, email, and password are required' });
    }
    const exists = await q('SELECT id FROM users WHERE email=? LIMIT 1', [email]);
    if (exists.length > 0) {
      return res.status(400).json({ message: 'A user with this email already exists' });
    }
    const hash = await bcrypt.hash(password, 10);
    const result = await q(
      'INSERT INTO users (name, email, password_hash, role, status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, NOW(), NOW())',
      [name, email, hash, role, status]
    );
    res.json({ id: result.insertId, name, email, role, status });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.put('/api/users/:id', auth, admin, async (req, res) => {
  try {
    const { name, email, password, role, status } = req.body || {};
    const userId = req.params.id;
    if (password && String(password).trim()) {
      const hash = await bcrypt.hash(password, 10);
      await q('UPDATE users SET name=?, email=?, password_hash=?, role=?, status=?, updated_at=NOW() WHERE id=?', [name, email, hash, role, status, userId]);
    } else {
      await q('UPDATE users SET name=?, email=?, role=?, status=?, updated_at=NOW() WHERE id=?', [name, email, role, status, userId]);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.delete('/api/users/:id', auth, admin, async (req, res) => {
  try {
    const userId = req.params.id;
    if (String(req.user.id) === String(userId)) {
      return res.status(400).json({ message: 'You cannot delete your own account' });
    }
    await q('DELETE FROM users WHERE id=?', [userId]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Settings
app.get('/api/settings', auth, async (req, res) => {
  try {
    const rows = await q('SELECT setting_key, setting_value FROM settings');
    res.json(Object.fromEntries(rows.map(r => [r.setting_key, r.setting_value])));
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.put('/api/settings', auth, admin, async (req, res) => {
  try {
    for (const [k, v] of Object.entries(req.body || {})) {
      await q('INSERT INTO settings(setting_key,setting_value,updated_at) VALUES(?,?,NOW()) ON DUPLICATE KEY UPDATE setting_value=VALUES(setting_value),updated_at=NOW()', [k, String(v ?? '')]);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Notifications
app.get('/api/notifications', auth, async (req, res) => {
  try {
    const rows = await q('SELECT * FROM notifications ORDER BY id DESC LIMIT 100');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/notifications/read', auth, async (req, res) => {
  try {
    if (req.body?.id) await q('UPDATE notifications SET is_read=1 WHERE id=?', [req.body.id]);
    else await q('UPDATE notifications SET is_read=1');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/notifications/clear', auth, async (req, res) => {
  try {
    await q('DELETE FROM notifications');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.delete('/api/notifications', auth, async (req, res) => {
  try {
    await q('DELETE FROM notifications');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Activity Log
app.get('/api/activity', auth, async (req, res) => {
  try {
    const rows = await q('SELECT * FROM activity_log ORDER BY id DESC LIMIT 250');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Export & Import
app.get('/api/export/json', auth, async (req, res) => {
  try {
    const out = { format: 'site-control-webapp', exported_at: new Date().toISOString() };
    for (const e of Object.values(entities)) {
      out[e.table] = await q(`SELECT * FROM \`${e.table}\``);
    }
    res.setHeader('Content-Disposition', 'attachment; filename=site-control-backup.json');
    res.json(out);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

function parseGps(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return { gps: '', latitude: null, longitude: null };
  const pair = raw.match(/^\s*(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)\s*$/);
  if (pair) return { gps: `${parseFloat(pair[1])}, ${parseFloat(pair[2])}`, latitude: parseFloat(pair[1]), longitude: parseFloat(pair[2]) };
  const dms = /(\d+(?:\.\d+)?)\s*[°]\s*(\d+(?:\.\d+)?)?\s*['′]?\s*(\d+(?:\.\d+)?)?\s*["″]?\s*([NS]).*?(\d+(?:\.\d+)?)\s*[°]\s*(\d+(?:\.\d+)?)?\s*['′]?\s*(\d+(?:\.\d+)?)?\s*["″]?\s*([EW])/i.exec(raw);
  if (dms) {
    const cv = (d, m, s) => Number(d || 0) + Number(m || 0) / 60 + Number(s || 0) / 3600;
    const lat = cv(dms[1], dms[2], dms[3]) * (dms[4].toUpperCase() === 'S' ? -1 : 1);
    const lng = cv(dms[5], dms[6], dms[7]) * (dms[8].toUpperCase() === 'W' ? -1 : 1);
    return { gps: `${lat.toFixed(6)}, ${lng.toFixed(6)}`, latitude: lat, longitude: lng };
  }
  return { gps: raw, latitude: null, longitude: null };
}

function normalizeKey(str) {
  return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getRowValue(row, keys, fallback = '') {
  const map = {};
  for (const k of Object.keys(row || {})) {
    map[normalizeKey(k)] = row[k];
  }
  for (const k of keys) {
    const nk = normalizeKey(k);
    if (map[nk] !== undefined && String(map[nk]).trim() !== '') {
      return map[nk];
    }
  }
  return fallback;
}

function parseAmount(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const str = String(val || '').trim();
  if (!str) return 0;
  const lacMatch = str.match(/^([\d.]+)\s*(?:lac|lakh|l)s?$/i);
  if (lacMatch) {
    return Math.round(parseFloat(lacMatch[1]) * 100000);
  }
  const cleaned = str.replace(/[^0-9.]/g, '');
  return parseFloat(cleaned) || 0;
}

function parseDimensions(rawW, rawH, rawSize, rawSqFt) {
  let w = parseFloat(rawW) || null;
  let h = parseFloat(rawH) || null;
  const sizeStr = String(rawSize || '').trim();

  if ((!w || !h) && sizeStr) {
    const m = sizeStr.match(/(\d+(?:\.\d+)?)\s*(?:x|\*|X|by)\s*(\d+(?:\.\d+)?)/i);
    if (m) {
      if (!w) w = parseFloat(m[1]);
      if (!h) h = parseFloat(m[2]);
    }
  }

  let sqFt = parseFloat(rawSqFt) || null;
  if (!sqFt && w && h) {
    sqFt = Math.round(w * h * 100) / 100;
  }

  const finalSize = (w && h) ? `${w}x${h} ft` : (sizeStr || (sqFt ? `${sqFt} sq ft` : ''));
  return { w, h, sqFt, size: finalSize };
}

function parseAvailability(val) {
  if (val instanceof Date && !isNaN(val)) {
    const dd = String(val.getDate()).padStart(2, '0');
    const mm = String(val.getMonth() + 1).padStart(2, '0');
    return `${dd}.${mm}.${val.getFullYear()}`;
  }
  if (typeof val === 'number' && val > 20000) {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return `${String(d.d).padStart(2, '0')}.${String(d.m).padStart(2, '0')}.${d.y}`;
  }
  const str = String(val ?? '').trim().replace(/\.$/, '');
  if (!str) return 'Available';
  const isoMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (isoMatch) {
    return `${isoMatch[3].padStart(2, '0')}.${isoMatch[2].padStart(2, '0')}.${isoMatch[1]}`;
  }
  const dmyMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    return `${dmyMatch[1].padStart(2, '0')}.${dmyMatch[2].padStart(2, '0')}.${dmyMatch[3]}`;
  }
  const low = str.toLowerCase();
  if (low === 'yes' || low === 'y' || low === 'immediate' || low === 'ready' || low === 'vacant') return 'Available';
  if (low === 'no' || low === 'n' || low === 'occupied' || low === 'booked') return 'Occupied';
  return str;
}

export function parseSiteDimensions(sizeStr, w, h) {
  let width = parseFloat(w) || 0;
  let height = parseFloat(h) || 0;
  if ((!width || !height) && sizeStr) {
    const match = String(sizeStr).match(/(\d+(?:\.\d+)?)\s*(?:ft|')?\s*[xX*×/–-]\s*(\d+(?:\.\d+)?)/);
    if (match) {
      if (!width) width = parseFloat(match[1]);
      if (!height) height = parseFloat(match[2]);
    }
  }
  return { width, height };
}

export function normalizeTokens(str) {
  if (!str) return [];
  const stopWords = new Set([
    'the', 'a', 'an', 'and', 'or', 'at', 'in', 'on', 'to', 'from', 'of', 'for',
    'nr', 'near', 'opp', 'opposite', 'behind', 'beside', 'facing', 'fcg', 'towards',
    'road', 'rd', 'cross', 'crossroad', 'junction', 'jnc', 'circle', 'bridge', 'flyover',
    'traffic', 'highway', 'hw', 'hwy', 'street', 'st', 'lane', 'sector', 'sec'
  ]);
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, ' ')
    .split(/\s+/)
    .filter(t => t.length >= 2 && !stopWords.has(t));
}

export function matchSiteByLocationAndSize(location, size, width, height, siteList) {
  if (!Array.isArray(siteList) || siteList.length === 0) return null;
  const targetDim = parseSiteDimensions(size, width, height);
  const locTokens = normalizeTokens(location);
  const cleanLoc = String(location || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  let bestSite = null;
  let bestScore = -1;

  for (const site of siteList) {
    let score = 0;
    const siteDim = parseSiteDimensions(site.size, site.width, site.height);
    let sizeMatched = false;

    // 1. Size matching (highest weight: 45 points)
    if (targetDim.width > 0 && targetDim.height > 0 && siteDim.width > 0 && siteDim.height > 0) {
      const exactMatch = (Math.abs(targetDim.width - siteDim.width) < 0.5 && Math.abs(targetDim.height - siteDim.height) < 0.5);
      const flippedMatch = (Math.abs(targetDim.width - siteDim.height) < 0.5 && Math.abs(targetDim.height - siteDim.width) < 0.5);
      if (exactMatch || flippedMatch) {
        score += 45;
        sizeMatched = true;
      } else {
        const targetArea = targetDim.width * targetDim.height;
        const siteArea = siteDim.width * siteDim.height;
        if (targetArea > 0 && Math.abs(targetArea - siteArea) / targetArea < 0.05) {
          score += 25;
          sizeMatched = true;
        }
      }
    }

    // 2. Location token overlap
    const siteText = `${site.address || ''} ${site.area || ''} ${site.city || ''} ${site.location || ''} ${site.site_code || ''}`;
    const siteTokens = new Set(normalizeTokens(siteText));
    const siteRaw = siteText.toLowerCase().replace(/[^a-z0-9]/g, '');

    let tokenMatches = 0;
    for (const t of locTokens) {
      if (siteTokens.has(t)) {
        tokenMatches += 1;
        score += 15;
      } else if (siteRaw.includes(t)) {
        tokenMatches += 0.5;
        score += 8;
      }
    }

    // Substring match
    if (cleanLoc.length >= 4 && siteRaw.includes(cleanLoc)) {
      score += 30;
    }

    // Directional / position indicator match (Left / Right / Middle / 1 / 2 / 3)
    const locLower = String(location || '').toLowerCase();
    const siteLower = siteText.toLowerCase();
    ['left', 'right', 'middle', 'center', '(1)', '(2)', '(3)', '1-3', '2-1', '2-3'].forEach(pos => {
      if (locLower.includes(pos) && siteLower.includes(pos)) {
        score += 12;
      }
    });

    if (score > bestScore && (tokenMatches > 0 || (sizeMatched && (cleanLoc.length < 3 || siteRaw.includes(cleanLoc.slice(0, 3)))))) {
      bestScore = score;
      bestSite = site;
    }
  }

  return bestScore >= 20 ? bestSite : null;
}

function cleanStr(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function parseMysqlDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) {
    return val.toISOString().slice(0, 10);
  }
  if (typeof val === 'number' && val > 20000) {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) {
      const yy = String(d.y).padStart(4, '0');
      const mm = String(d.m).padStart(2, '0');
      const dd = String(d.d).padStart(2, '0');
      return `${yy}-${mm}-${dd}`;
    }
  }
  const str = String(val).trim();
  if (!str) return null;
  const iso = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) {
    return `${iso[1]}-${String(iso[2]).padStart(2, '0')}-${String(iso[3]).padStart(2, '0')}`;
  }
  const dmy = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmy) {
    return `${dmy[3]}-${String(dmy[2]).padStart(2, '0')}-${String(dmy[1]).padStart(2, '0')}`;
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }
  return null;
}

app.post('/api/import/xlsx', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook provided' });
  try {
    const wb = XLSX.readFile(req.file.path, { cellDates: true });
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      return res.status(400).json({ message: 'Excel workbook contains no sheets' });
    }

    // Header keywords across languages and formats
    const headerKeywords = [
      'area', 'landmark', 'zone', 'locality',
      'location', 'address', 'sitename', 'site name',
      'media', 'mediatype', 'displaytype',
      'light', 'illumination',
      'w', 'width', 'h', 'height', 'size', 'dimension', 'sqft', 'sq ft',
      'rate', 'selling', 'adv fee', 'fee', 'price', 'amount', 'rent', 'cost',
      'avail', 'status',
      'sr no', 'srno', 'sno', 'serial',
      'site id', 'site code', 'siteid', 'sitecode', 'code', 'mb code', 'mbcode', 'mb no', 'mbno',
      'latitude', 'longitude', 'gps', 'coords', 'coordinates',
      'city', 'town'
    ];

    let chosenSheet = null;
    let bestHeaderIdx = 0;
    let highestScore = -1;

    // Scan sheets to find the one with the best header match
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      for (let i = 0; i < Math.min(25, rawRows.length); i++) {
        const row = rawRows[i] || [];
        let score = 0;
        for (const cell of row) {
          const txt = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!txt) continue;
          for (const kw of headerKeywords) {
            const cleanKw = kw.replace(/[^a-z0-9]/g, '');
            if (txt.includes(cleanKw)) {
              score++;
              break;
            }
          }
        }
        if (score > highestScore) {
          highestScore = score;
          chosenSheet = sheet;
          bestHeaderIdx = i;
        }
      }
    }

    if (!chosenSheet) {
      chosenSheet = wb.Sheets[wb.SheetNames[0]];
      bestHeaderIdx = 0;
    }

    // If a confident header row was found (>= 2 keywords), use it as the range origin
    const headerRange = highestScore >= 2 ? bestHeaderIdx : 0;
    const rows = XLSX.utils.sheet_to_json(chosenSheet, { range: headerRange, defval: '', raw: true });

    let updatedCount = 0, newCount = 0;

    // Load active sites for matching
    const existingSites = await q(`SELECT id, site_code, city, area, address, media_type, lighting, size, width, height,
      availability, monthly_rate, monthly_cost, latitude, longitude, gps, flags FROM sites WHERE record_status="active"`);

    // Determine current highest MB-XX site number
    let maxMbNum = 0;
    for (const st of existingSites) {
      const m = String(st.site_code || '').match(/MB[ -]?(\d+)/i);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxMbNum) maxMbNum = n;
      }
    }

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      let city = String(getRowValue(r, ['CITY', 'City', 'Town', 'District'], 'Ahmedabad')).trim() || 'Ahmedabad';
      let area = String(getRowValue(r, ['AREA', 'Area', 'Area / Landmark', 'Area/Landmark', 'Landmark', 'Zone', 'Locality', 'City/Area'])).trim();
      let location = String(getRowValue(r, ['LOCATION', 'Location', 'Full Address', 'Address', 'Site Location', 'Site Address', 'Location / Landmark', 'Road / Location', 'Site Name', 'Description'])).trim();

      // If one is missing, infer from the other
      if (!area && location) {
        const parts = location.split(/[-–,]/);
        if (parts.length > 1 && parts[0].trim().length > 2 && parts[0].trim().length < 35) {
          area = parts[0].trim();
        } else {
          area = city;
        }
      } else if (area && !location) {
        location = area;
      }

      if (!area && !location) continue; // Skip completely empty line

      // Media Type
      const rawMedia = String(getRowValue(r, ['MEDIA', 'Media', 'Media Type', 'MediaType', 'Type', 'Media_Type', 'Display Type'], '')).trim();
      let mediaType = 'Hoarding';
      if (/gantry/i.test(rawMedia)) mediaType = 'Gantry';
      else if (/unipole/i.test(rawMedia)) mediaType = 'Unipole';
      else if (/billboard/i.test(rawMedia)) mediaType = 'Billboard';
      else if (/dooh|digital|led/i.test(rawMedia)) mediaType = 'DOOH';
      else if (/kiosk/i.test(rawMedia)) mediaType = 'Kiosk';
      else if (/bridge/i.test(rawMedia)) mediaType = 'Bridge Panel';
      else if (rawMedia) mediaType = rawMedia;

      // Lighting
      const rawLight = String(getRowValue(r, ['LIGHT', 'Lighting', 'Light', 'Illumination', 'Light Type'], '')).trim();
      let lighting = 'BL';
      if (/front\s*lit|fl/i.test(rawLight)) lighting = 'FL';
      else if (/back\s*lit|bl/i.test(rawLight)) lighting = 'BL';
      else if (/non\s*lit|nl|unlit/i.test(rawLight)) lighting = 'NL';
      else if (/led|digital/i.test(rawLight)) lighting = 'LED';
      else if (rawLight) lighting = rawLight.toUpperCase();

      // Dimensions & Size
      const rawW = getRowValue(r, ['W', 'Width', 'W (ft)', 'Width (ft)', 'Width(ft)', 'W (FT)', 'Width in Feet'], '');
      const rawH = getRowValue(r, ['H', 'Height', 'H (ft)', 'Height (ft)', 'Height(ft)', 'H (FT)', 'Height in Feet'], '');
      const rawSize = getRowValue(r, ['Size', 'Dimensions', 'Size (ft)', 'Size (WxH)', 'Size(ft)', 'Dimension'], '');
      const rawSqFt = getRowValue(r, ['SQ FT', 'SQFT', 'Sq Ft', 'SqFt', 'Area (Sq Ft)', 'Area (Sqft)', 'Total Sqft', 'Total Sq Ft'], '');
      const dim = parseDimensions(rawW, rawH, rawSize, rawSqFt);

      // Availability
      const rawAvail = getRowValue(r, ['AVAILABLITY', 'AVAILABILITY', 'Availability', 'Status', 'Avail', 'Available From', 'Vacant From'], 'Available');
      const availability = parseAvailability(rawAvail);

      // Rate / Adv. Fee Per Month
      const rawSelling = getRowValue(r, [
        'Selling Amount', 'SellingAmount', 'Selling Amount (₹)', 'Adv. Fee Per Month', 'Adv Fee Per Month',
        'Adv Fee', 'Adv. Fee', 'Rate', 'Monthly Rate', 'Rate Per Month', 'Price', 'Rental', 'Cost', 'Amount'
      ], '0');
      const monthlyRate = parseAmount(rawSelling);

      // GPS & Coordinates
      const rawGps = getRowValue(r, ['Latitude Longitude', 'LatitudeLongitude', 'GPS', 'Coordinates', 'Latitude, Longitude', 'Coords', 'Lat/Long', 'Lat Long', 'Geo Coordinates'], '');
      const gpsObj = parseGps(rawGps);
      const lat = gpsObj.latitude || parseFloat(getRowValue(r, ['Latitude', 'Lat'], '')) || null;
      const lng = gpsObj.longitude || parseFloat(getRowValue(r, ['Longitude', 'Long', 'Lng', 'Lon'], '')) || null;
      const formattedGps = gpsObj.gps || (lat && lng ? `${lat}, ${lng}` : '');

      // Site Code in row
      const codeInRow = String(getRowValue(r, [
        'MB CODE', 'MB Code', 'MB_CODE', 'MB NO', 'MB No', 'MBNO', 'MB-CODE',
        'Site ID', 'site_code', 'SITE ID', 'Site Code', 'SiteCode', 'Site No', 'SITE NO',
        'Code', 'ID'
      ], '')).trim();

      // Match existing site
      let matched = null;
      const isExplicitMbCode = /^mb[-\s]?\d+/i.test(codeInRow);

      // 1. Check explicit MB-XX site code match
      if (codeInRow && isExplicitMbCode) {
        const cleanCode = cleanStr(codeInRow);
        matched = existingSites.find(s => cleanStr(s.site_code) === cleanCode);
      }

      // 2. Automatic site code detection by location & size
      if (!matched && (location || area || dim.size || dim.w || dim.h)) {
        matched = matchSiteByLocationAndSize(location || area, dim.size, dim.w, dim.h, existingSites);
      }

      // 3. Fallback direct site code / numeric match
      if (!matched && codeInRow) {
        const cleanCode = cleanStr(codeInRow);
        matched = existingSites.find(s => cleanStr(s.site_code) === cleanCode);
        if (!matched && /^\d+$/.test(codeInRow)) {
          const num = parseInt(codeInRow, 10);
          const mbVariant = `mb${String(num).padStart(2, '0')}`;
          matched = existingSites.find(s => cleanStr(s.site_code) === mbVariant || cleanStr(s.site_code) === `mb${num}`);
        }
      }

      // 4. Fallback exact or clean address match
      if (!matched && location) {
        const cleanLoc = cleanStr(location);
        matched = existingSites.find(s => cleanStr(s.address) === cleanLoc);
      }

      // 5. Fallback partial address containment
      if (!matched && location && cleanStr(location).length >= 10) {
        const cleanLoc = cleanStr(location);
        matched = existingSites.find(s => {
          const cleanAddr = cleanStr(s.address);
          if (!cleanAddr || cleanAddr.length < 10) return false;
          return cleanAddr.includes(cleanLoc) || cleanLoc.includes(cleanAddr);
        });
      }

      // 6. Fallback area + location landmark overlap
      if (!matched && area && location) {
        const cleanA = cleanStr(area);
        const cleanL = cleanStr(location);
        matched = existingSites.find(s => {
          if (cleanStr(s.area) !== cleanA) return false;
          const cleanAddr = cleanStr(s.address);
          return cleanAddr.includes(cleanL.slice(0, 15)) || cleanL.includes(cleanAddr.slice(0, 15));
        });
      }

      if (matched) {
        let oldFlags = {};
        try {
          oldFlags = typeof matched.flags === 'string' ? JSON.parse(matched.flags || '{}') : (matched.flags || {});
        } catch (e) { oldFlags = {}; }
        if (Array.isArray(oldFlags)) oldFlags = { ppt_images: oldFlags };

        const effectiveRate = monthlyRate || oldFlags.ppt_rate || matched.monthly_rate || 0;
        const mergedFlags = JSON.stringify({
          ...oldFlags,
          ppt_availability: availability,
          ppt_rate: effectiveRate,
          ppt_images: oldFlags.ppt_images || []
        });

        const finalArea = area || matched.area || '';
        const finalAddress = location || matched.address || '';
        const finalMediaType = mediaType || matched.media_type || 'Billboard';
        const finalLighting = lighting || matched.lighting || 'FL';
        const finalSize = dim.size || matched.size || '';
        const finalW = dim.w != null ? dim.w : matched.width;
        const finalH = dim.h != null ? dim.h : matched.height;
        const finalLat = lat != null ? lat : matched.latitude;
        const finalLng = lng != null ? lng : matched.longitude;
        const finalGps = formattedGps || matched.gps || '';

        await q(`UPDATE sites SET
          area = ?,
          address = ?,
          media_type = ?,
          lighting = ?,
          size = ?,
          width = ?,
          height = ?,
          availability = ?,
          monthly_rate = ?,
          latitude = ?,
          longitude = ?,
          gps = ?,
          flags = ?,
          updated_at = NOW()
          WHERE id = ?`, [
          finalArea, finalAddress, finalMediaType, finalLighting, finalSize, finalW, finalH,
          availability, effectiveRate, finalLat, finalLng, finalGps, mergedFlags, matched.id
        ]);
        updatedCount++;
      } else {
        // Generate new site code
        let newSiteCode = null;
        if (codeInRow) {
          const mbMatch = codeInRow.match(/^(?:MB[-\s]?)?(\d+)$/i);
          if (mbMatch) {
            const num = parseInt(mbMatch[1], 10);
            newSiteCode = `MB-${String(num).padStart(2, '0')}`;
            if (num > maxMbNum) maxMbNum = num;
          } else {
            newSiteCode = codeInRow;
          }
        }
        if (!newSiteCode) {
          maxMbNum++;
          newSiteCode = `MB-${String(maxMbNum).padStart(2, '0')}`;
        }

        const flags = JSON.stringify({
          ppt_availability: availability,
          ppt_rate: monthlyRate,
          ppt_images: []
        });

        await q(`INSERT INTO sites (
          site_code, city, area, address, media_type, lighting, size, width, height,
          availability, monthly_cost, monthly_rate, latitude, longitude, gps, flags, record_status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', NOW(), NOW())`, [
          newSiteCode, city, area, location, mediaType, lighting, dim.size, dim.w, dim.h,
          availability, monthlyRate, monthlyRate, lat, lng, formattedGps, flags
        ]);
        newCount++;
      }
    }

    res.json({
      success: true,
      updated: updatedCount,
      created: newCount,
      rows: updatedCount + newCount,
      message: `Successfully processed ${updatedCount + newCount} sites (${updatedCount} updated, ${newCount} created).`
    });
  } catch (err) {
    console.error('Import XLSX error:', err);
    res.status(500).json({ message: 'Excel import error: ' + err.message });
  }
});

app.post('/api/import/electricity-xlsx', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook provided' });
  try {
    await ensureElectricityColumns();

    const wb = XLSX.readFile(req.file.path, { cellDates: true });
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      return res.status(400).json({ message: 'Excel workbook contains no sheets' });
    }

    const headerKeywords = [
      'meter', 'service', 'tnumber', 't_number', 'tno',
      'billingmonth', 'billdate', 'duedate', 'paiddate', 'paymentdate', 'ecs',
      'units', 'provider', 'billtype', 'paymentamount', 'billamount',
      'amount', 'paymentstatus', 'paymentref', 'status',
      'sitecode', 'siteid', 'location', 'size', 'ssv', 'mb'
    ];

    const candidateSheets = [];
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      let bestScore = 0, bestIdx = 0;
      for (let i = 0; i < Math.min(25, rawRows.length); i++) {
        const row = rawRows[i] || [];
        let score = 0;
        for (const cell of row) {
          const txt = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!txt) continue;
          for (const kw of headerKeywords) {
            if (txt.includes(kw)) {
              score++;
              break;
            }
          }
        }
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
      if (bestScore >= 2) {
        candidateSheets.push({ name, sheet, headerIdx: bestIdx, score: bestScore });
      }
    }

    if (candidateSheets.length === 0) {
      candidateSheets.push({ name: wb.SheetNames[0], sheet: wb.Sheets[wb.SheetNames[0]], headerIdx: 0, score: 0 });
    }

    let updatedCount = 0, newCount = 0;

    // Load active sites and active electricity records
    const existingSites = await q(`SELECT id, site_code, area, address, size, width, height, meter_no FROM sites WHERE record_status="active"`);
    const existingBills = await q(`SELECT id, site_id, site_code, meter_no, service_number, t_number, billing_month, due_date, amount, payment_status, location, ecs, payment_date, paid_date FROM electricity WHERE record_status="active"`);

    for (const cs of candidateSheets) {
      const sheetName = cs.name;
      const sheetHint = /ssv/i.test(sheetName) ? 'SSV' : (/mb\b|media\s*buzz/i.test(sheetName) ? 'MB' : '');
      const rows = XLSX.utils.sheet_to_json(cs.sheet, { range: cs.headerIdx, defval: '', raw: true });

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];

        const siteCode = String(getRowValue(r, ['SITE CODE', 'Site Code', 'SiteCode', 'Site ID', 'SiteID', 'Code', 'Site'], '')).trim();
        const location = String(getRowValue(r, ['LOCATION', 'Location', 'Address', 'Full Address', 'Site Location', 'Landmark'], '')).trim();
        const size = String(getRowValue(r, ['SIZE', 'Size', 'Dimensions'], '')).trim();
        const meterNo = String(getRowValue(r, ['METER NO', 'Meter No', 'Meter Number', 'Meter #', 'Meter', 'MTR'], '')).trim();
        const serviceNo = String(getRowValue(r, ['SERVICE NUMBER', 'Service Number', 'Service No', 'Consumer No', 'Consumer Number', 'Customer ID', 'Account Number'], '')).trim();
        const tNumber = String(getRowValue(r, ['T NUMBER', 'T Number', 'T-No', 'TNo', 'Tariff Number'], '')).trim();

        // 1. Check explicit provider/bill header variations
        const rawBillType = String(getRowValue(r, [
          'BILL / PROVIDER', 'Bill / Provider', 'Bill/Provider', 'Bill/provider', 'BILL PROVIDER', 'Bill Provider', 'bill provider',
          'SSV / MB', 'SSV/MB', 'SSV MB', 'SSV', 'MB', 'SSV/MB BILL', 'SSV / MB BILL',
          'BILL TYPE', 'Bill Type', 'bill type', 'BillType', 'Bill_Type',
          'PROVIDER', 'Provider', 'provider', 'Discom', 'Company', 'Firm', 'Entity', 'Party', 'Biller',
          'BILL', 'Bill', 'bill'
        ], '')).trim();

        let detectedBillType = rawBillType;

        // 2. If not found, inspect columns that might contain provider or firm info
        if (!detectedBillType) {
          for (const k of Object.keys(r || {})) {
            const nk = normalizeKey(k);
            if (nk.includes('provider') || nk.includes('discom') || nk.includes('biller') || nk.includes('ssv') || nk.includes('billtype') || nk.includes('firm') || nk.includes('entity') || nk.includes('party')) {
              const val = String(r[k] || '').trim();
              if (val) {
                detectedBillType = val;
                break;
              }
            }
          }
        }

        // 3. If still not found, check if any cell value strictly equals 'SSV' or 'MB'
        if (!detectedBillType) {
          for (const k of Object.keys(r || {})) {
            const val = String(r[k] || '').trim().toUpperCase();
            if (val === 'SSV' || val === 'MB') {
              detectedBillType = val;
              break;
            }
          }
        }

        // 4. Normalize provider value (SSV and MB as primary)
        let billType = sheetHint || 'SSV';
        if (detectedBillType) {
          const cleanUpper = detectedBillType.toUpperCase();
          if (/SSV/i.test(cleanUpper)) {
            billType = 'SSV';
          } else if (/^MB\b/i.test(cleanUpper) || /MEDIA\s*BUZZ/i.test(cleanUpper) || cleanUpper === 'MB') {
            billType = 'MB';
          } else if (/TORRENT/i.test(cleanUpper)) {
            billType = 'Torrent Power';
          } else if (/PGVCL/i.test(cleanUpper)) {
            billType = 'PGVCL';
          } else if (/UGVCL/i.test(cleanUpper)) {
            billType = 'UGVCL';
          } else if (/MGVCL/i.test(cleanUpper)) {
            billType = 'MGVCL';
          } else if (/DGVCL/i.test(cleanUpper)) {
            billType = 'DGVCL';
          } else {
            billType = detectedBillType;
          }
        }

        let billingMonth = String(getRowValue(r, ['BILLING MONTH', 'Billing Month', 'Month', 'Bill Month', 'Period'], '')).trim();
        const rawBillDate = getRowValue(r, ['BILL DATE', 'Bill Date', 'Date of Bill', 'Issue Date'], '');
        const billDate = parseMysqlDate(rawBillDate);
        const rawDueDate = getRowValue(r, ['DUE DATE', 'Due Date', 'Payment Due Date', 'Last Date'], '');
        let dueDate = parseMysqlDate(rawDueDate);

        // Fallback due date if missing (due_date is NOT NULL in MySQL)
        if (!dueDate) {
          if (billDate) {
            const bd = new Date(billDate);
            bd.setDate(bd.getDate() + 15);
            dueDate = bd.toISOString().slice(0, 10);
          } else {
            const now = new Date();
            now.setDate(now.getDate() + 15);
            dueDate = now.toISOString().slice(0, 10);
          }
        }

        // Fallback billing month if missing
        if (!billingMonth) {
          if (billDate) {
            billingMonth = new Date(billDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          } else if (dueDate) {
            billingMonth = new Date(dueDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          } else {
            billingMonth = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          }
        }

        const units = parseFloat(getRowValue(r, ['UNITS', 'Units', 'Consumption', 'KWH', 'Total Units'], '0')) || 0;
        const rate = parseFloat(getRowValue(r, ['RATE', 'Rate', 'Unit Rate', 'Tariff Rate'], '0')) || 0;
        const otherCharges = parseFloat(getRowValue(r, ['OTHER CHARGES', 'Other Charges', 'Fixed Charges', 'Surcharge'], '0')) || 0;

        const rawAmount = getRowValue(r, ['PAYMENT AMOUNT (₹)', 'PAYMENT AMOUNT', 'Payment Amount', 'AMOUNT', 'Amount', 'Bill Amount', 'Total Amount', 'Net Amount', 'Payable Amount'], '0');
        let amount = parseAmount(rawAmount);
        if (amount === 0 && units > 0 && rate > 0) {
          amount = Math.round((units * rate + otherCharges) * 100) / 100;
        }

        const rawStatus = String(getRowValue(r, ['STATUS', 'Status', 'Payment Status', 'Paid Status'], '')).trim();
        const rawPaidDate = getRowValue(r, ['PAYMENT DATE', 'Payment Date', 'PAID DATE', 'Paid Date', 'Cleared Date', 'Pay Date'], '');
        const paidDate = parseMysqlDate(rawPaidDate);
        const paymentDate = paidDate;
        const ecs = String(getRowValue(r, ['ECS', 'Ecs', 'ecs', 'ECS STATUS', 'ECS Mandate', 'ECS / AUTO DEBIT', 'ECS/NACH', 'ECS / NACH', 'NACH'], '')).trim();
        const paymentRef = String(getRowValue(r, ['PAYMENT REF', 'Payment Ref', 'Payment Reference', 'UTR', 'Ref No', 'Transaction ID', 'Cheque No'], '')).trim();
        const notes = String(getRowValue(r, ['NOTES', 'Notes', 'Remarks', 'Comments', 'Description'], '')).trim();

        let paymentStatus = 'Pending';
        if (/paid|cleared|done|yes/i.test(rawStatus)) paymentStatus = 'Paid';
        else if (/overdue/i.test(rawStatus)) paymentStatus = 'Overdue';
        else if (paidDate || paymentRef) paymentStatus = 'Paid';
        else if (dueDate && new Date(dueDate) < new Date()) paymentStatus = 'Pending';

        // Skip row if completely empty
        if (!siteCode && !location && !meterNo && !serviceNo && amount === 0) continue;

        // Link to sites table
        let matchedSite = null;
        const isExplicitMbCode = /^mb[-\s]?\d+/i.test(siteCode);

        // 1. Check explicit MB-XX site code
        if (siteCode && isExplicitMbCode) {
          const cleanC = cleanStr(siteCode);
          matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
        }

        // 2. Auto-detect site code from location and size
        if (!matchedSite && (location || size)) {
          matchedSite = matchSiteByLocationAndSize(location, size, '', '', existingSites);
        }

        // 3. Fallback direct code or numeric match
        if (!matchedSite && siteCode) {
          const cleanC = cleanStr(siteCode);
          matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
          if (!matchedSite && /^\d+$/.test(siteCode)) {
            const num = parseInt(siteCode, 10);
            matchedSite = existingSites.find(s => cleanStr(s.site_code) === `mb${String(num).padStart(2, '0')}` || cleanStr(s.site_code) === `mb${num}`);
          }
        }
        if (!matchedSite && meterNo) {
          matchedSite = existingSites.find(s => cleanStr(s.meter_no) === cleanStr(meterNo));
        }
        if (!matchedSite && location) {
          matchedSite = existingSites.find(s => cleanStr(s.address) === cleanStr(location));
        }

        const finalSiteId = matchedSite?.id || null;
        const finalSiteCode = siteCode || matchedSite?.site_code || '';
        const finalLocation = location || matchedSite?.address || matchedSite?.area || '';
        const finalSize = size || matchedSite?.size || (matchedSite?.width && matchedSite?.height ? `${matchedSite.width}x${matchedSite.height} ft` : '');
        const finalMeterNo = meterNo || matchedSite?.meter_no || '';
        const finalServiceNo = serviceNo || '';
        const finalTNumber = tNumber || '';

        // Check if site meter_no can be backfilled
        if (matchedSite && meterNo && !matchedSite.meter_no) {
          await q('UPDATE sites SET meter_no=? WHERE id=?', [meterNo, matchedSite.id]);
        }

        // Check if bill already exists - match by unique service_number, t_number, meter_no, site_code or location
        let matchedBill = null;
        const cleanMonth = cleanStr(billingMonth);

        if (finalServiceNo) {
          const cleanSrv = cleanStr(finalServiceNo);
          matchedBill = existingBills.find(b => cleanStr(b.service_number) === cleanSrv && (!cleanMonth || cleanStr(b.billing_month) === cleanMonth));
          if (!matchedBill) {
            matchedBill = existingBills.find(b => cleanStr(b.service_number) === cleanSrv);
          }
        }
        if (!matchedBill && finalTNumber) {
          const cleanT = cleanStr(finalTNumber);
          matchedBill = existingBills.find(b => cleanStr(b.t_number) === cleanT && (!cleanMonth || cleanStr(b.billing_month) === cleanMonth));
          if (!matchedBill) {
            matchedBill = existingBills.find(b => cleanStr(b.t_number) === cleanT);
          }
        }
        if (!matchedBill && finalMeterNo) {
          const cleanMtr = cleanStr(finalMeterNo);
          matchedBill = existingBills.find(b => cleanStr(b.meter_no) === cleanMtr && (!cleanMonth || cleanStr(b.billing_month) === cleanMonth));
          if (!matchedBill) {
            matchedBill = existingBills.find(b => cleanStr(b.meter_no) === cleanMtr);
          }
        }
        if (!matchedBill && finalSiteCode && cleanMonth) {
          matchedBill = existingBills.find(b => cleanStr(b.site_code) === cleanStr(finalSiteCode) && cleanStr(b.billing_month) === cleanMonth);
        }
        if (!matchedBill && finalLocation) {
          const cleanLoc = cleanStr(finalLocation);
          matchedBill = existingBills.find(b => cleanStr(b.location) === cleanLoc);
        }

      if (matchedBill) {
        const finalSiteCodeVal = finalSiteCode || matchedBill.site_code || '';
        const finalLocVal = finalLocation || matchedBill.location || '';
        const finalMeterVal = finalMeterNo || matchedBill.meter_no || '';
        const finalSizeVal = finalSize || matchedBill.size || '';
        const finalSrvVal = finalServiceNo || matchedBill.service_number || '';
        const finalTVal = finalTNumber || matchedBill.t_number || '';
        const finalBillTypeVal = billType || matchedBill.bill_type || '';
        const finalPayRef = paymentRef || matchedBill.payment_reference || '';
        const finalNotes = notes || matchedBill.notes || '';
        const finalBillDate = billDate || matchedBill.bill_date;
        const finalPaidDate = paidDate || matchedBill.payment_date || matchedBill.paid_date;
        const finalPaymentDate = paymentDate || matchedBill.payment_date || matchedBill.paid_date;
        const finalEcs = ecs || matchedBill.ecs || '';

        await q(`UPDATE electricity SET
          site_id = ?,
          site_code = ?,
          location = ?,
          meter_no = ?,
          size = ?,
          service_number = ?,
          t_number = ?,
          bill_type = ?,
          payment_amount = ?,
          amount = ?,
          billing_month = ?,
          bill_date = ?,
          due_date = ?,
          units = ?,
          rate = ?,
          other_charges = ?,
          payment_status = ?,
          ecs = ?,
          paid_date = ?,
          payment_date = ?,
          payment_reference = ?,
          notes = ?,
          record_status = 'active',
          updated_at = NOW()
          WHERE id = ?`, [
          finalSiteId || matchedBill.site_id, finalSiteCodeVal, finalLocVal, finalMeterVal, finalSizeVal, finalSrvVal, finalTVal,
          finalBillTypeVal, amount, amount, billingMonth, finalBillDate, dueDate, units, rate, otherCharges,
          paymentStatus, finalEcs, finalPaidDate, finalPaymentDate, finalPayRef, finalNotes, matchedBill.id
        ]);
        updatedCount++;
      } else {
        await q(`INSERT INTO electricity (
          site_id, site_code, location, meter_no, size, service_number, t_number,
          bill_type, payment_amount, billing_month, bill_date, due_date, units, rate, other_charges,
          amount, payment_status, ecs, paid_date, payment_date, payment_reference, notes, record_status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', NOW(), NOW())`, [
          finalSiteId, finalSiteCode, finalLocation, finalMeterNo, finalSize, finalServiceNo, finalTNumber,
          billType, amount, billingMonth, billDate, dueDate, units, rate, otherCharges,
          amount, paymentStatus, ecs, paidDate, paymentDate, paymentRef, notes
        ]);
        newCount++;
      }
    }
  }

  res.json({
      success: true,
      updated: updatedCount,
      created: newCount,
      rows: updatedCount + newCount,
      message: `Successfully processed ${updatedCount + newCount} electricity bills (${updatedCount} updated, ${newCount} created).`
    });
  } catch (err) {
    console.error('Import Electricity XLSX error:', err);
    res.status(500).json({ message: 'Electricity bill import error: ' + err.message });
  }
});

// ── Site Block Format Parser & Helpers for Excel Import ─────────────────
function isVacantClient(val) {
  if (!val) return true;
  const s = String(val).trim().toLowerCase();
  return (
    s === 'blank' ||
    s.startsWith('blank') ||
    s.includes('blank due to') ||
    s === 'vacant' ||
    s === 'available' ||
    s === '-' ||
    s === '—' ||
    s === 'n/a' ||
    s === 'na' ||
    s === 'nil' ||
    s === 'none'
  );
}

function parseFlexibleExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof val === 'number') {
    try {
      if (typeof XLSX !== 'undefined' && XLSX?.SSF) {
        const parsed = XLSX.SSF.parse_date_code(val);
        if (parsed) {
          return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
        }
      }
    } catch {}
    const d = new Date((val - 25569) * 86400000);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const dt = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${dt}`;
    }
  }
  let s = String(val).trim();
  if (!s) return null;

  // Don't parse site codes like MB-72 as dates
  if (/^[a-zA-Z]{1,6}[-_ ]\d{1,4}/i.test(s)) return null;

  // Numeric Excel serial date in string (e.g. "45000")
  if (/^\d{5}$/.test(s)) {
    const d = new Date((Number(s) - 25569) * 86400000);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }

  // DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY or D.M.YY or D.M.YYYY
  const dmy = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (dmy) {
    let day = parseInt(dmy[1], 10);
    let month = parseInt(dmy[2], 10);
    let yearStr = dmy[3];
    let year = parseInt(yearStr, 10);
    if (yearStr.length === 2) year = 2000 + year;
    else if (yearStr.length === 3 && yearStr.startsWith('20')) year = parseInt('20' + yearStr.slice(2).padStart(2, '2'), 10); // e.g. 206 -> 2026

    if (month >= 1 && month <= 12) {
      const maxDays = new Date(year, month, 0).getDate();
      day = Math.min(day, maxDays);
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // YYYY-MM-DD or YYYY/MM/DD
  const ymd = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/);
  if (ymd) {
    let year = parseInt(ymd[1], 10);
    let month = parseInt(ymd[2], 10);
    let day = parseInt(ymd[3], 10);
    if (month >= 1 && month <= 12) {
      const maxDays = new Date(year, month, 0).getDate();
      day = Math.min(day, maxDays);
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // DD-MMM-YYYY (e.g. 15-Jan-2026, 01-Oct-24)
  const monthNames = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const dMmmY = s.match(/^(\d{1,2})[\s./-]+([A-Za-z]{3,9})[\s./-]+(\d{2,4})$/);
  if (dMmmY) {
    let day = parseInt(dMmmY[1], 10);
    const mKey = dMmmY[2].toLowerCase().slice(0, 3);
    const month = monthNames[mKey];
    let year = parseInt(dMmmY[3], 10);
    if (dMmmY[3].length === 2) year = 2000 + year;
    if (month) {
      const maxDays = new Date(year, month, 0).getDate();
      day = Math.min(day, maxDays);
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  return null;
}

function parseSiteBlockFormat(rawRows) {
  if (!rawRows || rawRows.length === 0) return [];

  // Quick check: If the header row contains flat tabular columns (e.g. Client + MB Code + Up Date + Down Date),
  // this is a flat tabular sheet, NOT a site-wise block schedule.
  for (let i = 0; i < Math.min(5, rawRows.length); i++) {
    const rowStr = (rawRows[i] || []).map(c => String(c || '').toLowerCase()).join(' ');
    if ((rowStr.includes('client') || rowStr.includes('display')) &&
        (rowStr.includes('mb code') || rowStr.includes('site code')) &&
        (rowStr.includes('up date') || rowStr.includes('start date') || rowStr.includes('from date')) &&
        (rowStr.includes('down date') || rowStr.includes('end date') || rowStr.includes('to date'))) {
      return [];
    }
  }

  const blocks = [];
  let currentSite = null;
  let colMapping = null;

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i] || [];
    // Clean string values of all cells in this row
    const cells = row.map(c => (c !== undefined && c !== null ? String(c).trim() : ''));
    if (cells.every(c => !c)) continue;

    // Check if this row is the top header (e.g. ["Site Code", ...])
    const normRow = cells.map(c => c.toLowerCase().replace(/[^a-z]/g, ''));
    if (normRow.includes('sitecode') && !normRow.some(c => /update|downdate/.test(c))) {
      continue;
    }

    // Check if subheader row (Up Date / Down Date / Display / Client)
    const isSubheader = normRow.some(c => c.includes('update') || c.includes('start') || c === 'from') &&
                        normRow.some(c => c.includes('downdate') || c.includes('end') || c === 'to');
    if (isSubheader) {
      let upCol = -1, downCol = -1, dispCol = -1, clientCol = -1;
      cells.forEach((c, idx) => {
        const nc = normRow[idx] || '';
        if (nc.includes('update') || nc.includes('start') || nc === 'from') upCol = idx;
        else if (nc.includes('downdate') || nc.includes('end') || nc === 'to') downCol = idx;
        else if (nc.includes('display') || nc.includes('campaign') || nc.includes('brand')) dispCol = idx;
        else if (nc.includes('client') || nc.includes('party') || nc.includes('agency')) clientCol = idx;
      });
      colMapping = { upCol, downCol, dispCol, clientCol };
      continue;
    }

    // Check if this row is a Site Header row (contains site code like MB-01, MB-14, etc.)
    let siteCodeFound = null;
    let siteLocFound = '';
    for (let c = 0; c < Math.min(3, cells.length); c++) {
      const val = cells[c];
      const isSC = /^(?:MB|DEL|BOM|AHM|SUR|RAJ|SITE|HOARDING|H)[\s_-]*\d+[A-Z]?$/i.test(val) ||
                   /^[A-Z]{1,6}[-_ ]?\d{1,4}[A-Z]?$/i.test(val);
      const isDate = parseFlexibleExcelDate(val) !== null;
      const rowHasDates = cells.some((cellVal, idx) => idx !== c && parseFlexibleExcelDate(cellVal) !== null);
      if (isSC && !isDate && !normRow[c].includes('date') && !rowHasDates) {
        siteCodeFound = val;
        // The location/description is the next non-empty cell in the row
        for (let nextC = c + 1; nextC < cells.length; nextC++) {
          if (cells[nextC]) {
            siteLocFound = cells[nextC];
            break;
          }
        }
        break;
      }
    }

    if (siteCodeFound) {
      currentSite = {
        site_code: siteCodeFound,
        location: siteLocFound,
        bookings: [],
        vacancies: []
      };
      blocks.push(currentSite);
      continue;
    }

    // If we have a current site, check if this row has date cells
    if (currentSite) {
      let upDateVal = null, downDateVal = null, clientVal = '';

      if (colMapping && colMapping.upCol >= 0 && colMapping.downCol >= 0) {
        upDateVal = parseFlexibleExcelDate(cells[colMapping.upCol]);
        downDateVal = parseFlexibleExcelDate(cells[colMapping.downCol]) || upDateVal;

        if (colMapping.dispCol >= 0 && cells[colMapping.dispCol]) {
          clientVal = cells[colMapping.dispCol];
        } else if (colMapping.clientCol >= 0 && cells[colMapping.clientCol]) {
          clientVal = cells[colMapping.clientCol];
        } else {
          for (let nextC = Math.max(colMapping.upCol, colMapping.downCol) + 1; nextC < cells.length; nextC++) {
            if (cells[nextC] && parseFlexibleExcelDate(cells[nextC]) === null) {
              clientVal = cells[nextC];
              break;
            }
          }
        }
      } else {
        // Fallback: search for first cell that has a valid date
        let dateIdx = -1;
        for (let c = 0; c < Math.min(4, cells.length); c++) {
          if (parseFlexibleExcelDate(cells[c]) !== null) {
            dateIdx = c;
            break;
          }
        }

        if (dateIdx >= 0) {
          upDateVal = parseFlexibleExcelDate(cells[dateIdx]);
          downDateVal = parseFlexibleExcelDate(cells[dateIdx + 1]) || upDateVal;
          for (let nextC = dateIdx + 2; nextC < cells.length; nextC++) {
            if (cells[nextC] && parseFlexibleExcelDate(cells[nextC]) === null) {
              clientVal = cells[nextC];
              break;
            }
          }
          if (!clientVal) clientVal = cells[dateIdx + 2] || cells[dateIdx + 1] || '';
        }
      }

      if (upDateVal) {
        const vacant = isVacantClient(clientVal);
        const item = {
          site_code: currentSite.site_code,
          location: currentSite.location,
          start_date: upDateVal,
          end_date: downDateVal || upDateVal,
          client: vacant ? 'Blank' : clientVal,
          is_vacant: vacant
        };

        if (vacant) {
          currentSite.vacancies.push(item);
        } else {
          currentSite.bookings.push(item);
        }
      }
    }
  }

  return blocks;
}

app.post('/api/import/campaigns-xlsx', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook provided' });
  try {
    // Auto-migrate any missing columns before doing anything else
    await ensureCampaignColumns();

    const wb = XLSX.readFile(req.file.path, { cellDates: true, cellNF: false, cellText: false });
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      return res.status(400).json({ message: 'Excel workbook contains no sheets' });
    }

    let existingSites = await q('SELECT id, site_code, address, area, width, height, size FROM sites WHERE record_status="active"');
    let existingCampaigns = await q('SELECT id, site_code, client, display, start_date, end_date FROM campaigns WHERE record_status="active"');
    const campTableCols = await tableColumns('campaigns');

    // Check if the workbook contains the Site Block format (Site Code + Location header, followed by Up Date / Down Date / Display rows)
    const allBlocks = [];
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      const blocks = parseSiteBlockFormat(rawRows);
      if (blocks.length > 0 && blocks.some(b => b.bookings.length > 0 || b.vacancies.length > 0)) {
        allBlocks.push(...blocks);
      }
    }

    if (allBlocks.length > 0) {
      let newCount = 0;
      let updatedCount = 0;
      let vacantCount = 0;

      for (const b of allBlocks) {
        const cleanC = cleanStr(b.site_code);
        let site = existingSites.find(s => cleanStr(s.site_code) === cleanC);

        let width = null, height = null, size = '';
        let lighting = 'BL';
        if (b.location) {
          const dimMatch = b.location.match(/(\d+(?:\.\d+)?)\s*['"]?\s*[xX*×]\s*['"]?\s*(\d+(?:\.\d+)?)/);
          if (dimMatch) {
            width = parseFloat(dimMatch[1]);
            height = parseFloat(dimMatch[2]);
            size = `${width}x${height}`;
          }
          if (/\b(?:FL|Front\s*Lit)\b/i.test(b.location)) lighting = 'FL';
          else if (/\b(?:NL|Non\s*Lit)\b/i.test(b.location)) lighting = 'Non Lit';
          else if (/\b(?:BL|Back\s*Lit)\b/i.test(b.location)) lighting = 'BL';
        }

        if (!site) {
          const loc = b.location || `Site ${b.site_code}`;
          const ins = await q(
            `INSERT INTO sites (site_code, address, area, city, media_type, lighting, width, height, size, ownership, availability, record_status, created_at, updated_at)
             VALUES (?, ?, ?, 'Ahmedabad', 'Hoarding', ?, ?, ?, ?, 'Owned', 'Available', 'active', NOW(), NOW())`,
            [b.site_code, loc, loc, lighting, width, height, size]
          );
          site = { id: ins.insertId, site_code: b.site_code, address: loc, area: loc, width, height, size };
          existingSites.push(site);
        } else {
          const updateFields = [];
          const updateVals = [];
          if (b.location && (!site.address || site.address.startsWith('Site '))) {
            updateFields.push('address = ?', 'area = ?');
            updateVals.push(b.location, b.location);
          }
          if (size && !site.size) {
            updateFields.push('size = ?', 'width = ?', 'height = ?');
            updateVals.push(size, width, height);
          }
          if (updateFields.length > 0) {
            updateFields.push('updated_at = NOW()');
            updateVals.push(site.id);
            await q(`UPDATE sites SET ${updateFields.join(', ')} WHERE id = ?`, updateVals);
          }
        }

        // Process booked campaigns (real clients only, "Blank" means vacant)
        for (const bk of b.bookings) {
          const startDate = bk.start_date;
          const endDate = bk.end_date || startDate;
          let client = bk.client || '';
          let display = client;

          // Intelligently parse client and display/brand (e.g. Swagat Group (Renewal))
          if (client) {
            const parenMatch = client.match(/^([^(]+)\s*\(([^)]+)\)$/);
            if (parenMatch) {
              client = parenMatch[1].trim();
              display = parenMatch[2].trim();
            } else if (client.includes(' / ')) {
              const parts = client.split(' / ');
              client = parts[0].trim();
              display = parts.slice(1).join(' / ').trim();
            }
          }

          const location = b.location || site.address || '';
          const diffMs = new Date(endDate).getTime() - new Date(startDate).getTime();
          const days = Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1);
          let month = '';
          try {
            month = new Date(startDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          } catch {}

          const match = existingCampaigns.find(ec =>
            cleanStr(ec.site_code) === cleanC &&
            (String(ec.client || '').toLowerCase() === String(client).toLowerCase() ||
             String(ec.display || '').toLowerCase() === String(display).toLowerCase()) &&
            String(ec.start_date || '').slice(0, 10) === startDate &&
            String(ec.end_date || '').slice(0, 10) === endDate
          );

          if (match) {
            await q(
              `UPDATE campaigns SET location=?, client=?, display=?, days=?, month=?, record_status='active', updated_at=NOW() WHERE id=?`,
              [location, client, display, days, month, match.id]
            );
            updatedCount++;
          } else {
            const bookingCode = `MB-BK-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
            const insertCols = ['booking_code', 'site_id', 'site_code', 'client', 'display', 'location', 'start_date', 'end_date', 'booking_date', 'days', 'month', 'record_status'];
            const insertVals = [bookingCode, site.id, b.site_code, client, display, location, startDate, endDate, startDate, days, month, 'active'];
            if (campTableCols.has('status')) {
              insertCols.push('status');
              insertVals.push('active');
            }
            if (campTableCols.has('campaign_name')) {
              insertCols.push('campaign_name');
              insertVals.push(display || client || b.site_code);
            }
            if (campTableCols.has('brand')) {
              insertCols.push('brand');
              insertVals.push(client || display || '');
            }
            const placeholders = insertCols.map(() => '?').join(', ');
            await q(
              `INSERT INTO campaigns (${insertCols.join(', ')}, created_at, updated_at) VALUES (${placeholders}, NOW(), NOW())`,
              insertVals
            );
            newCount++;
          }
        }

        vacantCount += b.vacancies.length;
      }

      await syncSiteAvailability();

      return res.json({
        message: `✓ Successfully imported ${newCount + updatedCount} campaign booking records across ${allBlocks.length} site schedule blocks (${vacantCount} vacant periods recognized).`,
        blocks: allBlocks.length,
        created: newCount,
        updated: updatedCount,
        vacancies: vacantCount
      });
    }

    const headerKeywords = [
      'site', 'code', 'id', 'sr', 'sno', 'no', 'hoarding', 'board', 'media', 'asset',
      'month', 'period', 'date', 'day', 'duration', 'booking',
      'client', 'agency', 'advertiser', 'party', 'customer', 'account',
      'display', 'brand', 'campaign', 'creative', 'ad', 'matter', 'caption', 'product',
      'vendor', 'supplier', 'owner', 'landlord',
      'location', 'address', 'area', 'landmark', 'city', 'place', 'sitename',
      'size', 'dimension', 'measurement', 'w', 'width', 'h', 'height', 'breadth', 'sqft',
      'type', 'illumination', 'lighting',
      'start', 'end', 'from', 'to',
      'fees', 'fee', 'rent', 'rental', 'rate', 'charges', 'advt',
      'printing', 'mounting', 'p&m', 'pm', 'fabrication', 'installation', 'cost',
      'total', 'amount', 'revenue', 'value', 'gross', 'net', 'price',
      'po', 'ro', 'order', 'wo',
      'bill', 'invoice', 'inv',
      'pending', 'balance', 'due', 'unpaid', 'status'
    ];

    const candidateSheets = [];
    let overallBestSheet = wb.SheetNames[0];
    let overallBestScore = -1;
    let overallBestHeaderIdx = 0;

    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      let bestScore = 0, bestIdx = 0;
      for (let i = 0; i < Math.min(50, rawRows.length); i++) {
        const row = rawRows[i] || [];
        let score = 0;
        for (const cell of row) {
          const txt = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!txt) continue;
          for (const kw of headerKeywords) {
            if (txt === kw || txt.includes(kw)) {
              score++;
              break;
            }
          }
        }
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
      if (bestScore > overallBestScore) {
        overallBestScore = bestScore;
        overallBestSheet = name;
        overallBestHeaderIdx = bestIdx;
      }
      if (bestScore >= 1) {
        candidateSheets.push({ name, sheet, headerIdx: bestIdx, score: bestScore });
      }
    }

    if (candidateSheets.length === 0) {
      candidateSheets.push({
        name: overallBestSheet,
        sheet: wb.Sheets[overallBestSheet],
        headerIdx: Math.max(0, overallBestHeaderIdx),
        score: Math.max(0, overallBestScore)
      });
    }

    const campCols = await tableColumns('campaigns');

    let updatedCount = 0, newCount = 0;
    existingSites = await q('SELECT id, site_code, address, area, width, height, size FROM sites WHERE record_status="active"');

    const selectCols = ['id', 'site_code', 'client', 'start_date', 'end_date'];
    ['month', 'booking_date', 'display', 'vendor_name', 'location', 'po', 'total_amount'].forEach(c => {
      if (campCols.has(c)) selectCols.push(c);
    });
    existingCampaigns = await q(`SELECT ${selectCols.join(', ')} FROM campaigns WHERE record_status="active"`);

    const toYmd = (val) => {
      if (!val) return '';
      if (typeof val === 'string') {
        const m = val.match(/^(\d{4}-\d{2}-\d{2})/);
        if (m) return m[1];
      }
      try {
        const dt = new Date(val);
        if (isNaN(dt.getTime())) return '';
        const offset = dt.getTimezoneOffset() * 60000;
        return new Date(dt.getTime() - offset).toISOString().slice(0, 10);
      } catch (e) {
        return '';
      }
    };

    const parseNum = (val) => {
      if (typeof val === 'number') return isNaN(val) ? 0 : val;
      if (!val) return 0;
      const cleaned = String(val)
        .replace(/₹|\$|€|£|Rs\.?|INR|\/|-|\s/gi, '')
        .replace(/,/g, '')
        .trim();
      const n = parseFloat(cleaned);
      return isNaN(n) ? 0 : n;
    };

    const parseDate = (val) => {
      if (!val) return null;
      if (val instanceof Date) {
        if (isNaN(val.getTime())) return null;
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const d = String(val.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      if (typeof val === 'number') {
        const parsed = (XLSX?.SSF?.parse_date_code) ? XLSX.SSF.parse_date_code(val) : null;
        if (parsed) {
          const m = String(parsed.m).padStart(2, '0');
          const d = String(parsed.d).padStart(2, '0');
          return `${parsed.y}-${m}-${d}`;
        }
      }
      const s = String(val).trim();
      if (!s) return null;

      // YYYY-MM-DD or YYYY/MM/DD
      const ymd = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
      if (ymd) {
        const y = parseInt(ymd[1], 10);
        const m = parseInt(ymd[2], 10);
        let d = parseInt(ymd[3], 10);
        if (m >= 1 && m <= 12) {
          const maxDays = new Date(y, m, 0).getDate();
          d = Math.min(d, maxDays);
          return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        }
      }

      // DD-MM-YYYY or DD/MM/YYYY or DD.MM.YYYY
      const dmy = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
      if (dmy) {
        let day = parseInt(dmy[1], 10);
        const month = parseInt(dmy[2], 10);
        let year = dmy[3];
        if (year.length === 2) year = '20' + year;
        const y = parseInt(year, 10);
        if (month >= 1 && month <= 12) {
          const maxDays = new Date(y, month, 0).getDate();
          day = Math.min(day, maxDays);
          return `${y}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
      }

      // DD-MMM-YYYY or DD MMM YYYY (e.g. 15-Jan-2024, 01-Oct-24)
      const monthNames = {
        jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
        jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
      };
      const dMmmY = s.match(/^(\d{1,2})[\s./-]+([A-Za-z]{3,9})[\s./-]+(\d{2,4})/);
      if (dMmmY) {
        let day = parseInt(dMmmY[1], 10);
        const mKey = dMmmY[2].toLowerCase().substring(0, 3);
        const month = monthNames[mKey];
        if (month) {
          let year = dMmmY[3];
          if (year.length === 2) year = '20' + year;
          const y = parseInt(year, 10);
          const maxDays = new Date(y, parseInt(month, 10), 0).getDate();
          day = Math.min(day, maxDays);
          return `${y}-${month}-${String(day).padStart(2, '0')}`;
        }
      }

      const dt = new Date(s);
      if (!isNaN(dt.getTime())) {
        const y = dt.getFullYear();
        const m = String(dt.getMonth() + 1).padStart(2, '0');
        const d = String(dt.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      return null;
    };

    const normKey = (s) => String(s || '').toLowerCase().replace(/[\r\n\t_./#\-–—()[\]{}:;]/g, ' ').replace(/\s+/g, ' ').trim();
    const alphaKey = (s) => normKey(s).replace(/[^a-z0-9]/g, '');

    const failedRows = [];

    for (const cs of candidateSheets) {
      const rows = XLSX.utils.sheet_to_json(cs.sheet, { range: cs.headerIdx, defval: '', raw: true });

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        if (!r || typeof r !== 'object') continue;

        try {
          const rowKeys = Object.keys(r);

          const getVal = (patterns) => {
            // 1. Exact match
            for (const p of patterns) {
              const pNorm = normKey(p);
              const pAlpha = alphaKey(p);
              for (const key of rowKeys) {
                if (normKey(key) === pNorm || alphaKey(key) === pAlpha) {
                  const val = r[key];
                  if (val !== undefined && val !== null && String(val).trim() !== '') return val;
                }
              }
            }
            // 2. Token / word match
            for (const p of patterns) {
              const pNorm = normKey(p);
              for (const key of rowKeys) {
                const words = normKey(key).split(/\s+/);
                if (words.includes(pNorm) || normKey(key).startsWith(pNorm + ' ') || normKey(key).endsWith(' ' + pNorm)) {
                  const val = r[key];
                  if (val !== undefined && val !== null && String(val).trim() !== '') return val;
                }
              }
            }
            // 3. Substring match only for patterns > 3 chars
            for (const p of patterns) {
              const pAlpha = alphaKey(p);
              if (pAlpha.length <= 3) continue;
              for (const key of rowKeys) {
                if (alphaKey(key).includes(pAlpha)) {
                  const val = r[key];
                  if (val !== undefined && val !== null && String(val).trim() !== '') return val;
                }
              }
            }
            return '';
          };

          const siteCode = String(getVal([
            'mb code / site code', 'mb code/site code', 'mb code', 'mbcode', 'mb_code',
            'site code', 'sitecode', 'site_code', 'site no', 'siteno', 'site number',
            'site id', 'siteid', 'hoarding no', 'hoarding code', 'board no', 'board code',
            'asset id', 'media id', 'location code', 'site', 'code', 'id',
            'sr no', 's no', 'sno', 'sl no'
          ])).trim();

          let month = String(getVal([
            'month', 'billing month', 'billing_month', 'bill month', 'mon', 'period'
          ])).trim();

          const dateVal = parseDate(getVal([
            'booking date', 'booking_date', 'order date', 'ro date', 'po date', 'agreement date', 'date', 'dt'
          ]));

          let client = String(getVal([
            'client name/display', 'client name / display', 'client name', 'client_name',
            'client / display', 'client/display', 'client_display',
            'client/agency name', 'client / agency name', 'client / agency', 'client/agency',
            'client', 'agency name', 'agency_name', 'agency',
            'advertiser name', 'advertiser', 'customer name', 'customer',
            'party name', 'party', 'account name', 'account', 'bill to'
          ])).trim();

          let display = String(getVal([
            'display / brand', 'display/brand',
            'display', 'campaign name', 'campaign_name', 'campaign', 'brand name', 'brand_name',
            'brand', 'product name', 'product', 'creative name', 'creative', 'ad name', 'ad title',
            'ad content', 'caption', 'matter', 'description'
          ])).trim();

          if (!display && client) {
            const parenMatch = client.match(/^([^(]+)\s*\(([^)]+)\)$/);
            if (parenMatch) {
              client = parenMatch[1].trim();
              display = parenMatch[2].trim();
            } else if (client.includes(' / ')) {
              const parts = client.split(' / ');
              client = parts[0].trim();
              display = parts.slice(1).join(' / ').trim();
            }
          }

          const vendorName = String(getVal([
            'vendor name', 'vendor_name', 'vendor', 'supplier name', 'supplier', 'media owner', 'owner', 'landlord'
          ])).trim();

          let location = String(getVal([
            'location / specification', 'location/specification', 'location', 'specification',
            'site location', 'site name', 'sitename', 'address', 'site address',
            'area', 'landmark', 'place', 'city', 'locality', 'zone'
          ])).trim();

          let width = parseNum(getVal(['width', 'w', 'width (ft)', 'width(ft)', 'breadth', 'b']));
          let height = parseNum(getVal(['height', 'h', 'height (ft)', 'height(ft)', 'length', 'l']));
          let size = String(getVal([
            'size', 'dimension', 'dimensions', 'size in ft', 'size(ft)', 'size (ft)', 'measurement', 'sqft', 'sq ft'
          ])).trim();

          if (!size && width > 0 && height > 0) {
            size = `${width}x${height} ft`;
          } else if (size && (!width || !height)) {
            const dimMatch = size.match(/(\d+(?:\.\d+)?)\s*['"]?\s*[xX*×]\s*['"]?\s*(\d+(?:\.\d+)?)/);
            if (dimMatch) {
              if (!width) width = parseFloat(dimMatch[1]);
              if (!height) height = parseFloat(dimMatch[2]);
            }
          }

          let lighting = 'BL';
          if (location) {
            if (/\b(?:FL|Front\s*Lit)\b/i.test(location)) lighting = 'FL';
            else if (/\b(?:NL|Non\s*Lit)\b/i.test(location)) lighting = 'Non Lit';
            else if (/\b(?:BL|Back\s*Lit)\b/i.test(location)) lighting = 'BL';
            if (!size) {
              const dimMatch = location.match(/(\d+(?:\.\d+)?)\s*['"]?\s*[xX*×]\s*['"]?\s*(\d+(?:\.\d+)?)/);
              if (dimMatch) {
                width = parseFloat(dimMatch[1]);
                height = parseFloat(dimMatch[2]);
                size = `${width}x${height} ft`;
              }
            }
          }

          const type = String(getVal([
            'type', 'media type', 'media_type', 'media', 'display type', 'nature', 'format'
          ]) || 'Hoarding').trim();

          let startDate = parseDate(getVal([
            'up date', 'update', 'up_date', 'up dt', 'up',
            'start date', 'start_date', 'from date', 'from_date', 'booking start',
            'commencement date', 'period from', 'date from', 'live date', 'from', 'start', 'start dt'
          ]));

          let endDate = parseDate(getVal([
            'down date', 'downdate', 'down_date', 'down dt', 'down',
            'end date', 'end_date', 'to date', 'to_date', 'booking end',
            'completion date', 'period to', 'date to', 'expiry date', 'to', 'end', 'end dt'
          ]));

          let days = parseInt(getVal([
            'days', 'duration days', 'duration_days', 'total days', 'duration', 'no of days', 'num of days', 'tenure'
          ]), 10);

          if (isNaN(days) || days <= 0) {
            if (startDate && endDate) {
              const diffMs = new Date(endDate).getTime() - new Date(startDate).getTime();
              days = Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1);
            } else {
              days = 30;
            }
          }

          if (!startDate && dateVal) {
            startDate = dateVal;
          }

          if (!startDate && month) {
            const mMatch = month.match(/([a-zA-Z]+)[ -_]?(\d{2,4})?/);
            if (mMatch) {
              const mKey = mMatch[1].toLowerCase().substring(0, 3);
              const mNum = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' }[mKey] || '01';
              let yNum = mMatch[2] || new Date().getFullYear();
              if (String(yNum).length === 2) yNum = '20' + yNum;
              startDate = `${yNum}-${mNum}-01`;
            }
          }

          if (!startDate) {
            startDate = new Date().toISOString().slice(0, 10);
          }

          if (!endDate && startDate) {
            endDate = startDate;
          }

          if (!month && startDate) {
            try {
              const sd = new Date(startDate);
              month = sd.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            } catch (e) {
              month = '';
            }
          }

          let advtFees = parseNum(getVal([
            'advt fees per month', 'advt. fees per month', 'advt fees', 'advt. fees', 'advt_fees',
            'advt rate', 'monthly rent', 'rent per month', 'rent', 'rental', 'display charges',
            'hire charges', 'fees', 'rate', 'charges'
          ]));

          let printingMounting = parseNum(getVal([
            'printing & mounting', 'printing and mounting', 'printing & mounting cost',
            'printing_mounting_cost', 'printing mounting', 'p&m', 'p & m', 'pm cost',
            'printing cost', 'mounting cost', 'printing', 'mounting', 'fabrication', 'installation', 'production'
          ]));

          let totalAmount = parseNum(getVal([
            'total amount', 'total_amount', 'total', 'amount', 'revenue',
            'gross amount', 'net amount', 'invoice amount', 'bill amount', 'value', 'cost', 'grand total', 'deal value'
          ]));

          if (totalAmount === 0 && (advtFees > 0 || printingMounting > 0)) {
            totalAmount = advtFees + printingMounting;
          } else if (totalAmount > 0 && advtFees === 0) {
            advtFees = Math.max(0, totalAmount - printingMounting);
          }

          const po = String(getVal([
            'po', 'po no', 'po number', 'po_number', 'p.o.', 'p.o. no', 'purchase order',
            'ro', 'ro no', 'ro number', 'ro_number', 'r.o.', 'release order', 'work order', 'wo', 'order no'
          ])).trim();

          const bill = String(getVal([
            'bill', 'bill no', 'bill_no', 'bill number', 'invoice', 'invoice no', 'invoice_no',
            'invoice number', 'inv no', 'inv_no', 'bill status'
          ])).trim();

          const pending = parseNum(getVal([
            'pending', 'pending amount', 'pending_amount', 'balance', 'balance amount', 'due amount', 'outstanding', 'unpaid'
          ]));

          const remarks = String(getVal([
            'remarks', 'remark', 'notes', 'note', 'comment', 'comments'
          ])).trim();

          // Skip completely empty rows
          if (!siteCode && !client && !location && !display && totalAmount === 0 && !startDate) {
            continue;
          }

          // Skip summary / footer rows at bottom of Excel (e.g. Total, Grand Total)
          const lowClient = (client || '').toLowerCase().trim();
          const lowLoc = (location || '').toLowerCase().trim();
          const lowCode = (siteCode || '').toLowerCase().trim();
          if (
            lowClient === 'total' || lowClient === 'grand total' || lowClient === 'subtotal' || lowClient === 'sub total' ||
            lowLoc === 'total' || lowLoc === 'grand total' ||
            lowCode === 'total' || lowCode === 'grand total'
          ) {
            continue;
          }

          // Vacant / blank client rows ("Blank" in Excel means vacant site) are preserved as records
          // so the full occupancy timeline is retained without data loss.
          if (isVacantClient(client) && isVacantClient(display)) {
            client = client || 'Blank';
            display = display || client || 'Vacant';
          }
          if (!client && display) client = display;
          if (!display && client) display = client;
          if (!client && !display) {
            client = 'Blank';
            display = 'Blank';
          }

          // Link to sites table - first check explicit site code, then auto-detect by Location & Size
          let matchedSite = null;
          const isExplicitMbCode = /^mb[-\s]?\d+/i.test(siteCode);

          // 1. Explicit MB-XX site code
          if (siteCode && isExplicitMbCode) {
            const cleanC = cleanStr(siteCode);
            matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
          }

          // 2. Automatic site detection: match site code according to location and size!
          if (!matchedSite && (location || size || width || height)) {
            matchedSite = matchSiteByLocationAndSize(location, size, width, height, existingSites);
          }

          // 3. Fallback direct or plain numeric row code match
          if (!matchedSite && siteCode) {
            const cleanC = cleanStr(siteCode);
            matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
            if (!matchedSite && /\d+/.test(siteCode)) {
              const numMatch = siteCode.match(/\d+/);
              if (numMatch) {
                const num = parseInt(numMatch[0], 10);
                matchedSite = existingSites.find(s => cleanStr(s.site_code) === `mb${String(num).padStart(2, '0')}` || cleanStr(s.site_code) === `mb${num}`);
              }
            }
          }

          // 4. Auto-create site in database if it doesn't exist yet (e.g. MB-86, MB-87)
          if (!matchedSite && siteCode) {
            const cleanC = cleanStr(siteCode);
            matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
            if (!matchedSite) {
              const loc = location || `Site ${siteCode}`;
              try {
                const ins = await q(
                  `INSERT INTO sites (site_code, address, area, city, media_type, lighting, width, height, size, ownership, availability, record_status, created_at, updated_at)
                   VALUES (?, ?, ?, 'Ahmedabad', ?, ?, ?, ?, ?, 'Owned', 'Available', 'active', NOW(), NOW())`,
                  [siteCode, loc, loc, type || 'Hoarding', lighting || 'BL', width || null, height || null, size || '']
                );
                matchedSite = { id: ins.insertId, site_code: siteCode, address: loc, area: loc, width, height, size, media_type: type };
                existingSites.push(matchedSite);
              } catch (err) {
                console.warn(`[Import] Could not auto-create site ${siteCode}:`, err.message);
              }
            }
          }

          const finalSiteId = matchedSite?.id || 0;
          const finalSiteCode = matchedSite?.site_code || siteCode || (location ? `MB-${cleanStr(location).slice(0, 10).toUpperCase()}` : `MB-${i + 1}`);

          // Auto-fill missing specs from matched site
          if (matchedSite) {
            if (!location && (matchedSite.address || matchedSite.area)) {
              location = matchedSite.address || matchedSite.area;
            }
            if (!size && matchedSite.size) {
              size = matchedSite.size;
            }
            if (!width && matchedSite.width) {
              width = matchedSite.width;
            }
            if (!height && matchedSite.height) {
              height = matchedSite.height;
            }
            if ((!type || type === 'Hoarding') && matchedSite.media_type) {
              type = matchedSite.media_type;
            }
          }

          // Match existing campaign - match by PO, or site+start_date+end_date+client, or client+display+location+start_date+end_date
          // end_date & client are included so distinct bookings for the same site are never overwritten.
          const matched = existingCampaigns.find(c => {
            if (po && c.po && String(c.po).trim().toLowerCase() === po.trim().toLowerCase()) return true;
            if (finalSiteCode && c.site_code && String(c.site_code).trim().toLowerCase() === finalSiteCode.trim().toLowerCase()) {
              if (startDate && c.start_date) {
                return toYmd(c.start_date) === toYmd(startDate) &&
                       toYmd(c.end_date) === toYmd(endDate) &&
                       String(c.client || '').trim().toLowerCase() === client.trim().toLowerCase();
              }
              if (month && c.month) {
                return cleanStr(c.month) === cleanStr(month) &&
                       String(c.client || '').trim().toLowerCase() === client.trim().toLowerCase();
              }
            }
            if (client && display && location &&
                String(c.client || '').trim().toLowerCase() === client.trim().toLowerCase() &&
                String(c.display || '').trim().toLowerCase() === display.trim().toLowerCase() &&
                String(c.location || '').trim().toLowerCase() === location.trim().toLowerCase() &&
                startDate && c.start_date && toYmd(c.start_date) === toYmd(startDate) &&
                toYmd(c.end_date) === toYmd(endDate)) {
              return true;
            }
            return false;
          });

          if (matched) {
            const updateFields = [];
            const updateVals = [];

            if (finalSiteId) { updateFields.push('`site_id` = ?'); updateVals.push(finalSiteId); }
            if (finalSiteCode) { updateFields.push('`site_code` = ?'); updateVals.push(finalSiteCode); }
            if (campCols.has('month') && month) { updateFields.push('`month` = ?'); updateVals.push(month); }
            if (campCols.has('booking_date') && dateVal) { updateFields.push('`booking_date` = ?'); updateVals.push(dateVal); }
            if (campCols.has('client') && client) { updateFields.push('`client` = ?'); updateVals.push(client); }
            if (campCols.has('display') && display) { updateFields.push('`display` = ?'); updateVals.push(display); }
            if (campCols.has('campaign_name') && display) { updateFields.push('`campaign_name` = ?'); updateVals.push(display); }
            if (campCols.has('brand') && (client || display)) { updateFields.push('`brand` = ?'); updateVals.push(client || display); }
            if (campCols.has('vendor_name') && vendorName) { updateFields.push('`vendor_name` = ?'); updateVals.push(vendorName); }
            if (campCols.has('location') && location) { updateFields.push('`location` = ?'); updateVals.push(location); }
            if (campCols.has('width') && width > 0) { updateFields.push('`width` = ?'); updateVals.push(width); }
            if (campCols.has('height') && height > 0) { updateFields.push('`height` = ?'); updateVals.push(height); }
            if (campCols.has('size') && size) { updateFields.push('`size` = ?'); updateVals.push(size); }
            if (campCols.has('type') && type) { updateFields.push('`type` = ?'); updateVals.push(type); }
            if (campCols.has('start_date') && startDate) { updateFields.push('`start_date` = ?'); updateVals.push(startDate); }
            if (campCols.has('end_date') && endDate) { updateFields.push('`end_date` = ?'); updateVals.push(endDate); }
            if (campCols.has('days') && days > 0) { updateFields.push('`days` = ?'); updateVals.push(days); }
            if (campCols.has('advt_fees') && advtFees > 0) { updateFields.push('`advt_fees` = ?'); updateVals.push(advtFees); }
            if (campCols.has('printing_mounting_cost') && printingMounting > 0) { updateFields.push('`printing_mounting_cost` = ?'); updateVals.push(printingMounting); }
            if (campCols.has('total_amount') && totalAmount > 0) { updateFields.push('`total_amount` = ?'); updateVals.push(totalAmount); }
            if (campCols.has('revenue') && totalAmount > 0) { updateFields.push('`revenue` = ?'); updateVals.push(totalAmount); }
            if (campCols.has('po') && po) { updateFields.push('`po` = ?'); updateVals.push(po); }
            if (campCols.has('bill') && bill) { updateFields.push('`bill` = ?'); updateVals.push(bill); }
            if (campCols.has('pending') && pending !== 0) { updateFields.push('`pending` = ?'); updateVals.push(pending); }
            if (campCols.has('notes') && remarks) { updateFields.push('`notes` = ?'); updateVals.push(remarks); }

            updateFields.push('`record_status` = "active"', '`updated_at` = NOW()');
            updateVals.push(matched.id);

            await q(`UPDATE campaigns SET ${updateFields.join(', ')} WHERE id = ?`, updateVals);
            matched.site_code = finalSiteCode;
            matched.start_date = startDate;
            matched.end_date = endDate;
            matched.month = month;
            updatedCount++;
          } else {
            const bookingCode = `MB-BK-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
            const insertData = {
              booking_code: bookingCode,
              site_id: finalSiteId || 0,
              site_code: finalSiteCode,
              record_status: 'active'
            };
            if (campCols.has('month')) insertData.month = month;
            if (campCols.has('booking_date')) insertData.booking_date = dateVal;
            if (campCols.has('client')) insertData.client = client || display || 'Direct Client';
            if (campCols.has('display')) insertData.display = display;
            if (campCols.has('campaign_name')) insertData.campaign_name = display || client;
            if (campCols.has('brand')) insertData.brand = client || display;
            if (campCols.has('vendor_name')) insertData.vendor_name = vendorName;
            if (campCols.has('location')) insertData.location = location;
            if (campCols.has('width')) insertData.width = width || null;
            if (campCols.has('height')) insertData.height = height || null;
            if (campCols.has('size')) insertData.size = size;
            if (campCols.has('type')) insertData.type = type;
            if (campCols.has('start_date')) insertData.start_date = startDate;
            if (campCols.has('end_date')) insertData.end_date = endDate;
            if (campCols.has('days')) insertData.days = days;
            if (campCols.has('advt_fees')) insertData.advt_fees = advtFees;
            if (campCols.has('printing_mounting_cost')) insertData.printing_mounting_cost = printingMounting;
            if (campCols.has('total_amount')) insertData.total_amount = totalAmount;
            if (campCols.has('revenue')) insertData.revenue = totalAmount;
            if (campCols.has('po')) insertData.po = po;
            if (campCols.has('bill')) insertData.bill = bill;
            if (campCols.has('pending')) insertData.pending = pending;
            if (campCols.has('notes') && remarks) insertData.notes = remarks;

            const cols = Object.keys(insertData);
            const placeholders = cols.map(() => '?').join(', ');
            const vals = cols.map(c => insertData[c]);
            const insertRes = await q(
              `INSERT INTO campaigns (${cols.map(c => '`' + c + '`').join(', ')}, created_at, updated_at) VALUES (${placeholders}, NOW(), NOW())`,
              vals
            );
            existingCampaigns.push({
              id: insertRes?.insertId,
              site_code: finalSiteCode,
              start_date: startDate,
              end_date: endDate,
              month,
              po, client, display, location
            });
            newCount++;
          }
        } catch (rowErr) {
          failedRows.push({ row: i + 2, reason: rowErr.message });
          console.warn(`[Import Campaigns] Row ${i + 2} failed:`, rowErr.message);
        }
      }
    }

    // Auto-sync all linked campaigns and site availability after Excel campaign import
    await syncAllLinkedCampaigns(q);
    await syncSiteAvailability();

    const distinctSites = new Set(existingCampaigns.map(c => cleanStr(c.site_code)).filter(Boolean)).size;
    const distinctClients = new Set(existingCampaigns.map(c => String(c.client || '').trim().toLowerCase()).filter(c => c && !isVacantClient(c))).size;

    const failedNote = failedRows.length > 0 ? ` (${failedRows.length} row${failedRows.length === 1 ? '' : 's'} could not be imported - see failedRows)` : '';
    res.json({
      success: true,
      updated: updatedCount,
      created: newCount,
      rows: updatedCount + newCount,
      sitesCount: distinctSites,
      clientsCount: distinctClients,
      failedRows,
      message: `✓ Successfully processed ${updatedCount + newCount} campaigns across ${distinctSites} MB sites (${distinctClients} clients structured).${failedNote}`
    });
  } catch (err) {
    console.error('Import Campaigns XLSX error:', err);
    res.status(500).json({ message: 'Campaign import error: ' + err.message });
  }
});

app.post('/api/import/occupancy-xlsx', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook provided' });
  try {
    await q(`CREATE TABLE IF NOT EXISTS occupancy_records (
      id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
      site_code VARCHAR(60) NOT NULL DEFAULT '',
      location VARCHAR(255) NOT NULL DEFAULT '',
      city VARCHAR(120) NOT NULL DEFAULT '',
      area VARCHAR(190) NOT NULL DEFAULT '',
      size VARCHAR(60) NOT NULL DEFAULT '',
      client VARCHAR(190) NOT NULL DEFAULT '',
      brand VARCHAR(190) NOT NULL DEFAULT '',
      display VARCHAR(190) NOT NULL DEFAULT '',
      month VARCHAR(60) NOT NULL DEFAULT '',
      start_date DATE NULL,
      end_date DATE NULL,
      days INT NOT NULL DEFAULT 30,
      occupancy_pct DECIMAL(5,2) NULL,
      total_amount DECIMAL(15,2) NOT NULL DEFAULT 0,
      pending DECIMAL(15,2) NOT NULL DEFAULT 0,
      po VARCHAR(100) NOT NULL DEFAULT '',
      bill VARCHAR(100) NOT NULL DEFAULT '',
      status VARCHAR(40) NOT NULL DEFAULT 'active',
      notes TEXT NULL,
      record_status VARCHAR(20) NOT NULL DEFAULT 'active',
      created_at DATETIME NOT NULL,
      updated_at DATETIME NOT NULL,
      KEY site_code(site_code),
      KEY month(month),
      KEY record_status(record_status)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`);

    if (req.query.clear === 'true' || req.body?.clear === 'true' || req.body?.clear === true) {
      await q(`DELETE FROM occupancy_records`);
    }

    const wb = XLSX.readFile(req.file.path, { cellDates: true, cellNF: false, cellText: false });
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      return res.status(400).json({ message: 'Excel workbook contains no sheets' });
    }

    let existingSites = await q('SELECT id, site_code, address, area, city, size, width, height FROM sites WHERE record_status="active"');
    let existingCampaigns = await q('SELECT id, site_id, site_code, client, display, location, start_date, end_date, booking_date FROM campaigns WHERE record_status="active"');

    // Check if the workbook contains the Site Block format (Site Code + Location header, followed by Up Date / Down Date / Client rows)
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      const blocks = parseSiteBlockFormat(rawRows);

      if (blocks.length > 0 && blocks.some(b => b.bookings.length > 0 || b.vacancies.length > 0)) {
        let newCount = 0;
        let vacantCount = 0;

        for (const b of blocks) {
          const cleanC = cleanStr(b.site_code);
          let site = existingSites.find(s => cleanStr(s.site_code) === cleanC);

          let width = null, height = null, size = '';
          if (b.location) {
            const dimMatch = b.location.match(/(\d+(?:\.\d+)?)\s*['"]?\s*[xX*×]\s*(\d+(?:\.\d+)?)/);
            if (dimMatch) {
              width = parseFloat(dimMatch[1]);
              height = parseFloat(dimMatch[2]);
              size = `${width}x${height}`;
            }
          }

          if (!site) {
            const loc = b.location || `Site ${b.site_code}`;
            const ins = await q(
              `INSERT INTO sites (site_code, address, area, city, media_type, lighting, width, height, size, ownership, availability, record_status, created_at, updated_at)
               VALUES (?, ?, ?, 'Ahmedabad', 'Hoarding', 'BL', ?, ?, ?, 'Owned', 'Available', 'active', NOW(), NOW())`,
              [b.site_code, loc, loc, width, height, size]
            );
            site = { id: ins.insertId, site_code: b.site_code, address: loc, area: loc, width, height, size };
            existingSites.push(site);
          } else {
            const updateFields = [];
            const updateVals = [];
            if (b.location && (!site.address || site.address.startsWith('Site '))) {
              updateFields.push('address = ?', 'area = ?');
              updateVals.push(b.location, b.location);
            }
            if (size && !site.size) {
              updateFields.push('size = ?', 'width = ?', 'height = ?');
              updateVals.push(size, width, height);
            }
            if (updateFields.length > 0) {
              updateFields.push('updated_at = NOW()');
              updateVals.push(site.id);
              await q(`UPDATE sites SET ${updateFields.join(', ')} WHERE id = ?`, updateVals);
            }
          }

          // Real client bookings
          for (const bk of b.bookings) {
            const startDate = bk.start_date;
            const endDate = bk.end_date || startDate;
            const client = bk.client;
            const display = client;
            const location = b.location || site.address || '';
            const diffMs = new Date(endDate).getTime() - new Date(startDate).getTime();
            const days = Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1);
            let month = '';
            try {
              month = new Date(startDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            } catch {}

            await q(`INSERT INTO occupancy_records (
              site_code, location, city, area, size, client, brand, display,
              month, start_date, end_date, days, occupancy_pct, total_amount,
              pending, po, bill, status, record_status, created_at, updated_at
            ) VALUES (?, ?, 'Ahmedabad', ?, ?, ?, ?, ?, ?, ?, ?, ?, 100, 0, 0, '', '', 'active', 'active', NOW(), NOW())`, [
              b.site_code, location, location, site.size || size, client, client, display,
              month, startDate, endDate, days
            ]);

            // Mirror into campaigns table so Campaign Tracker and Occupancy remain 100% linked
            try {
              const match = existingCampaigns.find(ec =>
                cleanStr(ec.site_code) === cleanC &&
                String(ec.client || '').toLowerCase() === String(client).toLowerCase() &&
                String(ec.start_date || '').slice(0, 10) === startDate &&
                String(ec.end_date || '').slice(0, 10) === endDate
              );
              if (match) {
                await q(
                  `UPDATE campaigns SET location=?, display=?, days=?, month=?, record_status='active', updated_at=NOW() WHERE id=?`,
                  [location, display, days, month, match.id]
                );
              } else {
                const campCols = await tableColumns('campaigns');
                const bookingCode = `MB-BK-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
                const insertCols = ['booking_code', 'site_id', 'site_code', 'client', 'display', 'location', 'start_date', 'end_date', 'booking_date', 'days', 'month', 'record_status'];
                const insertVals = [bookingCode, site.id, b.site_code, client, display, location, startDate, endDate, startDate, days, month, 'active'];
                if (campCols.has('status')) {
                  insertCols.push('status');
                  insertVals.push('active');
                }
                const placeholders = insertCols.map(() => '?').join(', ');
                const insCamp = await q(
                  `INSERT INTO campaigns (${insertCols.join(', ')}, created_at, updated_at) VALUES (${placeholders}, NOW(), NOW())`,
                  insertVals
                );
                existingCampaigns.push({
                  id: insCamp.insertId,
                  site_id: site.id,
                  site_code: b.site_code,
                  client,
                  start_date: startDate,
                  end_date: endDate
                });
              }
            } catch (cErr) {
              console.warn('Could not mirror occupancy booking into campaigns:', cErr.message);
            }

            newCount++;
          }

          // Vacant periods ("Blank" in Excel means vacant)
          for (const vk of b.vacancies) {
            const startDate = vk.start_date;
            const endDate = vk.end_date || startDate;
            const location = b.location || site.address || '';
            const diffMs = new Date(endDate).getTime() - new Date(startDate).getTime();
            const days = Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1);
            let month = '';
            try {
              month = new Date(startDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            } catch {}

            await q(`INSERT INTO occupancy_records (
              site_code, location, city, area, size, client, brand, display,
              month, start_date, end_date, days, occupancy_pct, total_amount,
              pending, po, bill, status, record_status, created_at, updated_at
            ) VALUES (?, ?, 'Ahmedabad', ?, ?, 'Blank', '', 'Blank', ?, ?, ?, ?, 0, 0, 0, '', '', 'vacant', 'active', NOW(), NOW())`, [
              b.site_code, location, location, site.size || size,
              month, startDate, endDate, days
            ]);
            vacantCount++;
          }
        }

        try {
          await syncAllLinkedCampaigns(q);
          await syncSiteAvailability();
        } catch (_) {}

        return res.json({
          message: `✓ Successfully imported ${newCount} booking records across ${blocks.length} sites (${vacantCount} vacant periods recognized as vacant).`,
          blocks: blocks.length,
          created: newCount,
          vacancies: vacantCount
        });
      }
    }

    const headerKeywords = [
      'site', 'code', 'hoarding', 'board', 'media',
      'month', 'period', 'date', 'day', 'duration',
      'client', 'agency', 'advertiser', 'customer',
      'display', 'brand', 'campaign', 'creative',
      'location', 'address', 'area', 'city', 'landmark',
      'size', 'dimension', 'w', 'width', 'h', 'height',
      'start', 'end', 'from', 'to',
      'amount', 'total', 'rent', 'fees', 'revenue',
      'po', 'bill', 'inv', 'invoice',
      'pending', 'balance', 'occupancy', 'occ', 'status'
    ];

    const candidateSheets = [];
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      let bestScore = 0, bestIdx = 0;
      for (let i = 0; i < Math.min(30, rawRows.length); i++) {
        const row = rawRows[i] || [];
        let score = 0;
        for (const cell of row) {
          const txt = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!txt) continue;
          for (const kw of headerKeywords) {
            if (txt.includes(kw)) { score++; break; }
          }
        }
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
      if (bestScore >= 1) {
        candidateSheets.push({ name, sheet, headerIdx: bestIdx, score: bestScore });
      }
    }

    if (candidateSheets.length === 0) {
      candidateSheets.push({ name: wb.SheetNames[0], sheet: wb.Sheets[wb.SheetNames[0]], headerIdx: 0, score: 0 });
    }

    existingSites = await q('SELECT id, site_code, address, area, city, size, width, height FROM sites WHERE record_status="active"');

    const parseNum = (val) => {
      if (typeof val === 'number') return isNaN(val) ? 0 : val;
      if (!val) return 0;
      const cleaned = String(val).replace(/₹|\$|€|£|Rs\.?|INR|\/|-|\s/gi, '').replace(/,/g, '').trim();
      const n = parseFloat(cleaned);
      return isNaN(n) ? 0 : n;
    };

    const parseDate = (val) => {
      if (!val) return null;
      if (val instanceof Date) {
        if (isNaN(val.getTime())) return null;
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const d = String(val.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      if (typeof val === 'number') {
        const parsed = XLSX.SSF.parse_date_code(val);
        if (parsed) {
          return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
        }
      }
      const str = String(val).trim();
      const d = new Date(str);
      if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const dt = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${dt}`;
      }
      return null;
    };

    let importedCount = 0;

    for (const cs of candidateSheets) {
      const rows = XLSX.utils.sheet_to_json(cs.sheet, { range: cs.headerIdx, defval: '', raw: true });

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const rowMap = {};
        for (const k of Object.keys(r || {})) {
          rowMap[String(k).toLowerCase().replace(/[^a-z0-9]/g, '')] = r[k];
        }

        const getVal = (aliases) => {
          for (const a of aliases) {
            const na = a.toLowerCase().replace(/[^a-z0-9]/g, '');
            if (rowMap[na] !== undefined && String(rowMap[na]).trim() !== '') {
              return rowMap[na];
            }
          }
          return '';
        };

        const rawCode = String(getVal(['site code', 'sitecode', 'site_code', 'site id', 'code', 'site', 'hoarding no', 'hoarding', 'board no'])).trim();
        let location = String(getVal(['location', 'address', 'area', 'landmark', 'site location', 'site name'])).trim();
        let city = String(getVal(['city', 'town'])).trim() || 'Ahmedabad';
        let size = String(getVal(['size', 'dimension', 'dimensions', 'measurement', 'wxh', 'w x h'])).trim();
        const client = String(getVal(['client', 'agency', 'client name', 'advertiser', 'party', 'customer'])).trim();
        const brand = String(getVal(['brand', 'product'])).trim();
        const display = String(getVal(['display', 'campaign', 'creative', 'ad matter', 'subject'])).trim() || brand || client || 'Standard Display';
        const month = String(getVal(['month', 'period', 'billing month'])).trim();
        let startDate = parseDate(getVal(['start date', 'start_date', 'from date', 'from_date', 'booking start', 'start', 'from', 'date']));
        let endDate = parseDate(getVal(['end date', 'end_date', 'to date', 'to_date', 'booking end', 'end', 'to']));
        let days = parseInt(getVal(['days', 'duration', 'total days', 'no of days']), 10);
        if (isNaN(days) || days <= 0) {
          if (startDate && endDate) {
            const diffMs = new Date(endDate).getTime() - new Date(startDate).getTime();
            days = Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1);
          } else {
            days = 30;
          }
        }
        if (!startDate && month) {
          const mMatch = month.match(/([a-zA-Z]+)[ -_]?(\d{2,4})?/);
          if (mMatch) {
            const mKey = mMatch[1].toLowerCase().substring(0, 3);
            const mNum = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' }[mKey] || '01';
            let yNum = mMatch[2] || new Date().getFullYear();
            if (String(yNum).length === 2) yNum = '20' + yNum;
            startDate = `${yNum}-${mNum}-01`;
            if (!endDate) {
              const dEnd = new Date(Number(yNum), Number(mNum), 0);
              endDate = `${yNum}-${mNum}-${String(dEnd.getDate()).padStart(2, '0')}`;
            }
          }
        }

        const rawOcc = getVal(['occupancy', 'occ', 'occupancy rate', 'pct', 'percent', 'utilization']);
        let occPct = null;
        if (rawOcc !== '') {
          const numOcc = parseFloat(String(rawOcc).replace('%', ''));
          if (!isNaN(numOcc)) occPct = Math.min(100, Math.max(0, numOcc));
        }

        const totalAmount = parseNum(getVal(['total amount', 'total', 'amount', 'rent', 'rental', 'revenue', 'value']));
        const pending = parseNum(getVal(['pending', 'balance', 'due', 'unpaid']));
        const po = String(getVal(['po', 'po number', 'ro', 'order no'])).trim();
        const bill = String(getVal(['bill', 'bill number', 'invoice', 'inv no'])).trim();
        const rawStatus = String(getVal(['status', 'booking status'])).trim().toLowerCase();
        let status = 'active';
        if (/past|complete|finished|done/i.test(rawStatus)) status = 'past';
        else if (/upcoming|future|booked/i.test(rawStatus)) status = 'upcoming';

        if (!rawCode && !location && !client && totalAmount === 0 && !startDate) continue;
        if (/total|grand total|subtotal/i.test(client) || /total|grand total/i.test(rawCode) || /total|grand total/i.test(location)) continue;

        // Auto-match site code if missing or implicit
        let finalSiteCode = rawCode;
        let matchedSite = null;
        if (rawCode) {
          const cleanC = cleanStr(rawCode);
          matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
        }
        if (!matchedSite && (location || size)) {
          matchedSite = matchSiteByLocationAndSize(location, size, '', '', existingSites);
        }
        if (matchedSite) {
          finalSiteCode = matchedSite.site_code;
          if (!location) location = matchedSite.address || matchedSite.area;
          if (!city) city = matchedSite.city || 'Ahmedabad';
          if (!size) size = matchedSite.size;
        }
        if (!finalSiteCode) {
          finalSiteCode = location ? `MB-${cleanStr(location).slice(0, 8).toUpperCase()}` : `MB-OCC-${i + 1}`;
        }

        const isVacant = isVacantClient(client);
        if (isVacant) {
          await q(`INSERT INTO occupancy_records (
            site_code, location, city, area, size, client, brand, display,
            month, start_date, end_date, days, occupancy_pct, total_amount,
            pending, po, bill, status, record_status, created_at, updated_at
          ) VALUES (?, ?, ?, ?, ?, 'Blank', '', 'Blank', ?, ?, ?, ?, 0, 0, 0, '', '', 'vacant', 'active', NOW(), NOW())`, [
            finalSiteCode, location, city, location, size,
            month, startDate, endDate, days
          ]);
          importedCount++;
          continue;
        }

        await q(`INSERT INTO occupancy_records (
          site_code, location, city, area, size, client, brand, display,
          month, start_date, end_date, days, occupancy_pct, total_amount,
          pending, po, bill, status, record_status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', NOW(), NOW())`, [
          finalSiteCode, location, city, location, size, client, brand, display,
          month, startDate, endDate, days, occPct, totalAmount,
          pending, po, bill, status
        ]);

        // Mirror into campaigns table so Campaign Tracker and Occupancy remain 100% linked
        try {
          const bookingCode = `MB-BK-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
          await q(`INSERT INTO campaigns (
            booking_code, site_code, client, display, location, city, size,
            month, start_date, end_date, days, total_amount,
            pending, po, bill, status, record_status, created_at, updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', NOW(), NOW())`, [
            bookingCode, finalSiteCode, client || 'Standard Client', display, location, city, size,
            month, startDate, endDate, days, totalAmount,
            pending, po, bill, status
          ]);
        } catch (cErr) {
          console.warn('Could not mirror occupancy record into campaigns:', cErr.message);
        }

        importedCount++;
      }
    }

    try {
      await syncAllLinkedCampaigns(q);
      await syncSiteAvailability();
    } catch (_) {}

    res.json({
      success: true,
      count: importedCount,
      message: `Successfully imported ${importedCount} occupancy records into Occupancy tracker.`
    });
  } catch (err) {
    console.error('Import Occupancy XLSX error:', err);
    res.status(500).json({ message: 'Occupancy import error: ' + err.message });
  }
});

// ── Occupancy Records CRUD ─────────────────────────────────────────────────

// GET all occupancy records
app.get('/api/occupancy', auth, async (req, res) => {
  try {
    // Ensure table exists (in case server hasn't imported yet)
    await q(`CREATE TABLE IF NOT EXISTS occupancy_records (
      id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
      site_code VARCHAR(60) NOT NULL DEFAULT '',
      location VARCHAR(255) NOT NULL DEFAULT '',
      city VARCHAR(120) NOT NULL DEFAULT '',
      area VARCHAR(190) NOT NULL DEFAULT '',
      size VARCHAR(60) NOT NULL DEFAULT '',
      client VARCHAR(190) NOT NULL DEFAULT '',
      brand VARCHAR(190) NOT NULL DEFAULT '',
      display VARCHAR(190) NOT NULL DEFAULT '',
      month VARCHAR(60) NOT NULL DEFAULT '',
      start_date DATE NULL,
      end_date DATE NULL,
      days INT NOT NULL DEFAULT 30,
      occupancy_pct DECIMAL(5,2) NULL,
      total_amount DECIMAL(15,2) NOT NULL DEFAULT 0,
      pending DECIMAL(15,2) NOT NULL DEFAULT 0,
      po VARCHAR(100) NOT NULL DEFAULT '',
      bill VARCHAR(100) NOT NULL DEFAULT '',
      status VARCHAR(40) NOT NULL DEFAULT 'active',
      notes TEXT NULL,
      record_status VARCHAR(20) NOT NULL DEFAULT 'active',
      created_at DATETIME NOT NULL,
      updated_at DATETIME NOT NULL,
      KEY site_code(site_code),
      KEY month(month),
      KEY record_status(record_status)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`);
    try {
      await q('UPDATE occupancy_records SET client="Blank", display="Blank" WHERE status="vacant" AND (client="" OR client IS NULL OR client="Vacant");');
    } catch {}
    const rows = await q(`SELECT * FROM occupancy_records WHERE record_status="active" ORDER BY site_code ASC, start_date ASC, id ASC`);
    res.json(rows);
  } catch (err) {
    console.error('GET /api/occupancy error:', err);
    res.status(500).json({ message: 'Failed to fetch occupancy records: ' + err.message });
  }
});

// DELETE single occupancy record
app.delete('/api/occupancy/:id', auth, managerOrAdmin, async (req, res) => {
  try {
    await q(`DELETE FROM occupancy_records WHERE id = ?`, [req.params.id]);
    res.json({ success: true });
  } catch (err) {
    console.error('DELETE /api/occupancy/:id error:', err);
    res.status(500).json({ message: 'Failed to delete occupancy record: ' + err.message });
  }
});

// Batch delete occupancy records
app.post('/api/occupancy/batch-delete', auth, managerOrAdmin, async (req, res) => {
  try {
    const { ids, hard } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      // If no ids provided, delete ALL records
      await q(`DELETE FROM occupancy_records`);
      return res.json({ success: true, message: 'All occupancy records deleted.' });
    }
    if (hard) {
      await q(`DELETE FROM occupancy_records WHERE id IN (${ids.map(() => '?').join(',')})`, ids);
    } else {
      await q(`UPDATE occupancy_records SET record_status='archived', updated_at=NOW() WHERE id IN (${ids.map(() => '?').join(',')})`, ids);
    }
    res.json({ success: true, deleted: ids.length });
  } catch (err) {
    console.error('POST /api/occupancy/batch-delete error:', err);
    res.status(500).json({ message: 'Failed to delete occupancy records: ' + err.message });
  }
});

app.post('/api/import/json', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No JSON file provided' });
  try {
    const raw = JSON.parse(fs.readFileSync(req.file.path, 'utf8'));
    let totalImported = 0;
    
    // If it's a backup with table keys
    if (raw.sites && Array.isArray(raw.sites)) {
      for (const s of raw.sites) {
        if (!s.site_code) continue;
        const flags = typeof s.flags === 'string' ? s.flags : JSON.stringify(s.flags || {});
        await q(`INSERT INTO sites (
          site_code, city, area, address, size, media_type, lighting, facing,
          ownership, availability, vendor_name, meter_no, monthly_cost, monthly_rate,
          latitude, longitude, gps, notes, flags, record_status, created_at, updated_at
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'active', NOW(), NOW())
        ON DUPLICATE KEY UPDATE city=VALUES(city), area=VALUES(area), address=VALUES(address),
        size=VALUES(size), media_type=VALUES(media_type), lighting=VALUES(lighting),
        availability=VALUES(availability), monthly_rate=VALUES(monthly_rate),
        latitude=VALUES(latitude), longitude=VALUES(longitude), flags=VALUES(flags)`, [
          s.site_code, s.city || 'Ahmedabad', s.area || '', s.address || '', s.size || '',
          s.media_type || 'Hoarding', s.lighting || 'BL', s.facing || '', s.ownership || 'Owned',
          s.availability || 'Available', s.vendor_name || '', s.meter_no || '',
          Number(s.monthly_cost || 0), Number(s.monthly_rate || 0), s.latitude || null, s.longitude || null,
          s.gps || '', s.notes || '', flags
        ]);
        totalImported++;
      }
    }
    
    // Also restore clients, campaigns, electricity, vendors if present
    for (const ent of ['clients', 'campaigns', 'electricity', 'vendors', 'invoices']) {
      if (Array.isArray(raw[ent])) {
        for (const item of raw[ent]) {
          const keys = Object.keys(item).filter(k => k !== 'id');
          if (!keys.length) continue;
          const vals = keys.map(k => typeof item[k] === 'object' && item[k] !== null ? JSON.stringify(item[k]) : item[k]);
          await q(`INSERT INTO \`${ent}\` (${keys.map(k => '`' + k + '`').join(',')}, created_at, updated_at) VALUES (${keys.map(() => '?').join(',')}, NOW(), NOW())`, vals);
          totalImported++;
        }
      }
    }

    res.json({ success: true, count: totalImported });
  } catch (err) {
    res.status(500).json({ message: 'JSON import error: ' + err.message });
  }
});

app.post('/api/upload', auth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No file received' });
  const url = `/uploads/${req.file.filename}`;
  res.json({ url, name: req.file.originalname, size: req.file.size });
});

// Storage & Archives Endpoints
app.get('/api/storage', auth, async (req, res) => {
  try {
    const { category } = req.query;
    let sql = 'SELECT * FROM storage_archives WHERE record_status != "archived"';
    const params = [];
    if (category && category !== 'all') {
      sql += ' AND category = ?';
      params.push(category);
    }
    sql += ' ORDER BY id DESC LIMIT 500';
    const rows = await q(sql, params);
    const parsed = rows.map(r => {
      let meta = {};
      if (typeof r.meta_json === 'string') {
        try { meta = JSON.parse(r.meta_json); } catch {}
      } else if (r.meta_json && typeof r.meta_json === 'object') {
        meta = r.meta_json;
      }
      return { ...r, meta };
    });
    res.json(parsed);
  } catch (err) {
    console.error('Storage list error:', err);
    res.status(500).json({ message: 'Storage list error: ' + err.message });
  }
});

app.post('/api/storage', auth, notViewer, async (req, res) => {
  try {
    const { category = 'other', title, filename, file_url = '', file_size = '—', format = 'other', meta_json = {} } = req.body;
    if (!title || !filename) {
      return res.status(400).json({ message: 'Title and filename are required' });
    }
    const metaStr = typeof meta_json === 'string' ? meta_json : JSON.stringify(meta_json || {});
    const r = await q(
      'INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, "active", NOW(), NOW())',
      [category, title, filename, file_url, file_size, format.toLowerCase(), metaStr]
    );
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [r.insertId]))[0];
    let meta = {};
    try { meta = JSON.parse(item.meta_json); } catch {}
    res.json({ ...item, meta });
  } catch (err) {
    console.error('Storage create error:', err);
    res.status(500).json({ message: 'Storage create error: ' + err.message });
  }
});

app.post('/api/storage/upload', auth, notViewer, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ message: 'No file received' });
    const originalName = req.file.originalname || 'document';
    const ext = path.extname(originalName).toLowerCase();
    const fileUrl = `/uploads/${req.file.filename}`;

    // Calculate readable size
    const bytes = req.file.size || 0;
    let readableSize = '1 KB';
    if (bytes >= 1024 * 1024) readableSize = (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    else if (bytes >= 1024) readableSize = Math.round(bytes / 1024) + ' KB';
    else readableSize = bytes + ' B';

    // Guess format and category
    let format = ext.replace('.', '') || 'bin';
    let category = req.body.category || 'other';
    if (!req.body.category) {
      if (['pptx', 'ppt'].includes(format)) category = 'ppt';
      else if (['xlsx', 'xls', 'csv'].includes(format)) category = 'excel';
      else if (['json'].includes(format)) category = 'occupancy';
    }

    const title = req.body.title || originalName.replace(/\.[^/.]+$/, '').replace(/[_-]/g, ' ');
    let meta = {};
    try {
      if (req.body.meta_json) meta = typeof req.body.meta_json === 'string' ? JSON.parse(req.body.meta_json) : req.body.meta_json;
    } catch {}

    const r = await q(
      'INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, "active", NOW(), NOW())',
      [category, title, originalName, fileUrl, readableSize, format, JSON.stringify(meta)]
    );
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [r.insertId]))[0];
    res.json({ ...item, meta });
  } catch (err) {
    console.error('Storage upload error:', err);
    res.status(500).json({ message: 'Storage upload error: ' + err.message });
  }
});

app.post('/api/storage/snapshot-occupancy', auth, notViewer, async (req, res) => {
  try {
    const sites = await q('SELECT * FROM sites WHERE record_status = "active"');
    const campaigns = await q('SELECT * FROM campaigns WHERE record_status = "active"');
    const totalSites = sites.length || 22;
    
    // Live campaigns running right now
    const now = new Date();
    const activeCampaigns = campaigns.filter(c => {
      const start = c.start_date ? new Date(c.start_date) : null;
      const end = c.end_date ? new Date(c.end_date) : null;
      return (!start || start <= now) && (!end || end >= now);
    });

    const occupiedCodes = new Set();
    let totalRevenue = 0;
    activeCampaigns.forEach(c => {
      if (c.site_code) occupiedCodes.add(c.site_code.toUpperCase());
      totalRevenue += Number(c.total_amount || c.revenue || 0);
    });

    const occupiedCount = Math.max(occupiedCodes.size, sites.filter(s => ['occupied', 'booked', 'active'].includes(String(s.availability || '').toLowerCase())).length);
    const vacantCount = Math.max(0, totalSites - occupiedCount);
    const occupancyPct = totalSites > 0 ? Math.round((occupiedCount / totalSites) * 100) : 0;

    const monthLabel = req.body.month || now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    const title = `${monthLabel} Site Occupancy Snapshot`;
    const filename = `Occupancy_${monthLabel.replace(/\s+/g, '_')}.json`;

    const meta = {
      month: monthLabel,
      totalSites,
      occupiedSites: occupiedCount,
      vacantSites: vacantCount,
      occupancyPct,
      revenue: totalRevenue,
      activeCampaignsCount: activeCampaigns.length,
      timestamp: new Date().toISOString()
    };

    const r = await q(
      'INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, "active", NOW(), NOW())',
      ['occupancy', title, filename, '', '12 KB', 'json', JSON.stringify(meta)]
    );
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [r.insertId]))[0];
    res.json({ ...item, meta });
  } catch (err) {
    console.error('Storage snapshot error:', err);
    res.status(500).json({ message: 'Snapshot error: ' + err.message });
  }
});

app.post('/api/storage/batch-delete', auth, managerOrAdmin, async (req, res) => {
  try {
    const { ids } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ message: 'No IDs provided for batch deletion' });
    }
    const cleanIds = ids.map(Number).filter(n => !isNaN(n) && n > 0);
    if (!cleanIds.length) return res.status(400).json({ message: 'Invalid IDs' });

    const items = await q(`SELECT id, file_url FROM storage_archives WHERE id IN (${cleanIds.map(() => '?').join(',')})`, cleanIds);
    for (const item of items) {
      if (item.file_url && item.file_url.startsWith('/uploads/')) {
        const p = path.resolve(uploadDir, path.basename(item.file_url));
        if (fs.existsSync(p)) {
          try { fs.unlinkSync(p); } catch {}
        }
      }
    }

    await q(`DELETE FROM storage_archives WHERE id IN (${cleanIds.map(() => '?').join(',')})`, cleanIds);
    res.json({ success: true, count: cleanIds.length });
  } catch (err) {
    console.error('Storage batch-delete error:', err);
    res.status(500).json({ message: 'Batch delete error: ' + err.message });
  }
});

app.delete('/api/storage/:id', auth, managerOrAdmin, async (req, res) => {
  try {
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [req.params.id]))[0];
    if (!item) return res.status(404).json({ message: 'Record not found' });
    await q('DELETE FROM storage_archives WHERE id = ?', [req.params.id]);
    if (item.file_url && item.file_url.startsWith('/uploads/')) {
      const p = path.resolve(uploadDir, path.basename(item.file_url));
      if (fs.existsSync(p)) {
        try { fs.unlinkSync(p); } catch {}
      }
    }
    res.json({ success: true, id: Number(req.params.id) });
  } catch (err) {
    console.error('Storage delete error:', err);
    res.status(500).json({ message: 'Delete error: ' + err.message });
  }
});

// Generic CRUD endpoints for entities
app.get('/api/:entity', auth, async (req, res) => {
  try {
    const e = safeEntity(req, res);
    if (!e) return;
    const limit = Math.min(Number(req.query.limit || 1000), 5000);
    let rows;
    if (e.table === 'electricity') {
      rows = await q(`
        SELECT 
          e.*,
          IF(e.site_code IS NOT NULL AND e.site_code != '', e.site_code, COALESCE(s.site_code, '')) AS site_code,
          IF(e.location IS NOT NULL AND e.location != '', e.location, IF(s.address IS NOT NULL AND s.address != '', s.address, COALESCE(s.area, s.city, ''))) AS location,
          IF(e.size IS NOT NULL AND e.size != '', e.size, COALESCE(s.size, '')) AS size,
          IF(e.meter_no IS NOT NULL AND e.meter_no != '', e.meter_no, COALESCE(s.meter_no, '')) AS meter_no,
          COALESCE(e.service_number, '') AS service_number,
          COALESCE(e.t_number, '') AS t_number,
          COALESCE(e.bill_type, '') AS bill_type,
          COALESCE(e.ecs, '') AS ecs,
          COALESCE(e.payment_date, e.paid_date) AS payment_date,
          COALESCE(e.paid_date, e.payment_date) AS paid_date,
          IF(e.amount != 0, e.amount, COALESCE(e.payment_amount, 0)) AS amount,
          IF(e.payment_amount != 0, e.payment_amount, COALESCE(e.amount, 0)) AS payment_amount
        FROM electricity e
        LEFT JOIN sites s ON (e.site_id = s.id OR (e.site_code != '' AND e.site_code = s.site_code))
        WHERE e.record_status = 'active'
        ORDER BY e.id DESC
        LIMIT ${limit}
      `);
    } else {
      const cols = await tableColumns(e.table);
      let where = cols.has('record_status') ? "record_status != 'archived'" : "1=1";
      if (e.table === 'campaigns') {
        where += " AND LOWER(TRIM(COALESCE(client, ''))) NOT IN ('blank', 'vacant', 'unassigned', '', '-') AND LOWER(TRIM(COALESCE(client, ''))) NOT LIKE 'blank%' AND LOWER(TRIM(COALESCE(client, ''))) NOT LIKE 'vacant%'";
      }
      rows = await q(`SELECT * FROM \`${e.table}\` WHERE ${where} ORDER BY ${e.order} LIMIT ${limit}`);
    }

    for (const r of rows) {
      for (const k of ['flags', 'images_json']) {
        if (typeof r[k] === 'string') {
          try { r[k] = JSON.parse(r[k]); } catch {}
        }
      }
    }
    res.json(rows);
  } catch (err) {
    console.error(`Error loading ${req.params.entity}:`, err.message);
    res.status(500).json({ message: 'Database error: ' + err.message });
  }
});

app.get('/api/:entity/:id', auth, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const rows = await q(`SELECT * FROM \`${e.table}\` WHERE id=? LIMIT 1`, [req.params.id]);
  if (!rows[0]) return res.status(404).json({ message: 'Record not found' });
  res.json(rows[0]);
});

app.post('/api/:entity', auth, notViewer, async (req, res) => {
  try {
    const e = safeEntity(req, res);
    if (!e) return;
    if (['proposals', 'invoices', 'clients'].includes(req.params.entity) && !['admin', 'manager'].includes(req.user?.role)) {
      return res.status(403).json({ message: 'Manager or Administrator access required' });
    }
    const data = await cleanData(e.table, req.body);
    if (req.params.entity === 'electricity') {
      if (data.amount && !data.payment_amount) data.payment_amount = data.amount;
      if (data.payment_amount && !data.amount) data.amount = data.payment_amount;
      if (data.payment_date && !data.paid_date) data.paid_date = data.payment_date;
      if (data.paid_date && !data.payment_date) data.payment_date = data.paid_date;
      if (!data.due_date) data.due_date = new Date().toISOString().slice(0, 10);
      if (data.site_code || data.site_id) {
        const siteRows = await q('SELECT * FROM sites WHERE id=? OR site_code=? LIMIT 1', [data.site_id || 0, data.site_code || '']);
        if (siteRows[0]) {
          const s = siteRows[0];
          if (!data.site_code) data.site_code = s.site_code;
          if (!data.site_id) data.site_id = s.id;
          if (!data.location) data.location = s.address || s.area || s.city || '';
          if (!data.size) data.size = s.size || '';
          if (!data.meter_no) data.meter_no = s.meter_no || '';
        }
      }
    }
    if (req.params.entity === 'campaigns') {
      if (!data.booking_code) data.booking_code = `MB-BK-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
      if (!data.start_date) data.start_date = new Date().toISOString().slice(0, 10);
      if (!data.end_date) {
        const d = new Date();
        d.setDate(d.getDate() + 30);
        data.end_date = d.toISOString().slice(0, 10);
      }
    }
    const keys = Object.keys(data);
    if (!keys.length) return res.status(400).json({ message: 'No valid fields' });
    const sql = `INSERT INTO \`${e.table}\` (${keys.map(k => '`' + k + '`').join(',')},created_at,updated_at) VALUES (${keys.map(() => '?').join(',')},NOW(),NOW())`;
    const result = await q(sql, keys.map(k => data[k]));
    await audit(req.user, 'Created', req.params.entity, result.insertId, null, data);
    // Auto-sync linked campaigns and site availability after any campaign change
    if (req.params.entity === 'campaigns') {
      await syncLinkedCampaigns(q, result.insertId);
      await syncSiteAvailability();

      // Automatically convert any overlapping vacant occupancy records to occupied
      if (data.site_code) {
        try {
          const cSite = cleanStr(data.site_code);
          const cStart = data.start_date || new Date().toISOString().slice(0, 10);
          const cEnd = data.end_date || cStart;
          const clientName = data.client || data.display || 'Booked Campaign';
          await q(`UPDATE occupancy_records 
                   SET client=?, brand=?, display=?, status='active', occupancy_pct=100, updated_at=NOW()
                   WHERE record_status='active' AND status='vacant' 
                     AND REPLACE(REPLACE(UPPER(site_code), '-', ''), ' ', '') = ?
                     AND (
                       (start_date <= ? AND end_date >= ?) OR
                       (start_date >= ? AND start_date <= ?) OR
                       (end_date >= ? AND end_date <= ?)
                     )`, [
            clientName,
            data.brand || '',
            data.display || data.campaign_name || clientName,
            cSite,
            cEnd, cStart, cStart, cEnd, cStart, cEnd
          ]);
        } catch (occErr) {
          console.warn('Error auto-syncing vacant occupancy on campaign insert:', occErr.message);
        }
      }
    }
    const rows = await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [result.insertId]);
    res.status(201).json(rows[0]);
  } catch (err) {
    console.error('Error creating entity ' + req.params.entity, err.message);
    res.status(500).json({ message: err.message });
  }
});

app.put('/api/:entity/:id', auth, notViewer, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  if (['proposals', 'invoices', 'clients'].includes(req.params.entity) && !['admin', 'manager'].includes(req.user?.role)) {
    return res.status(403).json({ message: 'Manager or Administrator access required' });
  }
  const before = (await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [req.params.id]))[0];
  if (!before) return res.status(404).json({ message: 'Record not found' });
  const data = await cleanData(e.table, req.body);
  if (req.params.entity === 'electricity') {
    if (data.amount && !data.payment_amount) data.payment_amount = data.amount;
    if (data.payment_amount && !data.amount) data.amount = data.payment_amount;
    if (data.payment_date && !data.paid_date) data.paid_date = data.payment_date;
    if (data.paid_date && !data.payment_date) data.payment_date = data.paid_date;
    if (data.site_code || data.site_id) {
      const siteRows = await q('SELECT * FROM sites WHERE id=? OR site_code=? LIMIT 1', [data.site_id || 0, data.site_code || '']);
      if (siteRows[0]) {
        const s = siteRows[0];
        if (!data.site_code && !before.site_code) data.site_code = s.site_code;
        if (!data.location && !before.location) data.location = s.address || s.area || s.city || '';
        if (!data.size && !before.size) data.size = s.size || '';
        if (!data.meter_no && !before.meter_no) data.meter_no = s.meter_no || '';
      }
    }
  }
  if (req.params.entity === 'campaigns') {
    if (data.total_amount && !data.revenue) data.revenue = data.total_amount;
    if (data.revenue && !data.total_amount) data.total_amount = data.revenue;
    if (data.display && !data.campaign_name) data.campaign_name = data.display;
    if (data.campaign_name && !data.display) data.display = data.campaign_name;
  }
  const keys = Object.keys(data);
  if (keys.length) await q(`UPDATE \`${e.table}\` SET ${keys.map(k => '`' + k + '`=?').join(',')},updated_at=NOW() WHERE id=?`, [...keys.map(k => data[k]), req.params.id]);
  const after = (await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [req.params.id]))[0];
  await audit(req.user, 'Updated', req.params.entity, req.params.id, before, after);
  // Auto-sync linked campaigns and site availability after any campaign update
  if (req.params.entity === 'campaigns') {
    await syncLinkedCampaigns(q, req.params.id);
    await syncSiteAvailability();

    const siteCode = after?.site_code || before?.site_code;
    if (siteCode && after?.record_status === 'active') {
      try {
        const cSite = cleanStr(siteCode);
        const cStart = after.start_date || before.start_date || new Date().toISOString().slice(0, 10);
        const cEnd = after.end_date || before.end_date || cStart;
        const clientName = after.client || after.display || 'Booked Campaign';
        await q(`UPDATE occupancy_records 
                 SET client=?, brand=?, display=?, status='active', occupancy_pct=100, updated_at=NOW()
                 WHERE record_status='active' AND status='vacant' 
                   AND REPLACE(REPLACE(UPPER(site_code), '-', ''), ' ', '') = ?
                   AND (
                     (start_date <= ? AND end_date >= ?) OR
                     (start_date >= ? AND start_date <= ?) OR
                     (end_date >= ? AND end_date <= ?)
                   )`, [
          clientName,
          after.brand || '',
          after.display || after.campaign_name || clientName,
          cSite,
          cEnd, cStart, cStart, cEnd, cStart, cEnd
        ]);
      } catch (e) {
        console.warn('Error auto-syncing vacant occupancy on campaign update:', e.message);
      }
    }
  }
  res.json(after);
});

app.delete('/api/:entity/:id', auth, managerOrAdmin, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const cols = await tableColumns(e.table);
  const isHardDelete = e.table === 'campaigns' || e.table === 'electricity' || req.query.hard === 'true' || !cols.has('record_status');
  if (isHardDelete) {
    await q(`DELETE FROM \`${e.table}\` WHERE id=?`, [req.params.id]);
    await audit(req.user, 'Deleted', req.params.entity, req.params.id);
  } else {
    await q(`UPDATE \`${e.table}\` SET record_status='archived',updated_at=NOW() WHERE id=?`, [req.params.id]);
    await audit(req.user, 'Archived', req.params.entity, req.params.id);
  }
  // Auto-delete linked campaigns and sync site availability after any campaign deletion
  if (req.params.entity === 'campaigns') {
    await deleteLinkedCampaigns(q, req.params.id);
    await syncSiteAvailability();
  }
  res.json({ success: true });
});

app.post('/api/:entity/batch-delete', auth, managerOrAdmin, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const { ids, hard } = req.body || {};
  if (!Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ message: 'No IDs provided for batch deletion' });
  }
  const cleanIds = ids.map(x => Number(x)).filter(x => !isNaN(x) && x > 0);
  if (cleanIds.length === 0) {
    return res.status(400).json({ message: 'Invalid IDs for batch deletion' });
  }
  const cols = await tableColumns(e.table);
  const isHardDelete = e.table === 'campaigns' || e.table === 'electricity' || hard === true || req.query.hard === 'true' || !cols.has('record_status');
  
  const placeholders = cleanIds.map(() => '?').join(',');
  if (isHardDelete) {
    await q(`DELETE FROM \`${e.table}\` WHERE id IN (${placeholders})`, cleanIds);
    for (const id of cleanIds) {
      await audit(req.user, 'Deleted', req.params.entity, id);
    }
  } else {
    await q(`UPDATE \`${e.table}\` SET record_status='archived', updated_at=NOW() WHERE id IN (${placeholders})`, cleanIds);
    for (const id of cleanIds) {
      await audit(req.user, 'Archived', req.params.entity, id);
    }
  }
  if (req.params.entity === 'campaigns') {
    await batchDeleteLinkedCampaigns(q, cleanIds);
    await syncSiteAvailability();
  }
  res.json({ success: true, count: cleanIds.length });
});

// Linked site conflict check & manual sync endpoints
app.get('/api/sites/conflicts/:code', auth, (req, res) => {
  res.json(getConflictSummary(req.params.code) || {});
});
app.post('/api/sites/sync-availability', auth, managerOrAdmin, async (req, res) => {
  await syncSiteAvailability();
  res.json({ success: true, message: 'All site availabilities synchronized.' });
});
const distCandidates = [
  path.resolve(__dirname, '../../client/dist'),
  path.resolve(__dirname, '../client/dist'),
  path.resolve(process.cwd(), 'client/dist'),
  path.resolve(process.cwd(), 'dist')
];
const clientDist = distCandidates.find(p => fs.existsSync(p));
if (clientDist) {
  console.log(`Serving static frontend from: ${clientDist}`);
  app.use(express.static(clientDist));
  app.get('*', (req, res) => res.sendFile(path.join(clientDist, 'index.html')));
} else {
  console.warn('Frontend build dist folder not found. API routes are active.');
  app.get('/', (req, res) => res.json({ status: 'API is running', endpoints: '/api/health' }));
}

const server = (typeof PORT === 'number')
  ? app.listen(PORT, () => console.log(`Site Control API & Web App listening on port ${PORT}`))
  : app.listen(PORT, () => console.log(`Site Control API & Web App listening on socket ${PORT}`));

server.keepAliveTimeout = 65000;
server.headersTimeout = 66000;

async function initDb() {
  try {
    await pool.query('SELECT 1');
    console.log('Database connected successfully');

    // Automatically create tables from schema.sql if they do not exist
    const schemaCandidates = [
      path.resolve(__dirname, '../../sql/schema.sql'),
      path.resolve(__dirname, '../sql/schema.sql'),
      path.resolve(process.cwd(), 'sql/schema.sql'),
      path.resolve(process.cwd(), '../sql/schema.sql')
    ];
    const schemaPath = schemaCandidates.find(p => fs.existsSync(p));
    if (schemaPath) {
      try {
        const sql = fs.readFileSync(schemaPath, 'utf8');
        const stmts = sql.split(';').map(s => s.trim()).filter(s => s.length > 0);
        for (const stmt of stmts) {
          try {
            await pool.query(stmt);
          } catch (e) {
            // Ignore benign statements
          }
        }
        console.log('Database tables verified and ready.');
      } catch (schemaErr) {
        console.warn('Auto-schema execution notice:', schemaErr.message);
      }
    }

    try {
      const emails = await q('SELECT COUNT(*) c FROM users');
      if (!emails[0].c && process.env.ADMIN_EMAIL && process.env.ADMIN_PASSWORD) {
        const hash = await bcrypt.hash(process.env.ADMIN_PASSWORD, 12);
        await q('INSERT INTO users(name,email,password_hash,role,status,created_at,updated_at) VALUES(?,?,?,?,"active",NOW(),NOW())', ['Administrator', process.env.ADMIN_EMAIL, hash, 'admin']);
        console.log('Initial admin user created');
      }
    } catch (tblErr) {
      console.warn('User table check notice:', tblErr.message);
    }

    try {
      await q('ALTER TABLE electricity MODIFY due_date DATE NULL DEFAULT NULL');
      await q('ALTER TABLE campaigns MODIFY start_date DATE NULL DEFAULT NULL, MODIFY end_date DATE NULL DEFAULT NULL');
    } catch (_) {}

    await ensureCampaignColumns();
    await ensureElectricityColumns();

    try {
      await q(`CREATE TABLE IF NOT EXISTS occupancy_records (
        id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
        site_code VARCHAR(60) NOT NULL DEFAULT '',
        location VARCHAR(255) NOT NULL DEFAULT '',
        city VARCHAR(120) NOT NULL DEFAULT '',
        area VARCHAR(190) NOT NULL DEFAULT '',
        size VARCHAR(60) NOT NULL DEFAULT '',
        client VARCHAR(190) NOT NULL DEFAULT '',
        brand VARCHAR(190) NOT NULL DEFAULT '',
        display VARCHAR(190) NOT NULL DEFAULT '',
        month VARCHAR(60) NOT NULL DEFAULT '',
        start_date DATE NULL,
        end_date DATE NULL,
        days INT NOT NULL DEFAULT 30,
        occupancy_pct DECIMAL(5,2) NULL,
        total_amount DECIMAL(15,2) NOT NULL DEFAULT 0,
        pending DECIMAL(15,2) NOT NULL DEFAULT 0,
        po VARCHAR(100) NOT NULL DEFAULT '',
        bill VARCHAR(100) NOT NULL DEFAULT '',
        status VARCHAR(40) NOT NULL DEFAULT 'active',
        notes TEXT NULL,
        record_status VARCHAR(20) NOT NULL DEFAULT 'active',
        created_at DATETIME NOT NULL,
        updated_at DATETIME NOT NULL,
        KEY site_code(site_code),
        KEY month(month),
        KEY record_status(record_status)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`);
    } catch (occTblErr) {
      console.warn('Occupancy table init notice:', occTblErr.message);
    }
    try {
      await q(`CREATE TABLE IF NOT EXISTS storage_archives (
        id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
        category VARCHAR(60) NOT NULL DEFAULT 'other',
        title VARCHAR(255) NOT NULL DEFAULT '',
        filename VARCHAR(255) NOT NULL DEFAULT '',
        file_url TEXT NULL,
        file_size VARCHAR(50) NOT NULL DEFAULT '—',
        format VARCHAR(40) NOT NULL DEFAULT 'other',
        meta_json LONGTEXT NULL,
        record_status VARCHAR(20) NOT NULL DEFAULT 'active',
        created_at DATETIME NOT NULL,
        updated_at DATETIME NOT NULL,
        KEY category(category),
        KEY format(format),
        KEY record_status(record_status)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`);
    } catch (storageTblErr) {
      console.warn('Storage archives table init notice:', storageTblErr.message);
    }

    // Seed default settings if empty (sites are managed by user / Excel import)
    try {
      const seedCandidates = [
        path.resolve(__dirname, 'data/all-embedded-data.json'),
        path.resolve(__dirname, '../data/all-embedded-data.json'),
        path.resolve(process.cwd(), 'server/src/data/all-embedded-data.json'),
        path.resolve(process.cwd(), 'sql/all-embedded-data.json')
      ];
      const seedPath = seedCandidates.find(p => fs.existsSync(p));
      if (seedPath) {
        const raw = JSON.parse(fs.readFileSync(seedPath, 'utf8'));
        if (raw.default_settings) {
          const settingsCount = await q('SELECT COUNT(*) c FROM settings');
          if (!settingsCount[0]?.c) {
            for (const [k, v] of Object.entries(raw.default_settings)) {
              await q('INSERT INTO settings (setting_key, setting_value, updated_at) VALUES (?,?,NOW()) ON DUPLICATE KEY UPDATE setting_value=VALUES(setting_value)', [k, String(v ?? '')]);
            }
            console.log('Seeded default Media Buzz settings successfully.');
          }
        }
      }

      // Seed Vendors if empty
      const vendorCount = await q('SELECT COUNT(*) c FROM vendors');
      if (!vendorCount[0]?.c) {
        await q(`INSERT INTO vendors (name, service, contact_person, phone, email, cities, rating, status, created_at, updated_at) VALUES
          ('Gujarat Printers & Signage', 'Printing', 'Mukesh Bhai', '+91 98240 55441', 'mukesh@gujaratprinters.com', 'Ahmedabad, Gandhinagar, Surat', 4.85, 'active', NOW(), NOW()),
          ('Ahmedabad Neon & LED Arts', 'Mounting & Electrical', 'Haresh Solanki', '+91 98980 12345', 'info@ahmedabadneon.com', 'Ahmedabad, Vadodara', 4.90, 'active', NOW(), NOW()),
          ('Om Sai Mounting Services', 'Mounting', 'Ramesh Parmar', '+91 94260 88776', 'omsaimounting@gmail.com', 'Ahmedabad, Rajkot', 4.65, 'active', NOW(), NOW()),
          ('Apex Outdoor Fabricators', 'Fabrication & Maintenance', 'Ketan Shah', '+91 98255 33221', 'ketan@apexfab.in', 'Ahmedabad', 4.75, 'active', NOW(), NOW())`);
        console.log('Seeded initial vendors.');
      }

      // Seed Electricity if empty
      const elecCount = await q('SELECT COUNT(*) c FROM electricity');
      if (!elecCount[0]?.c) {
        await q(`INSERT INTO electricity (
          site_code, location, meter_no, size, service_number, t_number, bill_type,
          payment_amount, billing_month, bill_date, due_date, units, rate, amount,
          payment_status, paid_date, payment_reference, notes, record_status, created_at, updated_at
        ) VALUES
          ('AMD-GT-001', 'Shivranjani Cross Roads, Ahmedabad', 'MTR-UGVCL-8841', '30x10 ft', 'SRV-998241', 'T-4401', 'UGVCL', 7850, 'Aug 2026', DATE_SUB(CURDATE(), INTERVAL 18 DAY), DATE_SUB(CURDATE(), INTERVAL 2 DAY), 850, 9.23, 7850, 'Pending', NULL, '', 'Overdue meter bill', 'active', NOW(), NOW()),
          ('AMD-UP-002', 'SG Highway near YMCA Club, Ahmedabad', 'MTR-TORRENT-3312', '40x20 ft', 'SRV-887412', 'T-5512', 'Torrent Power', 12400, 'Aug 2026', DATE_SUB(CURDATE(), INTERVAL 10 DAY), DATE_ADD(CURDATE(), INTERVAL 6 DAY), 1320, 9.39, 12400, 'Pending', NULL, '', 'Due soon', 'active', NOW(), NOW()),
          ('AMD-HD-005', 'Sindhubhavan Road, Ahmedabad', 'MTR-TORRENT-1198', '50x20 ft', 'SRV-665209', 'T-2209', 'Torrent Power', 15800, 'Jul 2026', DATE_SUB(CURDATE(), INTERVAL 40 DAY), DATE_SUB(CURDATE(), INTERVAL 25 DAY), 1680, 9.40, 15800, 'Paid', DATE_SUB(CURDATE(), INTERVAL 28 DAY), 'UPI-9923847291', 'Paid on time', 'active', NOW(), NOW())`);
        console.log('Seeded initial electricity bills.');
      }

      // Seed Notifications if empty
      const notifCount = await q('SELECT COUNT(*) c FROM notifications');
      if (!notifCount[0]?.c) {
        await q(`INSERT INTO notifications (user_id, type, object_type, title, message, is_read, fingerprint, created_at) VALUES
          (0, 'electricity_due', 'electricity', 'Electricity Bill Overdue: MTR-UGVCL-8841', 'Bill amount ₹7,850 for AMD-GT-001 is past due date.', 0, 'notif-elec-003', NOW())`);
        console.log('Seeded initial notifications.');
      }

      // Seed Storage Archives if empty
      try {
        const storageCount = await q('SELECT COUNT(*) c FROM storage_archives');
        if (!storageCount[0]?.c) {
          await q(`INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES
            ('ppt', 'Diwali 2026 Prime Sites Pitch Deck', 'MediaBuzz_Automated-PPT_Diwali2026.pptx', '', '4.8 MB', 'pptx', '{"slides":8,"client":"All Prime Sites","sites":["AMD-GT-001","AMD-UP-002","AMD-HD-005"],"orientation":"16:9 Widescreen"}', 'active', DATE_SUB(NOW(), INTERVAL 2 DAY), NOW()),
            ('ppt', 'Rajyash Group - Ahmedabad Outdoor Showcase', 'Rajyash_Group_Outdoor_Showcase.pptx', '', '3.2 MB', 'pptx', '{"slides":5,"client":"Rajyash Group","sites":["AMD-GT-001","AMD-HD-005"]}', 'active', DATE_SUB(NOW(), INTERVAL 5 DAY), NOW()),
            ('ppt', 'Tata Motors EV Launch Presentation', 'Tata_Motors_EV_Launch_Deck.pptx', '', '2.9 MB', 'pptx', '{"slides":6,"client":"Tata Motors EV","sites":["AMD-UP-002"]}', 'active', DATE_SUB(NOW(), INTERVAL 9 DAY), NOW()),
            ('excel', 'Campaign Tracker Master Sheet (20 Columns)', 'MediaBuzz_Campaign_Tracker_Sep2026.xlsx', '', '380 KB', 'xlsx', '{"rows":24,"columns":20,"type":"Campaigns","month":"Sep 2026"}', 'active', DATE_SUB(NOW(), INTERVAL 1 DAY), NOW()),
            ('excel', 'Consolidated Electricity Bills (UGVCL & Torrent)', 'MediaBuzz_Electricity_Bills_Aug2026.xlsx', '', '210 KB', 'xlsx', '{"rows":18,"totalAmount":184500,"type":"Electricity"}', 'active', DATE_SUB(NOW(), INTERVAL 4 DAY), NOW()),
            ('excel', 'Master Sites Portfolio & Geo-Coordinates', 'MediaBuzz_Sites_Catalog_2026.xlsx', '', '512 KB', 'xlsx', '{"rows":22,"type":"Sites Catalog","cities":["Ahmedabad"]}', 'active', DATE_SUB(NOW(), INTERVAL 12 DAY), NOW()),
            ('occupancy', 'September 2026 Site Occupancy Snapshot', 'Occupancy_September_2026.json', '', '14 KB', 'json', '{"month":"September 2026","totalSites":22,"occupiedSites":19,"vacantSites":3,"occupancyPct":86,"revenue":1480000,"activeCampaignsCount":6}', 'active', DATE_SUB(NOW(), INTERVAL 8 DAY), NOW()),
            ('occupancy', 'August 2026 Site Occupancy Snapshot', 'Occupancy_August_2026.json', '', '14 KB', 'json', '{"month":"August 2026","totalSites":22,"occupiedSites":17,"vacantSites":5,"occupancyPct":77,"revenue":1290000,"activeCampaignsCount":5}', 'active', DATE_SUB(NOW(), INTERVAL 38 DAY), NOW()),
            ('occupancy', 'July 2026 Site Occupancy Snapshot', 'Occupancy_July_2026.json', '', '14 KB', 'json', '{"month":"July 2026","totalSites":22,"occupiedSites":16,"vacantSites":6,"occupancyPct":73,"revenue":1150000,"activeCampaignsCount":5}', 'active', DATE_SUB(NOW(), INTERVAL 69 DAY), NOW()),
            ('occupancy', 'June 2026 Site Occupancy Snapshot', 'Occupancy_June_2026.json', '', '14 KB', 'json', '{"month":"June 2026","totalSites":22,"occupiedSites":15,"vacantSites":7,"occupancyPct":68,"revenue":1020000,"activeCampaignsCount":4}', 'active', DATE_SUB(NOW(), INTERVAL 99 DAY), NOW())`);
          console.log('Seeded initial storage archives with generated PPTs, Excels, and historical occupancy snapshots.');
        }
      } catch (storageSeedErr) {
        console.warn('Storage archives seed notice:', storageSeedErr.message);
      }
      // Remove any accidental vacant / blank rows from campaigns table (blank shows in Occupancy only!)
      try {
        await q(`DELETE FROM campaigns WHERE LOWER(TRIM(COALESCE(client, ''))) IN ('blank', 'vacant', 'unassigned', '', '-') OR LOWER(TRIM(COALESCE(client, ''))) LIKE 'blank%' OR LOWER(TRIM(COALESCE(client, ''))) LIKE 'vacant%'`);
      } catch (_) {}

      // Synchronize site availability on startup
      await syncSiteAvailability();
    } catch (seedErr) {
      console.warn('Site seed notice:', seedErr.message);
    }
  } catch (err) {
    console.error('Database connection warning (check DB_HOST, DB_USER, DB_PASSWORD, DB_NAME):', err.message);
  }
}

initDb();

