require('dotenv').config();

const express = require('express');
const cors = require('cors');
const fs = require('fs');
const os = require('os');
const path = require('path');
const multer = require('multer');
const xlsx = require('xlsx');
const mysql = require('mysql2/promise');
const { Pool: PgPool } = require('pg');
const PDFDocument = require('pdfkit');
const QRCode = require('qrcode');
const archiver = require('archiver');
const nodemailer = require('nodemailer');
const crypto = require('crypto');

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || '0.0.0.0';
const ROOT = __dirname;
const FRONTEND_DIR = path.join(ROOT, '..', 'frontend');
const DATA_DIR = process.env.DATA_DIR || ROOT;
const DIRS = {
    templates: path.join(FRONTEND_DIR, 'templates'),
    uploaded: path.join(DATA_DIR, 'uploaded_templates'),
    logos: path.join(DATA_DIR, 'uploaded_logos'),
    uploads: path.join(DATA_DIR, 'uploads'),
    certificates: path.join(DATA_DIR, 'certificates')
};
const LOCAL_ADMIN_FILE = path.join(ROOT, 'admin.local.json');
const CERTIFICATE_CONFIG_FILE = path.join(DATA_DIR, 'certificate-config.json');

Object.values(DIRS).forEach(dir => fs.mkdirSync(dir, { recursive: true }));

if (process.env.TRUST_PROXY === 'true') app.set('trust proxy', 1);

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use((req, res, next) => {
    if (req.path === '/' || req.path.endsWith('.html')) {
        res.set('Cache-Control', 'no-store');
    }
    next();
});
app.use('/templates', express.static(DIRS.templates));
app.use('/uploaded_templates', express.static(DIRS.uploaded));
app.use('/uploaded_logos', express.static(DIRS.logos));
app.use('/certificates', express.static(DIRS.certificates));
app.use(express.static(FRONTEND_DIR));
app.get('/', (_, res) => res.sendFile(path.join(FRONTEND_DIR, 'index.html')));
app.get('/login.html', (_, res) => res.redirect('/index.html'));

const DB_CLIENT = (process.env.DB_CLIENT || (process.env.DATABASE_URL || '').split(':')[0] || 'mysql').toLowerCase();
const isPostgres = ['postgres', 'postgresql'].includes(DB_CLIENT);

function toPostgresQuery(sql) {
    let index = 0;
    return sql.replace(/\?/g, () => `$${++index}`);
}

// These values let the app run locally or on a hosted database without editing code.
function createDb() {
    const sslEnabled = process.env.DB_SSL === 'true' || process.env.MYSQL_SSL === 'true' || isPostgres;

    if (isPostgres) {
        const pool = new PgPool({
            connectionString: process.env.DATABASE_URL,
            ssl: sslEnabled ? { rejectUnauthorized: process.env.DB_SSL_REJECT_UNAUTHORIZED !== 'false' } : undefined,
            max: Number(process.env.DB_CONNECTION_LIMIT) || 10
        });

        return {
            query: async (sql, params = []) => {
                const result = await pool.query(toPostgresQuery(sql), params);
                return [result.rows];
            }
        };
    }

    const common = {
        waitForConnections: true,
        connectionLimit: Number(process.env.DB_CONNECTION_LIMIT) || 10,
        ssl: sslEnabled ? { rejectUnauthorized: process.env.DB_SSL_REJECT_UNAUTHORIZED !== 'false' } : undefined
    };

    const pool = process.env.DATABASE_URL
        ? mysql.createPool({ uri: process.env.DATABASE_URL, ...common })
        : mysql.createPool({
            host: process.env.DB_HOST || 'localhost',
            user: process.env.DB_USER || 'root',
            password: process.env.DB_PASSWORD ?? '',
            database: process.env.DB_NAME || 'certigen',
            port: Number(process.env.DB_PORT) || 3306,
            ...common
        });

    return { query: (sql, params = []) => pool.query(sql, params) };
}

const db = createDb();

const EMAIL_USER = (process.env.EMAIL_USER || '').trim();
const EMAIL_PASS = (process.env.EMAIL_PASS || '').trim();
const transporter = EMAIL_USER && EMAIL_PASS
    ? nodemailer.createTransport({
        host: 'smtp.gmail.com',
        port: 465,
        secure: true,
        auth: {
            user: EMAIL_USER,
            pass: EMAIL_PASS
        },
        tls: {
            rejectUnauthorized: false
        }
    })
    : null;

const hashPassword = password => {
    const salt = crypto.randomBytes(16).toString('hex');
    const derived = crypto.pbkdf2Sync(String(password), salt, 100000, 32, 'sha256').toString('hex');
    return `pbkdf2$${salt}$${derived}`;
};

const verifyPassword = (password, storedPassword) => {
    const stored = String(storedPassword || '');
    if (!stored.startsWith('pbkdf2$')) return String(password) === stored;
    const [, salt, expected] = stored.split('$');
    if (!salt || !expected) return false;
    const derived = crypto.pbkdf2Sync(String(password), salt, 100000, 32, 'sha256').toString('hex');
    if (derived.length !== expected.length) return false;
    return crypto.timingSafeEqual(Buffer.from(derived, 'hex'), Buffer.from(expected, 'hex'));
};

// This creates the required tables when a hosted database starts empty.
async function ensureSchema() {
    if (isPostgres) {
        await db.query(`
            CREATE TABLE IF NOT EXISTS employees (
                id SERIAL PRIMARY KEY,
                name VARCHAR(100) NOT NULL,
                email VARCHAR(150) NOT NULL UNIQUE,
                password VARCHAR(255) NOT NULL,
                phone VARCHAR(30),
                role VARCHAR(30) DEFAULT 'employee'
            )
        `);

        await db.query(`
            CREATE TABLE IF NOT EXISTS certificates (
                id SERIAL PRIMARY KEY,
                name VARCHAR(150) NOT NULL,
                reg_no VARCHAR(100) NOT NULL UNIQUE,
                course_name VARCHAR(150) NOT NULL,
                course_type VARCHAR(80) NOT NULL,
                score VARCHAR(30) NOT NULL,
                passing_marks VARCHAR(30),
                total_marks VARCHAR(30),
                roll_number VARCHAR(100),
                he_she VARCHAR(20),
                start_date DATE NOT NULL,
                end_date DATE NOT NULL,
                email VARCHAR(150),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        `);

        try { await db.query('ALTER TABLE employees ALTER COLUMN password TYPE VARCHAR(255)'); } catch (_) { /* already correct type */ }

        // Auto-migrate: add new columns to existing databases that predate these additions.
        const pgMigrations = [
            "ALTER TABLE certificates ADD COLUMN IF NOT EXISTS email VARCHAR(150)",
            "ALTER TABLE certificates ADD COLUMN IF NOT EXISTS passing_marks VARCHAR(30)",
            "ALTER TABLE certificates ADD COLUMN IF NOT EXISTS total_marks VARCHAR(30)",
            "ALTER TABLE certificates ADD COLUMN IF NOT EXISTS roll_number VARCHAR(100)",
            "ALTER TABLE certificates ADD COLUMN IF NOT EXISTS he_she VARCHAR(20)"
        ];
        for (const sql of pgMigrations) {
            try { await db.query(sql); } catch (_) { /* skip */ }
        }
        return;
    }

    await db.query(`
        CREATE TABLE IF NOT EXISTS employees (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(100) NOT NULL,
            email VARCHAR(150) NOT NULL UNIQUE,
            password VARCHAR(255) NOT NULL,
            phone VARCHAR(30),
            role VARCHAR(30) DEFAULT 'employee'
        )
    `);

    await db.query(`
        CREATE TABLE IF NOT EXISTS certificates (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(150) NOT NULL,
            reg_no VARCHAR(100) NOT NULL UNIQUE,
            course_name VARCHAR(150) NOT NULL,
            course_type VARCHAR(80) NOT NULL,
            score VARCHAR(30) NOT NULL,
            passing_marks VARCHAR(30),
            total_marks VARCHAR(30),
            roll_number VARCHAR(100),
            he_she VARCHAR(20),
            start_date DATE NOT NULL,
            end_date DATE NOT NULL,
            email VARCHAR(150),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    `);

    try { await db.query('ALTER TABLE employees MODIFY password VARCHAR(255) NOT NULL'); } catch (_) { /* already correct type */ }

    // Auto-migrate: add new columns to existing databases that predate these additions.
    const migrations = [
        "ALTER TABLE certificates ADD COLUMN email VARCHAR(150)",
        "ALTER TABLE certificates ADD COLUMN passing_marks VARCHAR(30)",
        "ALTER TABLE certificates ADD COLUMN total_marks VARCHAR(30)",
        "ALTER TABLE certificates ADD COLUMN roll_number VARCHAR(100)",
        "ALTER TABLE certificates ADD COLUMN he_she VARCHAR(20)"
    ];
    for (const sql of migrations) {
        try { await db.query(sql); } catch (_) { /* column already exists, skip */ }
    }
}

async function upsertEmployee({ name, email, password, phone, role }) {
    if (isPostgres) {
        await db.query(
            `INSERT INTO employees (name, email, password, phone, role)
             VALUES (?, ?, ?, ?, ?)
             ON CONFLICT (email) DO UPDATE SET
             name = EXCLUDED.name,
             password = EXCLUDED.password,
             phone = EXCLUDED.phone,
             role = EXCLUDED.role`,
            [name, email, password, phone, role]
        );
        return;
    }

    await db.query(
        `INSERT INTO employees (name, email, password, phone, role)
         VALUES (?, ?, ?, ?, ?)
         ON DUPLICATE KEY UPDATE
         name = VALUES(name),
         password = VALUES(password),
         phone = VALUES(phone),
         role = VALUES(role)`,
        [name, email, password, phone, role]
    );
}

async function upsertCertificate(student) {
    const record = {
        ...student,
        course_type: student.course_type || '',
        score: student.score || '0',
        start_date: student.start_date || '1970-01-01',
        end_date: student.end_date || '1970-01-01'
    };
    if (isPostgres) {
        await db.query(
            `INSERT INTO certificates (name, reg_no, course_name, course_type, score, passing_marks, total_marks, roll_number, he_she, start_date, end_date, email)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
             ON CONFLICT (reg_no) DO UPDATE SET
             name = EXCLUDED.name,
             course_name = EXCLUDED.course_name,
             course_type = EXCLUDED.course_type,
             score = EXCLUDED.score,
             passing_marks = EXCLUDED.passing_marks,
             total_marks = EXCLUDED.total_marks,
             roll_number = EXCLUDED.roll_number,
             he_she = EXCLUDED.he_she,
             start_date = EXCLUDED.start_date,
             end_date = EXCLUDED.end_date,
             email = EXCLUDED.email`,
            [record.name, record.reg_no, record.course_name, record.course_type, record.score, record.passing_marks || null, record.total_marks || null, record.roll_number || null, record.he_she || null, record.start_date, record.end_date, record.email || null]
        );
        return;
    }

    await db.query(
        `INSERT INTO certificates (name, reg_no, course_name, course_type, score, passing_marks, total_marks, roll_number, he_she, start_date, end_date, email)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON DUPLICATE KEY UPDATE name=VALUES(name), course_name=VALUES(course_name), course_type=VALUES(course_type), score=VALUES(score), passing_marks=VALUES(passing_marks), total_marks=VALUES(total_marks), roll_number=VALUES(roll_number), he_she=VALUES(he_she), start_date=VALUES(start_date), end_date=VALUES(end_date), email=VALUES(email)`,
        [record.name, record.reg_no, record.course_name, record.course_type, record.score, record.passing_marks || null, record.total_marks || null, record.roll_number || null, record.he_she || null, record.start_date, record.end_date, record.email || null]
    );
}

// This reads the real admin details from env vars or an ignored local file.
function loadLocalAdmin() {
    if (process.env.ADMIN_EMAIL && process.env.ADMIN_PASSWORD) {
        return {
            name: String(process.env.ADMIN_NAME || 'Admin User').trim(),
            email: String(process.env.ADMIN_EMAIL).trim(),
            password: String(process.env.ADMIN_PASSWORD).trim(),
            phone: String(process.env.ADMIN_PHONE || '').trim(),
            role: 'admin'
        };
    }

    try {
        if (!fs.existsSync(LOCAL_ADMIN_FILE)) return null;
        const raw = JSON.parse(fs.readFileSync(LOCAL_ADMIN_FILE, 'utf8'));
        if (!raw?.email || !raw?.password) return null;
        return {
            name: String(raw.name || 'Local Admin').trim(),
            email: String(raw.email).trim(),
            password: String(raw.password).trim(),
            phone: String(raw.phone || '').trim(),
            role: 'admin'
        };
    } catch (error) {
        console.warn('Local admin file could not be read:', error.message);
        return null;
    }
}

// This creates default login accounts so a fresh setup can sign in right away.
async function ensureDefaultUsers() {
    try {
        const adminEmail = process.env.ADMIN_EMAIL || 'admin@certigen.local';
        const adminPassword = process.env.ADMIN_PASSWORD || 'admin123';
        const employeeEmail = process.env.DEMO_EMPLOYEE_EMAIL || 'employee@certigen.local';
        const employeePassword = process.env.DEMO_EMPLOYEE_PASSWORD || 'employee123';

        await upsertEmployee({
            name: process.env.ADMIN_NAME || 'Admin User',
            email: adminEmail,
            password: hashPassword(adminPassword),
            phone: process.env.ADMIN_PHONE || '0000000000',
            role: 'admin'
        });

        if (process.env.CREATE_DEMO_EMPLOYEE !== 'false') {
            await upsertEmployee({
                name: 'Employee User',
                email: employeeEmail,
                password: hashPassword(employeePassword),
                phone: '1111111111',
                role: 'employee'
            });
        }

        const localAdmin = loadLocalAdmin();
        if (localAdmin) {
            await upsertEmployee({
                name: localAdmin.name,
                email: localAdmin.email,
                password: hashPassword(localAdmin.password),
                phone: localAdmin.phone,
                role: localAdmin.role
            });
        }
    } catch (error) {
        console.warn('Default user setup skipped:', error.message);
    }
}

async function upgradePlainPassword(email, password) {
    await db.query('UPDATE employees SET password=? WHERE email=?', [hashPassword(password), email]);
}

const excelUpload = multer({ dest: DIRS.uploads });
const templateUpload = multer({
    storage: multer.diskStorage({
        destination: (_, __, cb) => cb(null, DIRS.uploaded),
        filename: (_, file, cb) => cb(null, `${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_')}`)
    })
});
const logoUpload = multer({
    storage: multer.diskStorage({
        destination: (_, __, cb) => cb(null, DIRS.logos),
        filename: (_, file, cb) => cb(null, `${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_')}`)
    }),
    fileFilter: (_, file, cb) => {
        const ok = /\.(png|jpe?g)$/i.test(file.originalname || '') && /^image\/(png|jpe?g)$/i.test(file.mimetype || '');
        cb(ok ? null : new Error('Logo must be a PNG, JPG, or JPEG file.'), ok);
    }
});

let selectedTemplate = { name: 'template1.jpg', source: 'built_in', path: path.join(DIRS.templates, 'template1.jpg') };
let uploadedStudents = [];
let uploadSummary = { totalRows: 0, validRows: 0, invalidRows: 0, courseNames: [], invalidParticipants: [] };
let lastGeneration = null;

const CANVAS_SIZE = { width: 1000, height: 707 };
const DEFAULT_PARAGRAPH = 'This certifies that {name} has successfully completed the course {course_name} and fulfilled all requirements.';
const DEFAULT_CERTIFICATE_TITLE = 'Certificate Of Completion';
const DEFAULT_INTRO_TEXT = 'We hereby proudly announce that';
const FONT_FAMILIES = ['Times New Roman', 'Georgia', 'Garamond', 'Arial', 'Helvetica', 'Verdana', 'Open Sans', 'Roboto', 'Inter', 'Playfair Display', 'Cinzel', 'Merriweather', 'Poppins', 'Montserrat', 'Courier'];
const DEFAULT_ELEMENT_ORDER = ['institutionName', 'certificateTitle', 'introText', 'studentName', 'rollNumber', 'customParagraph', 'courseName', 'heShe', 'courseType', 'startDate', 'endDate', 'score', 'signatory1', 'registrationNumber', 'logo', 'qrCode', 'qrLabel'];
const DEFAULT_CERTIFICATE_CONFIG = {
    institutionName: 'CertiGen College',
    certificateTitle: DEFAULT_CERTIFICATE_TITLE,
    introText: DEFAULT_INTRO_TEXT,
    minimumQualifyingScore: 35,
    paragraphTemplate: DEFAULT_PARAGRAPH,
    elements: {
        qrCode: true,
        qrLabel: true,
        institutionName: true,
        certificateTitle: true,
        introText: true,
        registrationNumber: true,
        rollNumber: true,
        heShe: true,
        courseName: true,
        courseType: true,
        startDate: true,
        endDate: true,
        score: true,
        signatory1: true,
        logo: false,
        studentName: true,
        customParagraph: true
    },
    elementOrder: DEFAULT_ELEMENT_ORDER,
    layout: {
        institutionName: { x: 210, y: 52, w: 580, h: 34, fontSize: 20, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: false, color: '#13284b', align: 'center', zIndex: 1, locked: false, deleted: false },
        certificateTitle: { x: 190, y: 93, w: 620, h: 72, fontSize: 34, fontFamily: 'Times New Roman', fontWeight: 'bold', italic: false, underline: false, color: '#13284b', align: 'center', zIndex: 2, locked: false, deleted: false },
        introText: { x: 190, y: 174, w: 620, h: 34, fontSize: 15, fontFamily: 'Helvetica', fontWeight: 'normal', italic: true, underline: false, color: '#536685', align: 'center', zIndex: 3, locked: false, deleted: false },
        studentName: { x: 190, y: 218, w: 620, h: 58, fontSize: 36, fontFamily: 'Times New Roman', fontWeight: 'bold', italic: false, underline: true, color: '#0d1e39', align: 'center', zIndex: 4, locked: false, deleted: false, kind: 'text' },
        rollNumber: { x: 330, y: 276, w: 340, h: 30, fontSize: 15, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: true, color: '#20385f', align: 'center', zIndex: 5, locked: false, deleted: false, kind: 'text' },
        customParagraph: { x: 190, y: 314, w: 620, h: 76, fontSize: 17, fontFamily: 'Helvetica', fontWeight: 'normal', italic: false, underline: false, color: '#536685', align: 'center', zIndex: 6, locked: false, deleted: false, kind: 'text' },
        courseName: { x: 230, y: 399, w: 540, h: 50, fontSize: 25, fontFamily: 'Times New Roman', fontWeight: 'normal', italic: false, underline: true, color: '#13284b', align: 'center', zIndex: 7, locked: false, deleted: false, kind: 'text' },
        heShe: { x: 706, y: 326, w: 82, h: 28, fontSize: 15, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: true, color: '#20385f', align: 'center', zIndex: 8, locked: false, deleted: false, kind: 'text' },
        courseType: { x: 320, y: 451, w: 360, h: 28, fontSize: 14, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: true, color: '#20385f', align: 'center', zIndex: 9, locked: false, deleted: false, kind: 'text' },
        startDate: { x: 288, y: 488, w: 190, h: 26, fontSize: 13, fontFamily: 'Helvetica', fontWeight: 'normal', italic: false, underline: true, color: '#20385f', align: 'center', zIndex: 10, locked: false, deleted: false, kind: 'text' },
        endDate: { x: 522, y: 488, w: 190, h: 26, fontSize: 13, fontFamily: 'Helvetica', fontWeight: 'normal', italic: false, underline: true, color: '#20385f', align: 'center', zIndex: 11, locked: false, deleted: false, kind: 'text' },
        score: { x: 405, y: 524, w: 190, h: 26, fontSize: 13, fontFamily: 'Helvetica', fontWeight: 'normal', italic: false, underline: true, color: '#20385f', align: 'center', zIndex: 12, locked: false, deleted: false, kind: 'text' },
        signatory1: { source: 'signatory', x: 96, y: 618, w: 210, h: 54, fontSize: 12, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: false, color: '#223a62', align: 'center', zIndex: 13, locked: false, deleted: false, kind: 'signatory', name: 'Authorized Signatory' },
        registrationNumber: { x: 400, y: 640, w: 220, h: 28, fontSize: 12, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: true, color: '#223a62', align: 'center', zIndex: 14, locked: false, deleted: false, kind: 'text' },
        logo: { x: 452, y: 42, w: 96, h: 96, fontSize: 12, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: false, color: '#223a62', align: 'center', zIndex: 15, locked: false, deleted: true, kind: 'image', imageUrl: '' },
        qrCode: { x: 804, y: 488, w: 106, h: 106, fontSize: 12, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: false, color: '#223a62', align: 'center', zIndex: 16, locked: false, deleted: false, kind: 'qr' },
        qrLabel: { x: 768, y: 616, w: 178, h: 26, fontSize: 12, fontFamily: 'Helvetica', fontWeight: 'bold', italic: false, underline: false, color: '#223a62', align: 'center', zIndex: 17, locked: false, deleted: false, kind: 'text' }
    },
    presets: [
        { id: 'course-completion', name: 'Course Completion', config: null },
        { id: 'internship-certificate', name: 'Internship Certificate', config: null },
        { id: 'workshop-certificate', name: 'Workshop Certificate', config: null },
        { id: 'participation-certificate', name: 'Participation Certificate', config: null },
        { id: 'appreciation-certificate', name: 'Appreciation Certificate', config: null }
    ]
};

function deepMerge(base, overrides) {
    if (!overrides || typeof overrides !== 'object') return { ...base };
    const merged = Array.isArray(base) ? [...base] : { ...base };
    Object.entries(overrides).forEach(([key, value]) => {
        if (value && typeof value === 'object' && !Array.isArray(value) && base[key] && typeof base[key] === 'object') {
            merged[key] = deepMerge(base[key], value);
        } else {
            merged[key] = value;
        }
    });
    return merged;
}

function loadCertificateConfig() {
    try {
        if (!fs.existsSync(CERTIFICATE_CONFIG_FILE)) return deepMerge(DEFAULT_CERTIFICATE_CONFIG, {});
        return normalizeCertificateConfig(JSON.parse(fs.readFileSync(CERTIFICATE_CONFIG_FILE, 'utf8')));
    } catch (error) {
        console.warn('Certificate config could not be read:', error.message);
        return deepMerge(DEFAULT_CERTIFICATE_CONFIG, {});
    }
}

let certificateConfig = loadCertificateConfig();

// This makes column names easier to match from different Excel files.
const normalize = value => String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
// This keeps PDF file names safe for Windows folders.
const safeName = value => String(value || 'certificate').replace(/[<>:"/\\|?*\x00-\x1F]/g, '').replace(/\s+/g, '_').slice(0, 80) || 'certificate';
const REQUIRED_EXCEL_COLUMN_GROUPS = [
    { label: 'name', aliases: ['name_of_candidate', 'candidate_name', 'student_name', 'name'] },
    { label: 'registration_number', aliases: ['id_number', 'reg_number', 'registration_number', 'reg_no', 'id_no', 'id'] },
    { label: 'course_name', aliases: ['name_of_course', 'course_name', 'course'] },
    { label: 'roll_number', aliases: ['roll_number'] },
    { label: 'he_she', aliases: ['he_she'] }
];

function cleanText(value, fallback = '') {
    const text = String(value ?? '').trim();
    return text || fallback;
}

function cleanNumber(value, fallback, min, max) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return fallback;
    return Math.min(Math.max(numeric, min), max);
}

function normalizeLayoutBox(box, fallback) {
    const safe = fallback || {};
    const source = cleanText(box?.source, safe.source || '');
    const kind = ['text', 'image', 'shape', 'signatory', 'qr'].includes(box?.kind) ? box.kind : (safe.kind || (source === 'signatory' ? 'signatory' : 'text'));
    return {
        source,
        kind,
        x: cleanNumber(box?.x, safe.x || 0, 0, CANVAS_SIZE.width),
        y: cleanNumber(box?.y, safe.y || 0, 0, CANVAS_SIZE.height),
        w: cleanNumber(box?.w, safe.w || 120, 24, CANVAS_SIZE.width),
        h: cleanNumber(box?.h, safe.h || 30, 18, CANVAS_SIZE.height),
        fontSize: cleanNumber(box?.fontSize, safe.fontSize || 16, 8, 96),
        fontFamily: FONT_FAMILIES.includes(box?.fontFamily) ? box.fontFamily : (safe.fontFamily || 'Helvetica'),
        fontWeight: ['normal', 'bold'].includes(box?.fontWeight) ? box.fontWeight : (safe.fontWeight || 'normal'),
        italic: typeof box?.italic === 'boolean' ? box.italic : Boolean(safe.italic),
        underline: typeof box?.underline === 'boolean' ? box.underline : Boolean(safe.underline),
        color: /^#[0-9a-f]{6}$/i.test(String(box?.color || '')) ? box.color : (safe.color || '#13213f'),
        align: ['left', 'center', 'right', 'justify'].includes(box?.align) ? box.align : (safe.align || 'center'),
        letterSpacing: cleanNumber(box?.letterSpacing, safe.letterSpacing || 0, -5, 30),
        opacity: cleanNumber(box?.opacity, safe.opacity ?? 1, 0, 1),
        rotation: cleanNumber(box?.rotation, safe.rotation || 0, -360, 360),
        zIndex: cleanNumber(box?.zIndex, safe.zIndex || 1, 0, 10000),
        locked: typeof box?.locked === 'boolean' ? box.locked : Boolean(safe.locked),
        deleted: typeof box?.deleted === 'boolean' ? box.deleted : Boolean(safe.deleted),
        name: cleanText(box?.name, safe.name || ''),
        text: box?.text !== undefined ? String(box.text) : (safe.text !== undefined ? String(safe.text) : undefined),
        imageUrl: cleanText(box?.imageUrl, safe.imageUrl || ''),
        shape: cleanText(box?.shape, safe.shape || 'rectangle')
    };
}

function normalizeCertificateConfig(input = {}) {
    const merged = deepMerge(DEFAULT_CERTIFICATE_CONFIG, input);
    if (merged.layout?.signature && !merged.layout?.signatory1) {
        merged.layout.signatory1 = {
            ...merged.layout.signature,
            source: 'signatory',
            kind: 'signatory',
            h: Math.max(Number(merged.layout.signature.h) || 28, 54),
            name: merged.layout.signature.name || 'Authorized Signatory'
        };
        merged.elements.signatory1 = merged.elements.signature !== false;
        merged.elementOrder = (merged.elementOrder || DEFAULT_ELEMENT_ORDER).map(key => key === 'signature' ? 'signatory1' : key);
    }
    const elements = { ...DEFAULT_CERTIFICATE_CONFIG.elements, ...(merged.elements || {}) };
    elements.institutionName = true;
    elements.certificateTitle = true;
    elements.courseName = true;
    elements.studentName = true;
    elements.customParagraph = true;
    elements.introText = true;
    elements.qrLabel = Boolean(elements.qrCode);
    Object.keys(merged.layout || {}).forEach(key => {
        if (merged.layout[key]?.source === 'signatory' || merged.layout[key]?.kind === 'signatory') elements[key] = merged.elements?.[key] !== false;
        if (merged.layout[key]?.kind === 'image') elements[key] = merged.elements?.[key] !== false;
    });

    const layout = {};
    Object.entries(DEFAULT_CERTIFICATE_CONFIG.layout).forEach(([key, fallback]) => {
        layout[key] = normalizeLayoutBox(merged.layout?.[key], fallback);
        if (layout[key].x + layout[key].w > CANVAS_SIZE.width) layout[key].x = CANVAS_SIZE.width - layout[key].w;
        if (layout[key].y + layout[key].h > CANVAS_SIZE.height) layout[key].y = CANVAS_SIZE.height - layout[key].h;
    });
    Object.entries(merged.layout || {}).forEach(([key, box]) => {
        if (key === 'signature') return;
        if (layout[key] || !box || typeof box !== 'object') return;
        const source = box.source && (DEFAULT_CERTIFICATE_CONFIG.layout[box.source] || box.source === 'signatory') ? box.source : 'customParagraph';
        const fallback = source === 'signatory' ? DEFAULT_CERTIFICATE_CONFIG.layout.signatory1 : DEFAULT_CERTIFICATE_CONFIG.layout[source];
        layout[key] = normalizeLayoutBox({ ...box, source }, fallback);
        if (layout[key].x + layout[key].w > CANVAS_SIZE.width) layout[key].x = CANVAS_SIZE.width - layout[key].w;
        if (layout[key].y + layout[key].h > CANVAS_SIZE.height) layout[key].y = CANVAS_SIZE.height - layout[key].h;
    });

    const elementOrder = Array.isArray(merged.elementOrder)
        ? [...new Set([...merged.elementOrder, ...DEFAULT_ELEMENT_ORDER, ...Object.keys(layout)])].filter(key => layout[key])
        : [...DEFAULT_ELEMENT_ORDER, ...Object.keys(layout).filter(key => !DEFAULT_ELEMENT_ORDER.includes(key))];

    const presets = Array.isArray(merged.presets)
        ? merged.presets.map((preset, index) => ({
            id: cleanText(preset.id, `preset-${index + 1}`).replace(/[^a-zA-Z0-9_-]/g, '-').toLowerCase(),
            name: cleanText(preset.name, `Preset ${index + 1}`),
            config: preset.config && typeof preset.config === 'object' ? normalizeCertificateConfig({ ...preset.config, presets: [] }) : null
        })).slice(0, 20)
        : DEFAULT_CERTIFICATE_CONFIG.presets;

    return {
        institutionName: cleanText(merged.institutionName, DEFAULT_CERTIFICATE_CONFIG.institutionName),
        certificateTitle: cleanText(merged.certificateTitle, DEFAULT_CERTIFICATE_TITLE),
        introText: cleanText(merged.introText, DEFAULT_INTRO_TEXT),
        minimumQualifyingScore: cleanNumber(merged.minimumQualifyingScore, DEFAULT_CERTIFICATE_CONFIG.minimumQualifyingScore, 0, 100000),
        paragraphTemplate: cleanText(merged.paragraphTemplate, DEFAULT_PARAGRAPH),
        elements,
        elementOrder,
        layout,
        presets
    };
}

function saveCertificateConfig(nextConfig) {
    certificateConfig = normalizeCertificateConfig(nextConfig);
    fs.writeFileSync(CERTIFICATE_CONFIG_FILE, JSON.stringify(certificateConfig, null, 2));
    return certificateConfig;
}

function resetCertificateConfig() {
    certificateConfig = normalizeCertificateConfig(DEFAULT_CERTIFICATE_CONFIG);
    fs.writeFileSync(CERTIFICATE_CONFIG_FILE, JSON.stringify(certificateConfig, null, 2));
    return certificateConfig;
}

function cloneConfigForPreset(config) {
    const copy = normalizeCertificateConfig(config);
    return {
        institutionName: copy.institutionName,
        certificateTitle: copy.certificateTitle,
        introText: copy.introText,
        minimumQualifyingScore: copy.minimumQualifyingScore,
        paragraphTemplate: copy.paragraphTemplate,
        elements: copy.elements,
        elementOrder: copy.elementOrder,
        layout: copy.layout
    };
}

// This finds the local network IP so QR links can work on other devices too.
function getLanIp() {
    const nets = os.networkInterfaces();
    for (const entries of Object.values(nets)) {
        for (const entry of entries || []) {
            if (entry.family === 'IPv4' && !entry.internal) return entry.address;
        }
    }
    return null;
}

// This builds the base URL used inside QR verification links.
function baseUrl(req) {
    if (process.env.PUBLIC_BASE_URL) return process.env.PUBLIC_BASE_URL;
    const host = req.get('host') || `localhost:${PORT}`;
    if (host.includes('localhost') || host.includes('127.0.0.1')) {
        const lanIp = getLanIp();
        if (lanIp) return `${req.protocol}://${lanIp}:${PORT}`;
    }
    return `${req.protocol}://${host}`;
}

// This turns different date formats into one simple YYYY-MM-DD format.
const formatDate = value => {
    if (value === undefined || value === null || value === '') return '';
    if (value instanceof Date && !Number.isNaN(value)) return value.toISOString().split('T')[0];
    if (typeof value === 'number') return new Date((value - 25569) * 86400 * 1000).toISOString().split('T')[0];
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? String(value).trim() : parsed.toISOString().split('T')[0];
};

// This picks the first filled value from a list of possible column names.
const pick = (row, keys) => keys.map(key => row[key]).find(v => v !== undefined && v !== null && String(v).trim() !== '') || '';

// This helps show why a participant did not get a certificate.
// When the Excel sheet has a passing_marks column, that value is used per-row.
// If passing_marks is absent, the fallback threshold of 35 is used.
function invalidReason(student) {
    const requiredCore = ['name', 'reg_no', 'course_name', 'roll_number', 'he_she'];
    const missingFields = requiredCore
        .filter(key => !String(student[key] || '').trim())
        .map(key => key.replace(/_/g, ' '));

    if (missingFields.length) {
        return `Missing: ${missingFields.join(', ')}`;
    }

    if (student.score && Number.isNaN(Number(student.score))) return 'Score is not a valid number';
    return 'Row could not be processed';
}

// This combines built-in and uploaded templates for the UI.
const templateList = () => [
    ...fs.readdirSync(DIRS.templates).filter(f => /\.(png|jpe?g)$/i.test(f)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).map(name => ({
        name, source: 'built_in', url: `/templates/${encodeURIComponent(name)}`, selected: selectedTemplate.name === name && selectedTemplate.source === 'built_in'
    })),
    ...fs.readdirSync(DIRS.uploaded).filter(f => /\.(png|jpe?g)$/i.test(f)).sort((a, b) => b.localeCompare(a)).map(name => ({
        name, source: 'uploaded', url: `/uploaded_templates/${encodeURIComponent(name)}`, selected: selectedTemplate.name === name && selectedTemplate.source === 'uploaded'
    }))
];

// This reads Excel rows and keeps students with enough data to generate.
// Score eligibility is applied later only when the Score element is enabled.
function mapStudents(rows) {
    const validRows = [];
    const courseNames = new Set();
    const invalidParticipants = [];
    let invalidRows = 0;

    rows.forEach((raw, index) => {
        const row = {};
        Object.entries(raw).forEach(([key, value]) => { row[normalize(key)] = value; });

        // Read the optional passing_marks and total_marks columns — they may not exist in every sheet.
        const rawPassing = String(pick(row, ['passing_marks', 'pass_marks', 'minimum_marks', 'min_marks', 'passing_score', 'pass_score'])).trim();
        const rawTotal   = String(pick(row, ['total_marks', 'maximum_marks', 'max_marks', 'total_score', 'out_of'])).trim();

        const student = {
            name:         String(pick(row, ['name_of_candidate', 'candidate_name', 'student_name', 'name'])).trim(),
            reg_no:       String(pick(row, ['id_number', 'reg_number', 'registration_number', 'reg_no', 'id_no', 'id'])).trim(),
            roll_number:  String(pick(row, ['roll_number', 'roll_no', 'student_roll_number', 'roll'])).trim(),
            he_she:       String(pick(row, ['he_she', 'pronoun', 'gender_pronoun'])).trim(),
            course_name:  String(pick(row, ['name_of_course', 'course_name', 'course'])).trim(),
            start_date:   formatDate(pick(row, ['starting_date_of_the_course', 'start_date', 'course_start_date', 'starting_date'])),
            end_date:     formatDate(pick(row, ['ending_date_of_the_course', 'end_date', 'course_end_date', 'ending_date'])),
            score:        String(pick(row, ['score_of_the_candidate_in_the_course', 'score_achieved', 'score', 'marks'])).trim(),
            email:        String(pick(row, ['email_of_the_participants', 'participant_email', 'email'])).trim(),
            course_type:  String(pick(row, ['course_type_online_offline', 'course_type', 'type_of_course'])).trim(),
            passing_marks: rawPassing,
            total_marks:   rawTotal
        };

        const scoreValue = Number(student.score);

        // Only the identity and mandatory course fields are required at upload time.
        const requiredCore = ['name', 'reg_no', 'course_name', 'roll_number', 'he_she'];
        const hasRequiredFields = requiredCore.every(key => Boolean(String(student[key] || '').trim()));
        const hasValidScore = !student.score || !Number.isNaN(scoreValue);

        if (hasRequiredFields && hasValidScore) {
            validRows.push(student);
            courseNames.add(student.course_name);
        } else {
            invalidRows += 1;
            invalidParticipants.push({
                rowNumber:    index + 2,
                name:         student.name         || 'No name',
                reg_no:       student.reg_no        || 'No registration number',
                roll_number:  student.roll_number   || 'No roll number',
                he_she:       student.he_she        || 'No he/she value',
                course_name:  student.course_name   || 'No course',
                email:        student.email         || 'No email',
                score:        student.score         || 'No score',
                passing_marks: student.passing_marks || '',
                total_marks:  student.total_marks   || '',
                course_type:  student.course_type   || 'No course type',
                reason:       invalidReason(student)
            });
        }
    });

    return {
        students: validRows,
        summary: {
            totalRows: rows.length,
            validRows: validRows.length,
            invalidRows,
            courseNames: [...courseNames],
            invalidParticipants
        }
    };
}

// This sends a mail update after certificates are generated.
async function sendGenerationMail({ employeeEmail, zipName, zipPath, zipUrl, generatedCount, invalidParticipants }) {
    if (!transporter) return { sent: false, reason: 'Email is not configured on the server.' };
    try {
        const [admins] = await db.query(
            "SELECT id, name, email, phone, role FROM employees WHERE role='admin' ORDER BY id DESC LIMIT 1"
        );
        const [employeeRows] = employeeEmail
            ? await db.query('SELECT id,name,email,phone,role FROM employees WHERE email=? LIMIT 1', [employeeEmail])
            : [[]];
        const employee = employeeRows[0] || null;
        const latestAdmin = admins[0] || null;
        const recipients = [latestAdmin?.email].filter(Boolean);
        if (!recipients.length) return { sent: false, reason: 'No admin email address was found.' };

        const invalidLines = (invalidParticipants || []).length
            ? invalidParticipants.map(person => {
                const passingInfo = person.passing_marks
                    ? ` | Passing Marks: ${person.passing_marks}${person.total_marks ? ` / ${person.total_marks}` : ''}`
                    : '';
                return `Row ${person.rowNumber}: ${person.name} | Reg No: ${person.reg_no} | Roll No: ${person.roll_number || ''} | He/She: ${person.he_she || ''} | Course: ${person.course_name} | Email: ${person.email} | Score: ${person.score}${passingInfo} | Reason: ${person.reason}`;
            }).join('\n')
            : 'No ineligible participants in this upload.';

        await transporter.sendMail({
            from: EMAIL_USER,
            to: recipients.join(','),
            subject: 'Certificates generated successfully',
            text: [
                'Certificate generation summary',
                '',
                `ZIP file: ${zipName}`,
                `ZIP download: ${zipUrl}`,
                `Certificates created: ${generatedCount}`,
                `Ineligible participants: ${(invalidParticipants || []).length}`,
                `Template used: ${selectedTemplate.name}`,
                '',
                'Latest admin details',
                `Name: ${latestAdmin?.name || 'Unknown admin'}`,
                `Email: ${latestAdmin?.email || 'Not available'}`,
                `Phone: ${latestAdmin?.phone || 'Not available'}`,
                '',
                'Employee details',
                `Name: ${employee?.name || 'Unknown employee'}`,
                `Email: ${employee?.email || employeeEmail || 'Not provided'}`,
                `Phone: ${employee?.phone || 'Not available'}`,
                `Role: ${employee?.role || 'employee'}`,
                '',
                'Ineligible participant details',
                invalidLines
            ].join('\n'),
            attachments: [{
                filename: zipName,
                path: zipPath
            }]
        });
        return { sent: true, reason: `Email sent to: ${recipients.join(', ')}` };
    } catch {
        console.warn('Email notification skipped. USER:', EMAIL_USER, 'PASS length:', (EMAIL_PASS||'').length);
        return { sent: false, reason: 'Email sending failed. Check EMAIL_USER, EMAIL_PASS, and Gmail app password.' };
    }
}

function studentPlaceholders(student, config) {
    return {
        NAME: student.name || '',
        name: student.name || '',
        COURSE_NAME: student.course_name || '',
        course_name: student.course_name || '',
        COURSE_TYPE: student.course_type || '',
        course_type: student.course_type || '',
        START_DATE: student.start_date || '',
        start_date: student.start_date || '',
        END_DATE: student.end_date || '',
        end_date: student.end_date || '',
        SCORE: student.score || '',
        score: student.score || '',
        REGISTRATION_NUMBER: student.reg_no || '',
        registration_number: student.reg_no || '',
        ROLL_NUMBER: student.roll_number || '',
        roll_number: student.roll_number || '',
        HE_SHE: student.he_she || '',
        he_she: student.he_she || '',
        INSTITUTION_NAME: config.institutionName || '',
        institution_name: config.institutionName || ''
    };
}

function missingExcelColumns(rows) {
    const headers = new Set();
    rows.forEach(row => Object.keys(row || {}).forEach(key => headers.add(normalize(key))));
    return REQUIRED_EXCEL_COLUMN_GROUPS
        .filter(group => !group.aliases.some(alias => headers.has(alias)))
        .map(group => group.label);
}

function renderTemplateText(template, student, config) {
    const values = studentPlaceholders(student, config);
    return String(template || '').replace(/\{([a-zA-Z_]+)\}/g, (_, key) => values[key] ?? values[key.toUpperCase()] ?? values[key.toLowerCase()] ?? '');
}

function pdfFontName(box) {
    const serifFonts = ['Times New Roman', 'Georgia', 'Garamond', 'Playfair Display', 'Cinzel', 'Merriweather'];
    const family = serifFonts.includes(box.fontFamily) ? 'Times' : box.fontFamily === 'Courier' ? 'Courier' : 'Helvetica';
    if (family === 'Times') {
        if (box.fontWeight === 'bold' && box.italic) return 'Times-BoldItalic';
        if (box.fontWeight === 'bold') return 'Times-Bold';
        if (box.italic) return 'Times-Italic';
        return 'Times-Roman';
    }
    if (family === 'Courier') {
        if (box.fontWeight === 'bold' && box.italic) return 'Courier-BoldOblique';
        if (box.fontWeight === 'bold') return 'Courier-Bold';
        if (box.italic) return 'Courier-Oblique';
        return 'Courier';
    }
    if (box.fontWeight === 'bold' && box.italic) return 'Helvetica-BoldOblique';
    if (box.fontWeight === 'bold') return 'Helvetica-Bold';
    if (box.italic) return 'Helvetica-Oblique';
    return 'Helvetica';
}

function scaledBox(doc, box) {
    const sx = doc.page.width / CANVAS_SIZE.width;
    const sy = doc.page.height / CANVAS_SIZE.height;
    return {
        x: box.x * sx,
        y: box.y * sy,
        w: box.w * sx,
        h: box.h * sy,
        fontSize: box.fontSize * Math.min(sx, sy)
    };
}

function drawFittedText(doc, text, box) {
    const b = scaledBox(doc, box);
    let fontSize = b.fontSize;
    const minSize = 7;
    doc.save();
    doc.rotate(Number(box.rotation) || 0, { origin: [b.x + b.w / 2, b.y + b.h / 2] });
    doc.opacity(box.opacity ?? 1);
    doc.font(pdfFontName(box));
    while (fontSize > minSize) {
        doc.fontSize(fontSize);
        const height = doc.heightOfString(text, { width: b.w, align: box.align, characterSpacing: Number(box.letterSpacing) || 0 });
        if (height <= b.h) break;
        fontSize -= 1;
    }
    doc.fillColor(box.color).fontSize(fontSize).text(text, b.x, b.y, {
        width: b.w,
        height: b.h,
        align: box.align,
        characterSpacing: Number(box.letterSpacing) || 0,
        underline: Boolean(box.underline),
        ellipsis: true
    });
    doc.restore();
}

function drawSignatory(doc, box) {
    const b = scaledBox(doc, box);
    doc.save();
    doc.rotate(Number(box.rotation) || 0, { origin: [b.x + b.w / 2, b.y + b.h / 2] });
    doc.opacity(box.opacity ?? 1);
    doc.strokeColor(box.color || '#223a62').lineWidth(1);
    const lineY = b.y + Math.min(b.h * 0.35, 22);
    doc.moveTo(b.x + b.w * 0.08, lineY).lineTo(b.x + b.w * 0.92, lineY).stroke();
    doc.restore();
    drawFittedText(doc, box.name || 'Authorized Signatory', { ...box, y: box.y + box.h * 0.42, h: box.h * 0.5 });
}

function imagePathFromUrl(imageUrl) {
    const value = String(imageUrl || '');
    if (!value.startsWith('/uploaded_logos/')) return null;
    const filePath = path.join(DIRS.logos, path.basename(decodeURIComponent(value)));
    return fs.existsSync(filePath) ? filePath : null;
}

function drawImageElement(doc, box) {
    const filePath = imagePathFromUrl(box.imageUrl);
    if (!filePath) return;
    const b = scaledBox(doc, box);
    doc.save();
    doc.rotate(Number(box.rotation) || 0, { origin: [b.x + b.w / 2, b.y + b.h / 2] });
    doc.opacity(box.opacity ?? 1);
    doc.image(filePath, b.x, b.y, { fit: [b.w, b.h], align: 'center', valign: 'center' });
    doc.restore();
}

function drawShapeElement(doc, box) {
    const b = scaledBox(doc, box);
    doc.save();
    doc.rotate(Number(box.rotation) || 0, { origin: [b.x + b.w / 2, b.y + b.h / 2] });
    doc.opacity(box.opacity ?? 1);
    doc.fillColor(box.color || '#20385f').rect(b.x, b.y, b.w, b.h).fill();
    doc.restore();
}

function elementText(key, student, config) {
    const box = config.layout?.[key] || {};
    const sourceKey = box.source || key;
    const values = {
        institutionName: config.institutionName,
        certificateTitle: config.certificateTitle,
        introText: config.introText,
        studentName: student.name,
        rollNumber: student.roll_number ? `Roll No: ${student.roll_number}` : '',
        customParagraph: renderTemplateText(config.paragraphTemplate, student, config),
        courseName: student.course_name,
        heShe: student.he_she || '',
        courseType: student.course_type ? `Course Type: ${student.course_type}` : '',
        registrationNumber: student.reg_no ? `Reg. No: ${student.reg_no}` : '',
        startDate: student.start_date ? `Start Date: ${student.start_date}` : '',
        endDate: student.end_date ? `End Date: ${student.end_date}` : '',
        score: student.score ? `Score Achieved: ${student.score}` : '',
        signatory: config.layout?.[key]?.name || '',
        qrLabel: 'QR Verification'
    };
    return values[sourceKey] || values[key] || renderTemplateText(box.text || '', student, config);
}

function certificateElementsForGeneration(config) {
    return (config.elementOrder || Object.keys(config.layout))
        .filter(key => config.layout[key])
        .sort((a, b) => (config.layout[a].zIndex || 0) - (config.layout[b].zIndex || 0))
        .filter(key => {
            if (config.layout[key].deleted) return false;
            const sourceKey = config.layout[key].source || key;
            if (sourceKey === 'institutionName' || sourceKey === 'courseName' || sourceKey === 'certificateTitle' || sourceKey === 'introText' || sourceKey === 'studentName' || sourceKey === 'customParagraph' || sourceKey === 'rollNumber' || sourceKey === 'heShe' || sourceKey === 'signatory') return true;
            return Boolean(config.elements[sourceKey]);
        });
}

function isScoreEligible(student, config) {
    if (!config.elements.score) return true;
    const score = Number(student.score);
    return Number.isFinite(score) && score >= Number(config.minimumQualifyingScore);
}

function eligibilityReason(student, config) {
    if (!student.score) return 'Score is required because Score Achieved is enabled.';
    if (Number.isNaN(Number(student.score))) return 'Score is not a valid number.';
    return `Score ${student.score} is below the qualifying score of ${config.minimumQualifyingScore}.`;
}

// This draws one full certificate page using the saved layout and optional QR verification.
async function drawCertificate(doc, student, verifyUrl, config = certificateConfig) {
    const w = doc.page.width;
    const h = doc.page.height;
    doc.image(selectedTemplate.path, 0, 0, { width: w, height: doc.page.height });

    for (const key of certificateElementsForGeneration(config)) {
        if (key === 'qrCode') continue;
        const box = config.layout[key];
        if (box.kind === 'image') {
            drawImageElement(doc, box);
            continue;
        }
        if (box.kind === 'shape') {
            drawShapeElement(doc, box);
            continue;
        }
        if ((box.source || key) === 'signatory' || box.kind === 'signatory') {
            drawSignatory(doc, box);
            continue;
        }
        const text = elementText(key, student, config);
        if (text) drawFittedText(doc, text, box);
    }

    if (config.elements.qrCode && !config.layout.qrCode.deleted) {
        const box = scaledBox(doc, config.layout.qrCode);
        const qrSize = Math.min(box.w, box.h);
        const qr = await QRCode.toBuffer(verifyUrl, { margin: 1, width: Math.max(80, Math.round(qrSize)) });
        doc.image(qr, box.x, box.y, { width: qrSize, height: qrSize });
    }
}

// This checks login details and sends back the employee role.
app.post('/login', async (req, res) => {
    try {
        const [rows] = await db.query('SELECT id,name,email,password,role FROM employees WHERE email=? LIMIT 1', [req.body.email]);
        const user = rows[0];
        if (!user || !verifyPassword(req.body.password, user.password)) return res.json({ success: false });
        if (!String(user.password || '').startsWith('pbkdf2$')) await upgradePlainPassword(user.email, req.body.password);
        const { password: _password, ...employee } = user;
        res.json({ success: true, role: employee.role || 'employee', employee });
    } catch {
        res.status(500).json({ success: false, message: 'Login failed' });
    }
});

app.get('/healthz', async (_, res) => {
    try {
        await db.query('SELECT 1');
        res.json({ ok: true });
    } catch {
        res.status(503).json({ ok: false });
    }
});

// This gives the frontend the full template list.
app.get('/templates-list', (_, res) => res.json(templateList()));

// This saves the template chosen by the employee.
app.post('/select-template', (req, res) => {
    const dir = req.body.source === 'uploaded' ? DIRS.uploaded : DIRS.templates;
    const filePath = path.join(dir, req.body.template || '');
    if (!req.body.template || !fs.existsSync(filePath)) return res.status(400).json({ success: false, message: 'Template not found' });
    selectedTemplate = { name: req.body.template, source: req.body.source === 'uploaded' ? 'uploaded' : 'built_in', path: filePath };
    res.json({ success: true, selectedTemplate });
});

// This uploads a custom template image and selects it right away.
app.post('/upload-template', templateUpload.single('template'), (req, res) => {
    if (!req.file) return res.status(400).json({ success: false, message: 'Template file is required' });
    selectedTemplate = { name: req.file.filename, source: 'uploaded', path: req.file.path };
    res.json({ success: true, template: { name: req.file.filename, source: 'uploaded', url: `/uploaded_templates/${encodeURIComponent(req.file.filename)}` } });
});

app.post('/upload-logo', (req, res) => {
    logoUpload.single('logo')(req, res, error => {
        if (error) return res.status(400).json({ success: false, message: error.message || 'Logo upload failed.' });
        if (!req.file) return res.status(400).json({ success: false, message: 'Logo file is required.' });
        res.json({ success: true, logo: { name: req.file.filename, url: `/uploaded_logos/${encodeURIComponent(req.file.filename)}` } });
    });
});

app.get('/certificate-config', (_, res) => {
    res.json({ success: true, config: certificateConfig, canvasSize: CANVAS_SIZE });
});

app.post('/certificate-config', (req, res) => {
    try {
        const saved = saveCertificateConfig(req.body || {});
        res.json({ success: true, config: saved, canvasSize: CANVAS_SIZE });
    } catch (error) {
        console.error('Certificate config save failed:', error?.message || error);
        res.status(400).json({ success: false, message: 'Certificate configuration could not be saved.' });
    }
});

app.post('/certificate-config/reset', (_, res) => {
    try {
        const saved = resetCertificateConfig();
        res.json({ success: true, config: saved, canvasSize: CANVAS_SIZE, message: 'Default layout restored.' });
    } catch (error) {
        console.error('Certificate config reset failed:', error?.message || error);
        res.status(500).json({ success: false, message: 'Certificate layout could not be reset.' });
    }
});

app.get('/certificate-presets', (_, res) => {
    res.json({ success: true, presets: certificateConfig.presets || [] });
});

app.post('/certificate-presets', (req, res) => {
    try {
        const name = cleanText(req.body?.name, 'Custom Preset');
        const id = `${name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'preset'}-${Date.now()}`;
        const next = deepMerge(certificateConfig, {});
        next.presets = [...(next.presets || []), { id, name, config: cloneConfigForPreset(req.body?.config || certificateConfig) }].slice(-20);
        const saved = saveCertificateConfig(next);
        res.json({ success: true, presets: saved.presets, preset: saved.presets.find(preset => preset.id === id) });
    } catch (error) {
        console.error('Preset save failed:', error?.message || error);
        res.status(400).json({ success: false, message: 'Preset could not be saved.' });
    }
});

app.post('/certificate-presets/:id/load', (req, res) => {
    try {
        const preset = (certificateConfig.presets || []).find(item => item.id === req.params.id);
        if (!preset?.config) return res.status(404).json({ success: false, message: 'Preset does not have a saved layout yet.' });
        const saved = saveCertificateConfig({ ...preset.config, presets: certificateConfig.presets });
        res.json({ success: true, config: saved, canvasSize: CANVAS_SIZE });
    } catch (error) {
        console.error('Preset load failed:', error?.message || error);
        res.status(400).json({ success: false, message: 'Preset could not be loaded.' });
    }
});

app.post('/certificate-presets/:id/duplicate', (req, res) => {
    try {
        const preset = (certificateConfig.presets || []).find(item => item.id === req.params.id);
        if (!preset) return res.status(404).json({ success: false, message: 'Preset not found.' });
        const id = `${preset.id}-copy-${Date.now()}`;
        const next = deepMerge(certificateConfig, {});
        next.presets = [...(next.presets || []), { id, name: `${preset.name} Copy`, config: preset.config ? cloneConfigForPreset(preset.config) : null }].slice(-20);
        const saved = saveCertificateConfig(next);
        res.json({ success: true, presets: saved.presets });
    } catch (error) {
        console.error('Preset duplicate failed:', error?.message || error);
        res.status(400).json({ success: false, message: 'Preset could not be duplicated.' });
    }
});

app.delete('/certificate-presets/:id', (req, res) => {
    try {
        const next = deepMerge(certificateConfig, {});
        next.presets = (next.presets || []).filter(item => item.id !== req.params.id);
        const saved = saveCertificateConfig(next);
        res.json({ success: true, presets: saved.presets });
    } catch (error) {
        console.error('Preset delete failed:', error?.message || error);
        res.status(400).json({ success: false, message: 'Preset could not be deleted.' });
    }
});

app.get('/sample-excel', (_, res) => {
    const worksheet = xlsx.utils.json_to_sheet([{
        Name: 'Lalith Kartheek',
        Registration_Number: 'REG-001',
        roll_number: '22A91A0501',
        he_she: 'He',
        Course_Name: 'Data Structures and Algorithms',
        Course_Type: 'Online',
        Start_Date: '2026-01-01',
        End_Date: '2026-03-31',
        Score: 86,
        Email: 'student@example.com'
    }], { header: ['Name', 'Registration_Number', 'roll_number', 'he_she', 'Course_Name', 'Course_Type', 'Start_Date', 'End_Date', 'Score', 'Email'] });
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Participants');
    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="CertiGen_Sample_Participants.xlsx"');
    res.send(buffer);
});

// This uploads the Excel file, reads the first sheet, and keeps valid rows.
app.post('/upload-excel', excelUpload.single('file'), (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Excel file is required' });
        const workbook = xlsx.readFile(req.file.path, { cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });
        const missingColumns = missingExcelColumns(rows);
        if (missingColumns.length) {
            return res.status(400).json({
                success: false,
                message: `Missing required Excel column${missingColumns.length > 1 ? 's' : ''}: ${missingColumns.join(', ')}. Download the latest sample Excel file and upload again.`
            });
        }
        const mapped = mapStudents(rows);
        uploadedStudents = mapped.students;
        uploadSummary = mapped.summary;
        res.json({ success: true, total: uploadSummary.totalRows, valid: uploadSummary.validRows, invalid: uploadSummary.invalidRows, courseNames: uploadSummary.courseNames });
    } catch (error) {
        console.error('Excel upload failed:', error?.message || error);
        res.status(500).json({ success: false, message: 'Failed to process Excel file' });
    } finally {
        if (req.file?.path && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    }
});

// This creates all certificate PDFs, zips them, and stores the latest batch info.
app.post('/generate', async (req, res) => {
    try {
        if (!uploadedStudents.length) return res.status(400).json({ success: false, message: 'Upload a valid Excel sheet first' });
        const activeConfig = req.body?.config ? saveCertificateConfig(req.body.config) : certificateConfig;
        const eligibleStudents = [];
        const ineligibleParticipants = [];

        uploadedStudents.forEach((student, index) => {
            if (isScoreEligible(student, activeConfig)) {
                eligibleStudents.push(student);
            } else {
                ineligibleParticipants.push({
                    rowNumber: index + 2,
                    name: student.name || 'No name',
                    reg_no: student.reg_no || 'No registration number',
                    roll_number: student.roll_number || 'No roll number',
                    he_she: student.he_she || 'No he/she value',
                    course_name: student.course_name || 'No course',
                    email: student.email || 'No email',
                    score: student.score || 'No score',
                    passing_marks: String(activeConfig.minimumQualifyingScore),
                    total_marks: student.total_marks || '',
                    course_type: student.course_type || 'No course type',
                    reason: eligibilityReason(student, activeConfig)
                });
            }
        });

        if (!eligibleStudents.length) {
            return res.status(400).json({
                success: false,
                message: 'No participants met the current certificate eligibility settings.',
                eligibleCount: 0,
                ineligibleCount: ineligibleParticipants.length,
                generatedCount: 0
            });
        }

        const zipName = `certificates_${Date.now()}.zip`;
        const zipPath = path.join(DIRS.certificates, zipName);
        const batchDir = path.join(DIRS.certificates, `batch_${Date.now()}`);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', { zlib: { level: 9 } });
        fs.mkdirSync(batchDir, { recursive: true });
        archive.pipe(output);

        for (const student of eligibleStudents) {
            await upsertCertificate(student);

            const pdfName = `${safeName(student.name)}_${safeName(student.reg_no)}.pdf`;
            const pdfPath = path.join(batchDir, pdfName);
            const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 0 });
            const pdfStream = fs.createWriteStream(pdfPath);
            doc.pipe(pdfStream);
            await new Promise((resolve, reject) => {
                pdfStream.on('finish', resolve);
                pdfStream.on('error', reject);
                doc.on('error', reject);
                drawCertificate(doc, student, `${baseUrl(req)}/verify/${encodeURIComponent(student.reg_no)}`, activeConfig).then(() => doc.end()).catch(reject);
            });
            archive.file(pdfPath, { name: pdfName });
        }

        await archive.finalize();
        await new Promise((resolve, reject) => { output.on('close', resolve); output.on('error', reject); });
        fs.rmSync(batchDir, { recursive: true, force: true });
        lastGeneration = {
            zipName,
            zipUrl: `/download-zip/${encodeURIComponent(zipName)}`,
            createdAt: new Date().toISOString(),
            generatedCount: eligibleStudents.length,
            eligibleCount: eligibleStudents.length,
            ineligibleCount: ineligibleParticipants.length,
            failedCount: uploadSummary.invalidRows + ineligibleParticipants.length
        };
        const emailStatus = await sendGenerationMail({
            employeeEmail: req.body.employeeEmail,
            zipName,
            zipPath,
            zipUrl: `${baseUrl(req)}${lastGeneration.zipUrl}`,
            generatedCount: eligibleStudents.length,
            invalidParticipants: [...uploadSummary.invalidParticipants, ...ineligibleParticipants]
        });
        res.json({
            success: true,
            eligibleCount: eligibleStudents.length,
            ineligibleCount: ineligibleParticipants.length,
            generatedCount: eligibleStudents.length,
            failedCount: uploadSummary.invalidRows + ineligibleParticipants.length,
            zipUrl: lastGeneration.zipUrl,
            emailSent: Boolean(emailStatus?.sent),
            emailMessage: emailStatus?.reason || 'Email status unavailable.'
        });
    } catch (error) {
        console.error('Generation failed:', error?.message || error);
        console.error('Stack:', error?.stack);
        res.status(500).json({ success: false, message: `Failed to generate certificates: ${error?.message || 'Unknown error'}` });
    }
});

// This downloads the ZIP file for the latest certificate batch.
app.get('/download-zip/:zipName', (req, res) => {
    const file = path.join(DIRS.certificates, path.basename(req.params.zipName || ''));
    if (!fs.existsSync(file)) return res.status(404).json({ success: false, message: 'ZIP file not found' });
    res.download(file);
});

// This opens the certificate verification page.
app.get('/verify/:reg_no', (_, res) => res.sendFile(path.join(FRONTEND_DIR, 'verify.html')));

// This checks one registration number in the database.
app.get('/api/verify/:reg_no', async (req, res) => {
    try {
        const [rows] = await db.query('SELECT * FROM certificates WHERE reg_no=? LIMIT 1', [req.params.reg_no]);
        if (!rows.length) return res.status(404).json({ verified: false, reg_no: req.params.reg_no, message: 'No valid certificate record was found.' });
        res.json({ verified: true, certificate: rows[0] });
    } catch {
        res.status(500).json({ verified: false, reg_no: req.params.reg_no, message: 'Verification failed' });
    }
});

// This sends dashboard numbers like totals, chart data, and last upload info.
app.get('/stats', async (_, res) => {
    try {
        const [totalResult, typesResult, coursesResult] = await Promise.all([
            db.query('SELECT COUNT(*) count FROM certificates'),
            db.query('SELECT course_type label, COUNT(*) value FROM certificates GROUP BY course_type ORDER BY value DESC'),
            db.query('SELECT course_name label, COUNT(*) value FROM certificates GROUP BY course_name ORDER BY value DESC')
        ]);
        const total = totalResult[0];
        const types = typesResult[0];
        const courses = coursesResult[0];
        res.json({
            totalCertificatesGenerated: Number(total[0]?.count) || 0,
            totalCandidatesNotCreated: uploadSummary.invalidRows || 0,
            courseTypes: types,
            courseNames: courses,
            currentUpload: uploadSummary,
            selectedTemplate,
            certificateConfig,
            latestGeneration: lastGeneration
        });
    } catch {
        res.status(500).json({ success: false, message: 'Failed to load dashboard stats' });
    }
});

// This adds a new employee account from the admin page.
app.post('/add-employee', async (req, res) => {
    try {
        await db.query('INSERT INTO employees (name,email,password,phone,role) VALUES (?, ?, ?, ?, ?)', [req.body.name, req.body.email, hashPassword(req.body.password), req.body.phone, 'employee']);
        res.json({ success: true });
    } catch {
        res.status(500).json({ success: false, message: 'Failed to add employee' });
    }
});

// This loads the employee list for the admin page.
app.get('/employees', async (_, res) => {
    try {
        const [rows] = await db.query("SELECT id,name,email,phone,role FROM employees WHERE role <> 'admin' OR role IS NULL ORDER BY id DESC");
        res.json(rows);
    } catch {
        res.status(500).json({ success: false, message: 'Failed to load employees' });
    }
});

// This deletes an employee account, but keeps admin accounts protected.
app.delete('/delete-employee/:id', async (req, res) => {
    try {
        await db.query("DELETE FROM employees WHERE id=? AND (role <> 'admin' OR role IS NULL)", [req.params.id]);
        res.json({ success: true });
    } catch {
        res.status(500).json({ success: false, message: 'Failed to delete employee' });
    }
});

// This starts the backend server and prepares default accounts for first-time use.
app.listen(PORT, HOST, async () => {
    try {
        await ensureSchema();
        await ensureDefaultUsers();
    } catch (error) {
        console.warn('Database startup setup failed:', error.message);
    }
    console.log(`Server running on http://localhost:${PORT}`);
});
