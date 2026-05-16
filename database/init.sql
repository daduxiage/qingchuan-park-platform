-- 平台核心数据库初始化
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

-- 用户表
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    display_name TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'admin',  -- admin / enterprise
    company TEXT DEFAULT '',
    status TEXT DEFAULT 'active',
    created_at TEXT DEFAULT (datetime('now','localtime'))
);

-- 默认用户: admin / admin123, jxly / jxly123
INSERT OR IGNORE INTO users (username, password_hash, display_name, role, company) VALUES
('admin', '$2b$10$placeholder_hash_admin123', '管理员', 'admin', ''),
('jxly', '$2b$10$placeholder_hash_jxly123', '佳兴铝业', 'enterprise', '佳兴铝业');

-- 操作事件日志表（替代 events.json）
CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    username TEXT NOT NULL DEFAULT '未知',
    event_type TEXT NOT NULL,
    detail TEXT,
    ip TEXT DEFAULT '--',
    page TEXT,
    ua TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
);

-- 信用评分表
CREATE TABLE IF NOT EXISTS credit_scores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sn TEXT NOT NULL UNIQUE,
    company TEXT NOT NULL,
    level TEXT NOT NULL,
    total_score INTEGER DEFAULT 80,
    score_law INTEGER DEFAULT 80,
    score_safety INTEGER DEFAULT 80,
    score_env INTEGER DEFAULT 80,
    score_social INTEGER DEFAULT 80,
    eval_date TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now','localtime'))
);

-- 证照管理表
CREATE TABLE IF NOT EXISTS licenses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sn TEXT NOT NULL UNIQUE,
    company TEXT NOT NULL,
    cert_name TEXT NOT NULL,
    status TEXT DEFAULT '有效',
    expire_date TEXT,
    entry_time TEXT DEFAULT (datetime('now','localtime')),
    entry_by TEXT DEFAULT '管理员',
    event_desc TEXT
);

-- 企业入驻表
CREATE TABLE IF NOT EXISTS settle_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sn TEXT NOT NULL UNIQUE,
    company TEXT NOT NULL,
    step TEXT NOT NULL DEFAULT '提交申请',
    company_type TEXT,
    files_count TEXT,
    apply_time TEXT DEFAULT (datetime('now','localtime')),
    duration TEXT DEFAULT '审核中'
);

-- 信息发布表
CREATE TABLE IF NOT EXISTS info_publish (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sn TEXT NOT NULL UNIQUE,
    module TEXT NOT NULL,
    event_type TEXT,
    detail TEXT,
    publish_time TEXT DEFAULT (datetime('now','localtime')),
    publisher TEXT DEFAULT '管理员',
    status TEXT DEFAULT '已发布'
);

-- 接口监控表
CREATE TABLE IF NOT EXISTS api_interfaces (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sn TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    url TEXT,
    type TEXT,
    auth_method TEXT,
    health TEXT DEFAULT '正常',
    access_time TEXT
);
