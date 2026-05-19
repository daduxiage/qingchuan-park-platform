const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec, spawn } = require('child_process');

const PORT = process.env.PORT || 8520;
const ROOT = __dirname;
const DB_PATH = path.join(ROOT, '.data', 'park.db');

// ===== 简易路由 =====
const routes = {};

// ===== SQLite 数据库 =====
let db = null;
function initDB() {
    const dataDir = path.join(ROOT, '.data');
    if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

    // 简易 JSON 数据库（不依赖 native sqlite3）
    const tables = {
        events: path.join(dataDir, 'events.json'),
        users: path.join(dataDir, 'users.json'),
        licenses: path.join(dataDir, 'licenses.json'),
        credit_scores: path.join(dataDir, 'credit_scores.json'),
        settle_requests: path.join(dataDir, 'settle_requests.json'),
        api_interfaces: path.join(dataDir, 'api_interfaces.json'),
        info_publish: path.join(dataDir, 'info_publish.json'),
    };

    // 初始化所有表文件
    const defaults = {
        events: '[]',
        users: JSON.stringify([{ username: 'admin', password: 'admin123', display_name: '管理员', role: 'admin', company: '' }, { username: 'jxly', password: 'jxly123', display_name: '佳兴铝业', role: 'enterprise', company: '佳兴铝业' }]),
        licenses: '[]', credit_scores: '[]', settle_requests: '[]',
        api_interfaces: '[]', info_publish: '[]'
    };

    const tableFiles = {};
    for (const [name, filePath] of Object.entries(tables)) {
        if (!fs.existsSync(filePath)) fs.writeFileSync(filePath, defaults[name] || '[]');
        try { JSON.parse(fs.readFileSync(filePath, 'utf-8')); } catch (e) { fs.writeFileSync(filePath, defaults[name] || '[]'); }
        tableFiles[name] = filePath;
    }

    db = {
        // 读取整个表
        all: (table) => JSON.parse(fs.readFileSync(tableFiles[table], 'utf-8')),
        // 查询（简单过滤）
        query: (table, filterFn) => {
            const rows = JSON.parse(fs.readFileSync(tableFiles[table], 'utf-8'));
            return rows.filter(filterFn);
        },
        // 插入
        insert: (table, row) => {
            const rows = JSON.parse(fs.readFileSync(tableFiles[table], 'utf-8'));
            row.id = rows.length ? (rows[rows.length - 1].id + 1) : 1;
            rows.push(row);
            fs.writeFileSync(tableFiles[table], JSON.stringify(rows));
            return row;
        },
        // 更新
        update: (table, id, data) => {
            const rows = JSON.parse(fs.readFileSync(tableFiles[table], 'utf-8'));
            const idx = rows.findIndex(r => r.id === id);
            if (idx >= 0) { rows[idx] = { ...rows[idx], ...data }; fs.writeFileSync(tableFiles[table], JSON.stringify(rows)); return true; }
            return false;
        },
        // 删除
        delete: (table, id) => {
            const rows = JSON.parse(fs.readFileSync(tableFiles[table], 'utf-8'));
            const newRows = rows.filter(r => r.id !== id);
            fs.writeFileSync(tableFiles[table], JSON.stringify(newRows));
            return newRows.length < rows.length;
        },
        // 数量
        count: (table) => JSON.parse(fs.readFileSync(tableFiles[table], 'utf-8')).length,
    };

    console.log('数据库初始化完成');
    console.log(`  events: ${db.count('events')} 条`);
    console.log(`  users: ${db.count('users')} 条`);
}

// ===== API 路由定义 =====

// 登录
routes['POST:/api/auth/login'] = (req, res) => {
    const { username, password } = req.body;
    const users = db.all('users');
    const user = users.find(u => u.username === username && u.password === password);
    if (user) {
        res.json({ ok: true, user: { username: user.username, display_name: user.display_name, role: user.role, company: user.company } });
    } else {
        res.status(401).json({ ok: false, error: '用户名或密码错误' });
    }
};

// 事件日志 - GET/POST
routes['GET:/api/log'] = (req, res) => {
    let events = db.all('events');
    if (req.query.user) events = events.filter(e => e.username === req.query.user);
    if (req.query.event) events = events.filter(e => e.event_type === req.query.event);
    res.json(events);
};
routes['POST:/api/log'] = (req, res) => {
    const entry = req.body;
    entry.ip = getIP(req);
    if (!entry.time) entry.time = new Date().toISOString();
    const row = db.insert('events', entry);
    res.json({ ok: true, id: row.id });
};

// 证照 CRUD
routes['GET:/api/licenses'] = (req, res) => res.json(db.all('licenses'));
routes['POST:/api/licenses'] = (req, res) => { const row = db.insert('licenses', req.body); res.json({ ok: true, id: row.id }); };
routes['PUT:/api/licenses/:id'] = (req, res) => { const ok = db.update('licenses', parseInt(req.params.id), req.body); res.json({ ok }); };
routes['DELETE:/api/licenses/:id'] = (req, res) => { const ok = db.delete('licenses', parseInt(req.params.id)); res.json({ ok }); };

// 信用评分 CRUD
routes['GET:/api/credit_scores'] = (req, res) => res.json(db.all('credit_scores'));
routes['POST:/api/credit_scores'] = (req, res) => { const row = db.insert('credit_scores', req.body); res.json({ ok: true, id: row.id }); };

// 企业入驻 CRUD
routes['GET:/api/settle_requests'] = (req, res) => res.json(db.all('settle_requests'));
routes['POST:/api/settle_requests'] = (req, res) => { const row = db.insert('settle_requests', req.body); res.json({ ok: true, id: row.id }); };

// 信息发布 CRUD
routes['GET:/api/info_publish'] = (req, res) => res.json(db.all('info_publish'));
routes['POST:/api/info_publish'] = (req, res) => { const row = db.insert('info_publish', req.body); res.json({ ok: true, id: row.id }); };

// 接口监控 CRUD
routes['GET:/api/interfaces'] = (req, res) => res.json(db.all('api_interfaces'));
routes['POST:/api/interfaces'] = (req, res) => { const row = db.insert('api_interfaces', req.body); res.json({ ok: true, id: row.id }); };

// 健康检查
routes['GET:/api/health'] = (req, res) => {
    res.json({ status: 'ok', uptime: process.uptime(), time: new Date().toISOString() });
};

// 运维面板 - 综合状态
routes['GET:/api/admin/status'] = (req, res) => {
    var info = {
        server: { uptime: process.uptime(), node: process.version, memory: process.memoryUsage(), pid: process.pid, time: new Date().toISOString() },
        git: { local: '', remote: '', behind: 0 },
        pm2: ''
    };
    exec('cd '+__dirname+' && git log -1 --format="%h %s (%cr)"', function(e, o) {
        info.git.local = (o || '').trim() || 'N/A';
        // 用 ls-remote 获取真实远程 HEAD（浅克隆下 git log origin/main 不准）
        exec('cd '+__dirname+' && git ls-remote origin main 2>/dev/null | cut -f1', function(e2, o2) {
            var remoteHash = (o2 || '').trim();
            if (remoteHash) {
                exec('cd '+__dirname+' && git log ' + remoteHash + ' -1 --format="%h %s (%cr)" 2>/dev/null || echo "' + remoteHash.slice(0,7) + '"', function(e2b, o2b) {
                    info.git.remote = (o2b || '').trim() || remoteHash.slice(0,7);
                    finish(remoteHash);
                });
            } else { info.git.remote = 'N/A'; finish(null); }
        });
        function finish(remoteHash) {
            var behindCmd = remoteHash ? 'git rev-list HEAD..'+remoteHash+' --count 2>/dev/null || echo 0' : 'echo 0';
            exec('cd '+__dirname+' && ' + behindCmd, function(e3, o3) {
                info.git.behind = parseInt((o3 || '0').trim()) || 0;
                exec('pm2 jlist 2>/dev/null || echo "[]"', function(e4, o4) {
                    try { var j = JSON.parse(o4); info.pm2 = j[0] ? j[0].pm2_env.status : 'N/A'; } catch(x) {}
                    res.json({ ok: true, data: info });
                });
            });
        }
    });
};

// 自动部署
routes['POST:/api/deploy'] = (req, res) => {
    var p = req.body || {};
    if (p.token !== 'park2024') {
        res.writeHead(403); res.end(JSON.stringify({ ok: false, error: 'invalid token' }));
        return;
    }
    exec('cd '+__dirname+' && git pull origin main', { timeout: 180000 }, function(err, stdout, stderr) {
        var result = (stdout + stderr).trim();
        res.end(JSON.stringify({ ok: !err, output: result }));
        if (!err) {
            var child = spawn('pm2', ['restart', 'park-platform'], { detached: true, stdio: 'ignore' });
            child.unref();
        }
    });
};

// GitHub 连通检测（服务器端代理）
routes['GET:/api/admin/github-check'] = (req, res) => {
    var opts = { hostname:'api.github.com', path:'/repos/daduxiage/qingchuan-park-platform', method:'HEAD', timeout:8000 };
    var h = (typeof require('https')!=='undefined')?require('https'):require('http');
    var r = h.request(opts, function(resp){ res.json({ok:true,status:resp.statusCode}); });
    r.on('error', function(){ res.json({ok:false,status:0}); });
    r.on('timeout', function(){ r.destroy(); res.json({ok:false,status:0}); });
    r.end();
};

// ===== 辅助函数 =====
function getIP(req) {
    const xff = req.headers['x-forwarded-for'];
    if (xff) return xff.split(',')[0].trim();
    const real = req.headers['x-real-ip'];
    if (real) return real.trim();
    return req.socket.remoteAddress.replace('::ffff:', '');
}

function parseJSON(str) {
    try { return JSON.parse(str); } catch (e) { return str; }
}

// ===== MIME 类型 =====
const MIME = {
    '.html': 'text/html;charset=utf-8', '.css': 'text/css;charset=utf-8',
    '.js': 'application/javascript;charset=utf-8', '.json': 'application/json;charset=utf-8',
    '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
    '.gif': 'image/gif', '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
    '.woff2': 'font/woff2', '.woff': 'font/woff', '.ttf': 'font/ttf',
    '.map': 'application/json', '.txt': 'text/plain;charset=utf-8',
};

// ===== HTTP 服务 =====
const server = http.createServer((req, res) => {
    // CORS
    const origin = req.headers.origin || '*';
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type,Authorization,X-Requested-With,ngrok-skip-browser-warning');

    if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

    const parsedUrl = req.url.split('?')[0];

    // 解析 Body
    if (['POST', 'PUT'].includes(req.method)) {
        let body = '';
        req.on('data', chunk => body += chunk);
        req.on('end', () => {
            req.body = parseJSON(body);
            handleRequest(req, res, parsedUrl);
        });
    } else {
        handleRequest(req, res, parsedUrl);
    }
});

function handleRequest(req, res, urlPath) {
    // 添加 res.json 辅助方法
    res.json = function(obj) {
        this.writeHead(this.statusCode || 200, { 'Content-Type': 'application/json;charset=utf-8' });
        this.end(JSON.stringify(obj));
    };
    res.status = function(code) { this.statusCode = code; return this; };
    // 先匹配 API 路由
    const routeKey = `${req.method}:${urlPath}`;
    if (routes[routeKey]) {
        try { routes[routeKey](req, res); } catch (e) { res.writeHead(500); res.end(JSON.stringify({ ok: false, error: e.message })); }
        return;
    }
    // 匹配带参数的路由
    for (const key of Object.keys(routes)) {
        const [method, pattern] = key.split(':');
        if (method !== req.method) continue;
        const patternParts = pattern.split('/');
        const urlParts = urlPath.split('/');
        if (patternParts.length !== urlParts.length) continue;
        let match = true;
        const params = {};
        for (let i = 0; i < patternParts.length; i++) {
            if (patternParts[i].startsWith(':')) {
                params[patternParts[i].slice(1)] = urlParts[i];
            } else if (patternParts[i] !== urlParts[i]) {
                match = false; break;
            }
        }
        if (match) { req.params = params; try { routes[key](req, res); } catch (e) { res.writeHead(500); res.end(JSON.stringify({ ok: false, error: e.message })); } return; }
    }

    // 静态文件
    // 修复重复的 /pages/ 路径 → 301 跳转
    let filePath = urlPath;
    var cleaned = filePath;
    while (cleaned.includes('/pages/pages/')) {
        cleaned = cleaned.replace('/pages/pages/', '/pages/');
    }
    if (cleaned !== filePath) {
        res.writeHead(301, { 'Location': cleaned });
        return res.end();
    }
    if (filePath === '/') filePath = '/index.html';
    const fullPath = path.join(ROOT, filePath);

    fs.readFile(fullPath, (err, data) => {
        if (err) {
            // SPA 或 404
            if (!urlPath.startsWith('/api/')) {
                fs.readFile(path.join(ROOT, 'index.html'), (err2, data2) => {
                    res.writeHead(err2 ? 404 : 200, { 'Content-Type': 'text/html;charset=utf-8' });
                    res.end(err2 ? '404 Not Found' : data2);
                });
            } else {
                res.writeHead(404, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: 'API not found' }));
            }
            return;
        }
        const stat = fs.statSync(fullPath);
        const lastMod = stat.mtime.toUTCString();
        const ifModSince = req.headers['if-modified-since'];
        if (ifModSince && new Date(ifModSince) >= stat.mtime) {
            res.writeHead(304, { 'Last-Modified': lastMod });
            res.end();
            return;
        }
        const ext = path.extname(fullPath);
        const ct = MIME[ext] || 'application/octet-stream';
        const maxAge = (ext === '.html' || ext === '.json') ? 'no-cache' : 'public,max-age=604800,immutable';
        res.writeHead(200, { 'Content-Type': ct, 'Cache-Control': maxAge, 'Last-Modified': lastMod });
        res.end(data);
    });
}

// ===== 初始化 =====
initDB();

server.listen(PORT, '0.0.0.0', () => {
    console.log(`✅ 青川智慧工业园平台已启动`);
    console.log(`   HTTP:  http://0.0.0.0:${PORT}`);
    console.log(`   API:   http://0.0.0.0:${PORT}/api/health`);
});
