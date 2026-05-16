/**
 * common.js — 青川县庄子上工业园区智慧园区平台 通用功能库
 * 所有页面共享：分页、新增、删除、查询、地图全屏、主题
 */

/* ═══ 0. 全局主题基础函数（最优先加载，覆盖各页面的破损定义）═══ */
(function() {
    /* 初始化主题：从 localStorage 恢复 */
    var saved = localStorage.getItem('ss_theme');
    if (saved !== 'dark') document.documentElement.classList.add('light-theme');

    /* 定义核心 toggleTheme（各页面若有自己的会在 common.js 之前定义，
       但 common.js 的 initPageHeaderBtns 会正确拦截） */
    if (!window._themeBaseInited) {
        window._themeBaseInited = true;
        /* 如果页面自己没有定义 toggleTheme，提供一个完整默认实现 */
        if (!window.toggleTheme) {
            window.toggleTheme = function() {
                var isLt = document.documentElement.classList.toggle('light-theme');
                localStorage.setItem('ss_theme', isLt ? 'light' : 'dark');
            };
        }
    }
})();

(function(window) {
'use strict';

/* ─────────────────── 1. Toast 通知 ─────────────────── */
window.showToast = function(msg, type, duration) {
    type = type || 'info';
    duration = duration || 2500;
    var t = document.getElementById('__toast');
    if (!t) {
        t = document.createElement('div');
        t.id = '__toast';
        t.style.cssText = 'position:fixed;top:24px;right:24px;z-index:99999;display:flex;flex-direction:column;gap:8px;pointer-events:none';
        document.body.appendChild(t);
    }
    var icons = {success:'✅',error:'❌',warning:'⚠️',info:'ℹ️'};
    var colors = {success:'#4cd964',error:'#ff4d4f',warning:'#ff9500',info:'#40c4ff'};
    var item = document.createElement('div');
    item.style.cssText = 'background:rgba(16,24,56,0.95);backdrop-filter:blur(12px);border:1px solid ' + (colors[type]||colors.info) + ';border-radius:8px;padding:10px 18px;color:#e8f4ff;font-size:13px;display:flex;align-items:center;gap:8px;box-shadow:0 4px 20px rgba(0,0,0,0.4);transform:translateX(100px);opacity:0;transition:all 0.3s;pointer-events:auto';
    item.innerHTML = '<span>' + (icons[type]||icons.info) + '</span><span>' + msg + '</span>';
    t.appendChild(item);
    requestAnimationFrame(function() { item.style.transform='translateX(0)'; item.style.opacity='1'; });
    setTimeout(function() {
        item.style.transform='translateX(100px)'; item.style.opacity='0';
        setTimeout(function() { if(item.parentNode) item.parentNode.removeChild(item); }, 300);
    }, duration);
};

/* ─────────────────── 2. 通用弹窗基础 ─────────────────── */
window.isLight = function() { return document.documentElement.classList.contains('light-theme'); };

function getModalTheme() {
    if (isLight()) {
        return {
            overlay: 'position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,0.35);display:flex;align-items:center;justify-content:center',
            modal: 'background:#fff;border:1px solid #d0dff0;border-radius:12px;min-width:400px;max-width:90vw;max-height:90vh;display:flex;flex-direction:column;box-shadow:0 8px 40px rgba(0,0,0,0.18)',
            header: 'padding:16px 20px;border-bottom:1px solid #e0ebf5;display:flex;align-items:center;justify-content:space-between',
            title: 'font-size:15px;font-weight:600;color:#1a2a4a',
            close: 'cursor:pointer;font-size:20px;color:#6080a0;line-height:1;padding:2px 6px;border-radius:4px',
            body: 'padding:20px;overflow-y:auto;flex:1;color:#1a2a4a',
            footer: 'padding:12px 20px;border-top:1px solid #e0ebf5;display:flex;justify-content:flex-end;gap:8px'
        };
    }
    return {
        overlay: 'position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,0.6);display:flex;align-items:center;justify-content:center',
        modal: 'background:#0d1633;border:1px solid rgba(43,104,224,0.4);border-radius:12px;min-width:400px;max-width:90vw;max-height:90vh;display:flex;flex-direction:column;box-shadow:0 8px 40px rgba(0,0,0,0.6)',
        header: 'padding:16px 20px;border-bottom:1px solid rgba(43,104,224,0.3);display:flex;align-items:center;justify-content:space-between',
        title: 'font-size:15px;font-weight:600;color:#e8f4ff',
        close: 'cursor:pointer;font-size:20px;color:#a0cfff;line-height:1;padding:2px 6px;border-radius:4px',
        body: 'padding:20px;overflow-y:auto;flex:1;color:#c8e0f4',
        footer: 'padding:12px 20px;border-top:1px solid rgba(43,104,224,0.3);display:flex;justify-content:flex-end;gap:8px'
    };
}

window.createModalBase = function(id) {
    var old = document.getElementById(id || '__globalModal');
    if (old) old.remove();
    var theme = getModalTheme();
    var overlay = document.createElement('div');
    overlay.id = id || '__globalModal';
    overlay.style.cssText = theme.overlay;
    var box = document.createElement('div');
    box.style.cssText = theme.modal;
    overlay.appendChild(box);
    overlay.addEventListener('click', function(e) { if (e.target === overlay) overlay.remove(); });
    // ESC 关闭
    var esc = function(e) { if (e.key === 'Escape') { overlay.remove(); document.removeEventListener('keydown', esc); } };
    document.addEventListener('keydown', esc);
    document.body.appendChild(overlay);
    return { overlay: overlay, box: box, theme: theme };
}

/* ─────────────────── 3. 确认弹窗 ─────────────────── */
window.createConfirmModal = function(title, content, onConfirm, onCancel) {
    var m = createModalBase();
    var t = m.theme;
    var header = document.createElement('div');
    header.style.cssText = t.header;
    header.innerHTML = '<span style="' + t.title + '">' + title + '</span>';
    var closeBtn = document.createElement('span');
    closeBtn.style.cssText = t.close;
    closeBtn.innerHTML = '&times;';
    closeBtn.onclick = function() { m.overlay.remove(); if (onCancel) onCancel(); };
    header.appendChild(closeBtn);

    var body = document.createElement('div');
    body.style.cssText = t.body + ';font-size:14px;line-height:1.8';
    body.innerHTML = content;

    var footer = document.createElement('div');
    footer.style.cssText = t.footer;
    var cancelBtn = document.createElement('button');
    cancelBtn.className = 'btn btn-default';
    cancelBtn.textContent = '取消';
    cancelBtn.onclick = function() { m.overlay.remove(); if (onCancel) onCancel(); };
    var confirmBtn = document.createElement('button');
    confirmBtn.className = 'btn btn-primary';
    confirmBtn.textContent = '确认';
    confirmBtn.onclick = function() { m.overlay.remove(); if (onConfirm) onConfirm(); };
    footer.appendChild(cancelBtn);
    footer.appendChild(confirmBtn);

    m.box.appendChild(header);
    m.box.appendChild(body);
    m.box.appendChild(footer);
};

/* ─────────────────── 4. 表单新增弹窗 ─────────────────── */
window.createFormModal = function(title, fields, onSave, extraBtns) {
    var m = createModalBase();
    var t = m.theme;
    m.box.style.cssText = m.box.style.cssText + ';min-width:500px;max-width:700px';

    var header = document.createElement('div');
    header.style.cssText = t.header;
    header.innerHTML = '<span style="' + t.title + '">' + title + '</span>';
    var closeBtn = document.createElement('span');
    closeBtn.style.cssText = t.close;
    closeBtn.innerHTML = '&times;';
    closeBtn.onclick = function() { m.overlay.remove(); };
    header.appendChild(closeBtn);

    var body = document.createElement('div');
    body.style.cssText = t.body;
    var formRow = document.createElement('div');
    formRow.className = 'form-row';
    formRow.style.cssText = 'flex-wrap:wrap;gap:12px';

    fields.forEach(function(f) {
        var item = document.createElement('div');
        item.className = 'form-item';
        item.style.cssText = 'min-width:' + (f.fullWidth ? '100%' : '200px') + ';flex:1';
        var label = document.createElement('div');
        label.className = 'form-label';
        label.textContent = f.label + (f.required ? ' *' : '');
        item.appendChild(label);

        var ctrl;
        if (f.type === 'select') {
            ctrl = document.createElement('select');
            ctrl.className = 'form-control';
            (f.options || []).forEach(function(o) {
                var opt = document.createElement('option');
                opt.value = typeof o === 'object' ? o.value : o;
                opt.textContent = typeof o === 'object' ? o.label : o;
                ctrl.appendChild(opt);
            });
            if (f.value) ctrl.value = f.value;
        } else if (f.type === 'textarea') {
            ctrl = document.createElement('textarea');
            ctrl.className = 'form-control';
            ctrl.rows = 3;
            ctrl.style.resize = 'vertical';
            ctrl.value = f.value || '';
            ctrl.placeholder = f.placeholder || '请输入';
            item.style.minWidth = '100%';
        } else {
            ctrl = document.createElement('input');
            ctrl.className = 'form-control';
            ctrl.type = f.type || 'text';
            ctrl.placeholder = f.placeholder || '请输入';
            ctrl.value = f.value || '';
        }
        ctrl.dataset.field = f.name;
        if (f.required) ctrl.required = true;
        item.appendChild(ctrl);
        formRow.appendChild(item);
    });
    body.appendChild(formRow);

    var footer = document.createElement('div');
    footer.style.cssText = t.footer;
    var cancelBtn = document.createElement('button');
    cancelBtn.className = 'btn btn-default';
    cancelBtn.textContent = '取消';
    cancelBtn.onclick = function() { m.overlay.remove(); };
    var saveBtn = document.createElement('button');
    saveBtn.className = 'btn btn-primary';
    saveBtn.textContent = '保存';
    saveBtn.onclick = function() {
        var data = {};
        var valid = true;
        formRow.querySelectorAll('[data-field]').forEach(function(el) {
            if (el.required && !el.value.trim()) {
                el.style.borderColor = '#ff4d4f';
                valid = false;
            } else {
                el.style.borderColor = '';
                data[el.dataset.field] = el.value;
            }
        });
        if (!valid) { showToast('请填写必填项', 'warning'); return; }
        m.overlay.remove();
        if (onSave) onSave(data);
    };
    if (extraBtns) extraBtns.forEach(function(b) { footer.appendChild(b); });
    footer.appendChild(cancelBtn);
    footer.appendChild(saveBtn);

    m.box.appendChild(header);
    m.box.appendChild(body);
    m.box.appendChild(footer);
};

/* ─────────────────── 5. 数据生成器（虚拟数据引擎） ─────────────────── */
var DataGen = window._DataGen = {
    surnames: '张李王赵钱孙周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施',
    names1: '建国明远大勇小飞刚强志远',
    names2: '华春伟民磊超',
    plates: ['川A','川B','川C','川F','川G','川H','川J','川K'],
    companies: ['龙门运输有限公司','彭州物流公司','通济建材运输队','新兴砂石运输公司','彭州市建材总公司','成都砂石运输集团','川西物流有限公司','蜀兴运输公司','峨眉运输集团','新都货运公司'],
    statuses: ['正常','临期','暂停','注销'],
    vehicleTypes: ['自卸货车','半挂车','搅拌车','平板车','箱式货车'],

    rand: function(arr) { return arr[Math.floor(Math.random() * arr.length)]; },
    randInt: function(a, b) { return Math.floor(Math.random() * (b - a + 1)) + a; },
    randDate: function(future) {
        var d = new Date();
        d.setDate(d.getDate() + (future ? DataGen.randInt(30, 730) : DataGen.randInt(-365, 365)));
        return d.toISOString().slice(0,10);
    },
    randPlate: function() {
        var p = DataGen.rand(DataGen.plates);
        var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ0123456789';
        var n = '';
        for (var i=0;i<5;i++) n += chars[DataGen.randInt(0, chars.length-1)];
        return p + n;
    },
    randPhone: function() { return '1' + DataGen.rand(['3','5','7','8','9']) + DataGen.rand(['3','5','6','7','8','9']) + '****' + DataGen.randInt(1000,9999); },
    randName: function() {
        return DataGen.surnames[DataGen.randInt(0, DataGen.surnames.length-1)] +
            DataGen.names1[DataGen.randInt(0, DataGen.names1.length-1)] +
            (Math.random()>0.4 ? DataGen.names2[DataGen.randInt(0, DataGen.names2.length-1)] : '');
    },
    randCompany: function() { return DataGen.rand(DataGen.companies); },
    randAmount: function(a, b) { return (DataGen.randInt(a*10, b*10)/10).toFixed(1); }
};

/* ─────────────────── 6. 分页引擎 ─────────────────── */
var PAGE_SIZE = 10;

function buildRow(template, page, idx, rowIdx) {
    // 通过循环复用模板行模拟翻页
    var cells = template.cells;
    var newRow = template.row.cloneNode(true);
    var tds = newRow.querySelectorAll('td');
    // 对部分列注入随机值让每行看起来不同
    tds.forEach(function(td, ci) {
        var orig = cells[ci] ? cells[ci].innerHTML : td.innerHTML;
        // 保留状态标签和操作按钮，只替换纯文字内容
        if (!td.querySelector('button,span.status-tag,input,select')) {
            // 对一些特征列做随机替换
            var txt = td.textContent.trim();
            // 车牌：川X+数字字母
            if (/^川[A-Z]\d/.test(txt)) {
                td.textContent = DataGen.randPlate();
            }
            // 手机号
            else if (/^1[3-9]\d{1}[\*\d]{4}\d{4}$/.test(txt) || /13[0-9]\*/.test(txt)) {
                td.textContent = DataGen.randPhone();
            }
            // 日期
            else if (/^\d{4}-\d{2}-\d{2}$/.test(txt)) {
                td.textContent = DataGen.randDate(true);
            }
            // 纯数字（非状态）
            else if (/^\d+(\.\d+)?$/.test(txt) && txt.length < 6) {
                var n = parseFloat(txt);
                if (n > 0) td.textContent = (DataGen.randInt(Math.max(1,n*0.7|0), n*1.3|0));
            }
        }
    });
    return newRow;
}

function initOnePagination(card) {
    var table = card.querySelector('.data-table');
    if (!table) return;
    var tbody = table.querySelector('tbody');
    if (!tbody) return;
    var pagination = card.querySelector('.pagination');

    // 读取 total
    var totalRecords = 0;
    var infoSpan = null;
    if (pagination) {
        infoSpan = pagination.querySelector('span');
        var txt = infoSpan ? infoSpan.textContent : '';
        var m = txt.match(/[\d,]+/);
        if (m) totalRecords = parseInt(m[0].replace(/,/g,''));
    }
    // 也从 card-header 里的"共 X 辆/条/件"里读取
    if (!totalRecords) {
        var hdr = card.querySelector('.card-header');
        if (hdr) {
            var hm = hdr.textContent.match(/共\s*([\d,]+)/);
            if (hm) totalRecords = parseInt(hm[1].replace(/,/g,''));
        }
    }

    // 收集模板行
    var templateRows = Array.from(tbody.querySelectorAll('tr'));
    if (!templateRows.length) return;
    totalRecords = totalRecords || templateRows.length;
    var totalPages = Math.max(1, Math.ceil(totalRecords / PAGE_SIZE));

    // 存储原始行数据作为模板
    var templates = templateRows.map(function(r) {
        return { row: r.cloneNode(true), cells: Array.from(r.querySelectorAll('td')).map(function(c) { return c.cloneNode(true); }) };
    });

    // 为操作按钮绑定事件（新增行时重新绑定）
    function bindRowActions(row) {
        // 删除按钮
        row.querySelectorAll('.btn-delete,[data-action="delete"]').forEach(function(btn) {
            btn.onclick = function(e) {
                e.stopPropagation();
                var rowEl = btn.closest('tr');
                createConfirmModal('确认删除', '确定要删除这条记录吗？删除后无法恢复。', function() {
                    if (rowEl) rowEl.style.animation = 'fadeOut 0.3s forwards';
                    setTimeout(function() {
                        if (rowEl && rowEl.parentNode) rowEl.parentNode.removeChild(rowEl);
                        totalRecords = Math.max(0, totalRecords-1);
                        templates = templates.filter(function(t) { return t.row !== rowEl; });
                        showToast('删除成功', 'success');
                        renderPage(currentPage);
                    }, 300);
                });
            };
        });
    }

    var currentPage = 1;
    // 确保pagination存在
    if (!pagination) {
        pagination = document.createElement('div');
        pagination.className = 'pagination';
        infoSpan = document.createElement('span');
        pagination.appendChild(infoSpan);
        var btnsDiv = document.createElement('div');
        btnsDiv.className = 'page-btns';
        pagination.appendChild(btnsDiv);
        card.appendChild(pagination);
    }
    if (!pagination.querySelector('.page-btns')) {
        var bd2 = document.createElement('div');
        bd2.className = 'page-btns';
        pagination.appendChild(bd2);
    }

    function renderPage(page) {
        currentPage = page;
        tbody.innerHTML = '';
        var tplCount = templates.length;
        if (!tplCount) return;
        var count = Math.min(PAGE_SIZE, totalRecords - (page-1)*PAGE_SIZE);
        for (var i = 0; i < count; i++) {
            var tpl = templates[i % tplCount];
            var row = buildRow(tpl, page, i, (page-1)*PAGE_SIZE + i);
            tbody.appendChild(row);
            bindRowActions(row);
        }
        // 更新info
        if (infoSpan) {
            var start = (page-1)*PAGE_SIZE+1;
            var end = Math.min(page*PAGE_SIZE, totalRecords);
            infoSpan.textContent = '共 ' + totalRecords.toLocaleString() + ' 条  第' + start + '-' + end + '条';
        }
        updateBtns(page, totalPages);
    }

    function updateBtns(page, total) {
        var btnsDiv = pagination.querySelector('.page-btns');
        if (!btnsDiv) return;
        btnsDiv.innerHTML = '';
        function mkBtn(label, pg, disabled, active) {
            var b = document.createElement('div');
            b.className = 'page-btn' + (active?' active':'') + (disabled?' disabled':'');
            b.textContent = label;
            if (!disabled && pg !== null) b.onclick = function() { renderPage(pg); };
            return b;
        }
        btnsDiv.appendChild(mkBtn('«', page-1, page<=1));
        var s = Math.max(1, page-2), e = Math.min(total, s+4);
        if (e-s < 4) s = Math.max(1, e-4);
        for (var p=s; p<=e; p++) btnsDiv.appendChild(mkBtn(p, p, false, p===page));
        btnsDiv.appendChild(mkBtn('»', page+1, page>=total));
    }

    // 绑定查询按钮
    var queryBtns = card.querySelectorAll('.btn-search,[data-action="search"]');
    queryBtns.forEach(function(btn) {
        btn.onclick = function() {
            currentPage = 1;
            renderPage(1);
            showToast('查询成功，共 ' + totalRecords.toLocaleString() + ' 条', 'success');
        };
    });

    // 绑定搜索表单的查询按钮（同一个 .card 内的查询按钮）
    var searchCard = card.previousElementSibling;
    if (searchCard && searchCard.classList.contains('card')) {
        var sBtns = searchCard.querySelectorAll('.btn-search,[data-action="search"]');
        sBtns.forEach(function(btn) {
            btn.onclick = function() {
                currentPage = 1;
                renderPage(1);
                showToast('查询成功', 'success');
            };
        });
    }

    renderPage(1);
    return { renderPage: renderPage, getTotal: function() { return totalRecords; }, addRow: function(row) { templates.unshift({row:row.cloneNode(true), cells:Array.from(row.querySelectorAll('td')).map(function(c){return c.cloneNode(true);})}); totalRecords++; renderPage(1); } };
}

/* ─────────────────── 7. 全局查询按钮 ─────────────────── */
function initSearchButtons() {
    // 找页面里标记为查询的按钮（card-header里的第一个btn-primary，或有 fa-search 图标的）
    document.querySelectorAll('.card').forEach(function(card) {
        var header = card.querySelector('.card-header');
        if (!header) return;
        var hasFormRow = card.querySelector('.form-row');
        if (!hasFormRow) return;
        // 找目标表格（同一个card或下一个card）
        var targetCard = card.nextElementSibling;
        var qtBtns = header.querySelectorAll('.btn-primary,.btn-search');
        qtBtns.forEach(function(btn) {
            if (btn.hasAttribute('data-search-init')) return;
            btn.setAttribute('data-search-init','1');
            var origClick = btn.onclick;
            btn.onclick = function(e) {
                if (origClick) origClick.call(btn, e);
                // 执行过滤
                var filters = {};
                card.querySelectorAll('[data-field],[name]').forEach(function(el) {
                    var k = el.dataset.field || el.name;
                    if (k && el.value) filters[k] = el.value.toLowerCase();
                });
                if (targetCard) {
                    var tbody = targetCard.querySelector('tbody');
                    if (tbody) {
                        var rows = tbody.querySelectorAll('tr');
                        rows.forEach(function(row) {
                            if (!Object.keys(filters).length) { row.style.display=''; return; }
                            var txt = row.textContent.toLowerCase();
                            var show = Object.values(filters).every(function(v) { return txt.includes(v); });
                            row.style.display = show ? '' : 'none';
                        });
                    }
                }
                showToast('查询完成', 'success');
            };
        });
    });
}

/* ─────────────────── 8. 新增按钮绑定 ─────────────────── */
function getFieldsFromTable(tableCard) {
    var headers = [];
    if (tableCard) {
        tableCard.querySelectorAll('thead th').forEach(function(th) {
            var txt = th.textContent.trim();
            if (txt && txt !== '操作') headers.push(txt);
        });
    }
    return headers.map(function(h) {
        var type = 'text';
        var opts = null;
        if (h.includes('状态') || h.includes('类型')) { type = 'select'; opts = ['请选择', h.includes('状态') ? '正常' : '类型A', h.includes('状态') ? '暂停' : '类型B']; }
        if (h.includes('日期') || h.includes('时间')) type = 'date';
        if (h.includes('电话') || h.includes('手机')) type = 'tel';
        if (h.includes('备注') || h.includes('描述')) type = 'textarea';
        return { name: h, label: h, type: type, options: opts, placeholder: '请输入' + h, required: false };
    });
}

function initAddButtons() {
    document.querySelectorAll('.btn-add,[data-action="add"]').forEach(function(btn) {
        if (btn.hasAttribute('data-add-init')) return;
        btn.setAttribute('data-add-init','1');
        btn.onclick = function(e) {
            e.stopPropagation();
            // 找对应表格
            var card = btn.closest('.card');
            var tableCard = null;
            if (card) {
                var sibling = card.nextElementSibling;
                if (sibling && sibling.querySelector('table')) tableCard = sibling;
                if (!tableCard && card.querySelector('table')) tableCard = card;
            }
            var fields = getFieldsFromTable(tableCard);
            if (!fields.length) fields = [{name:'name',label:'名称',type:'text',placeholder:'请输入名称',required:true}];
            // 尝试从 data-fields 属性获取字段定义
            var customFields = btn.dataset.fields;
            if (customFields) {
                try { fields = JSON.parse(customFields); } catch(e) {}
            }
            var modalTitle = btn.dataset.title || (btn.textContent.trim().replace(/^\s*[^\u4e00-\u9fa5]*/, '') || '新增记录');
            createFormModal(modalTitle, fields, function(data) {
                // 构建新行并插入表格
                if (tableCard) {
                    var tbody = tableCard.querySelector('tbody');
                    if (tbody) {
                        var existRow = tbody.querySelector('tr');
                        if (existRow) {
                            var newRow = existRow.cloneNode(true);
                            var tds = newRow.querySelectorAll('td');
                            var vals = Object.values(data);
                            tds.forEach(function(td, i) {
                                if (i < vals.length && !td.querySelector('button,span.status-tag')) {
                                    td.textContent = vals[i];
                                } else if (i === tds.length-1) {
                                    // 操作列保留
                                }
                            });
                            newRow.style.background = 'rgba(64,196,255,0.1)';
                            setTimeout(function() { newRow.style.background=''; }, 2000);
                            tbody.insertBefore(newRow, tbody.firstChild);
                        }
                    }
                }
                showToast('新增成功', 'success');
            });
        };
    });
}

/* ─────────────────── 9. 删除按钮绑定 ─────────────────── */
function initDeleteButtons() {
    document.querySelectorAll('.btn-danger,.btn-delete,[data-action="delete"]').forEach(function(btn) {
        if (btn.hasAttribute('data-del-init')) return;
        btn.setAttribute('data-del-init','1');
        if (!btn.textContent.includes('删') && !btn.dataset.action) return;
        btn.onclick = function(e) {
            e.stopPropagation();
            var row = btn.closest('tr');
            createConfirmModal('确认删除', '确定要删除这条记录吗？此操作不可恢复。', function() {
                if (row) {
                    row.style.transition = 'opacity 0.3s';
                    row.style.opacity = '0';
                    setTimeout(function() { if(row.parentNode) row.parentNode.removeChild(row); }, 300);
                }
                showToast('删除成功', 'success');
            });
        };
    });
}

/* ─────────────────── 10. 地图全屏 & 居中定位 ─────────────────── */
window.initMapControls = function(mapId, mapObj) {
    var el = document.getElementById(mapId);
    if (!el) return;
    var container = el.parentElement;

    // 添加控制按钮容器
    if (container.querySelector('.map-ctrl-btns')) return;
    var ctrlDiv = document.createElement('div');
    ctrlDiv.className = 'map-ctrl-btns';
    ctrlDiv.style.cssText = 'position:absolute;top:10px;right:10px;z-index:1000;display:flex;flex-direction:column;gap:6px';

    // 居中定位按钮
    var centerBtn = document.createElement('button');
    centerBtn.className = 'btn btn-default';
    centerBtn.style.cssText = 'padding:6px 10px;font-size:13px;background:rgba(13,22,51,0.85);backdrop-filter:blur(8px)';
    centerBtn.title = '居中定位';
    centerBtn.innerHTML = '⊕';
    centerBtn.onclick = function() {
        var m = mapObj || window[mapId + 'Obj'] || window._mapTrajectoryObj || window['_map' + mapId];
        if (m && m.setView) {
            m.setView([31.05, 103.95], 12);
            showToast('已居中定位', 'info');
        }
    };

    // 全屏按钮
    var fsBtn = document.createElement('button');
    fsBtn.className = 'btn btn-default';
    fsBtn.style.cssText = 'padding:6px 10px;font-size:13px;background:rgba(13,22,51,0.85);backdrop-filter:blur(8px)';
    fsBtn.title = '全屏地图';
    fsBtn.innerHTML = '⛶';
    fsBtn.onclick = function() { openFullscreenMap(mapId, mapObj); };

    ctrlDiv.appendChild(centerBtn);
    ctrlDiv.appendChild(fsBtn);
    if (container.style.position !== 'absolute' && container.style.position !== 'fixed') {
        container.style.position = 'relative';
    }
    container.appendChild(ctrlDiv);
};

window.openFullscreenMap = function(mapId, mapObj) {
    var m = createModalBase('__mapFullscreen');
    m.box.style.cssText = 'width:96vw;height:92vh;max-width:96vw;max-height:92vh;background:#0a1024;border:1px solid rgba(43,104,224,0.4);border-radius:12px;overflow:hidden;display:flex;flex-direction:column';
    if (isLight()) m.box.style.background = '#f0f4f8';

    var header = document.createElement('div');
    header.style.cssText = 'padding:10px 16px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid rgba(43,104,224,0.3);flex-shrink:0';
    if (isLight()) header.style.borderBottomColor = '#d0dff0';
    header.innerHTML = '<span style="font-weight:600;color:' + (isLight()?'#1a2a4a':'#e8f4ff') + '">地图全屏视图</span>';
    var closeSpan = document.createElement('span');
    closeSpan.style.cssText = 'cursor:pointer;font-size:20px;color:' + (isLight()?'#6080a0':'#a0cfff');
    closeSpan.innerHTML = '&times;';
    closeSpan.onclick = function() {
        m.overlay.remove();
        // 恢复原地图
        setTimeout(function() {
            var origMap = window['_mapObj_'+mapId];
            if (origMap && origMap.invalidateSize) origMap.invalidateSize();
        }, 100);
    };
    header.appendChild(closeSpan);

    var mapContainer = document.createElement('div');
    mapContainer.id = '__fullscreenMapContainer';
    mapContainer.style.cssText = 'flex:1;width:100%;min-height:0';

    var innerMap = document.createElement('div');
    innerMap.id = '__fullscreenMap';
    innerMap.style.cssText = 'width:100%;height:100%';
    mapContainer.appendChild(innerMap);

    m.box.appendChild(header);
    m.box.appendChild(mapContainer);

    // 等待DOM渲染后初始化地图
    setTimeout(function() {
        if (typeof L === 'undefined') return;
        var origMapObj = mapObj || window['_mapObj_'+mapId] || window._mapTrajectoryObj;
        var center = origMapObj ? origMapObj.getCenter() : [31.05, 103.95];
        var zoom = origMapObj ? origMapObj.getZoom() : 12;

        var fsMap = L.map('__fullscreenMap', { center: center, zoom: zoom, zoomControl: true });
        L.tileLayer('https://webrd0{s}.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}', {
            subdomains: ['1','2','3','4'], maxZoom: 18, attribution: '&copy; 高德地图'
        }).addTo(fsMap);

        // 复制原地图的图层
        if (origMapObj) {
            origMapObj.eachLayer(function(layer) {
                if (layer instanceof L.TileLayer) return;
                try { var clone = layer; fsMap.addLayer(clone); } catch(e) {}
            });
        }
        setTimeout(function() { fsMap.invalidateSize(); }, 200);
    }, 300);
};

/* ─────────────────── 11. 统一初始化入口 ─────────────────── */
function initAll() {
    // 分页
    document.querySelectorAll('.card').forEach(function(card) {
        if (card.querySelector('.data-table') && !card.dataset.pagiInited) {
            card.dataset.pagiInited = '1';
            initOnePagination(card);
        }
    });
    // 查询
    initSearchButtons();
    // 新增
    initAddButtons();
    // 删除
    initDeleteButtons();
    // 地图控制
    setTimeout(function() {
        var mapObjMap = {
            'mapLocation':'_mapLocationObj','mapTrip':'_mapTripObj',
            'mapTrajectory':'_mapTrajectoryObj','mapBoundary':'_mapBoundaryObj',
            'mapFixed':'_mapFixedObj','mapInTransit':'_mapInTransitObj'
        };
        Object.keys(mapObjMap).forEach(function(mid) {
            var el = document.getElementById(mid);
            if (el && el.classList.contains('leaflet-container')) {
                var mo = window[mapObjMap[mid]];
                initMapControls(mid, mo);
            }
        });
    }, 1500);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initAll);
} else {
    setTimeout(initAll, 0);
}

// 提供给各页面手动触发重新初始化
window._reinitPagination = function() {
    document.querySelectorAll('.card').forEach(function(card) {
        delete card.dataset.pagiInited;
    });
    initAll();
};

/* ─────────────────── 12. 图表统一主题 ─────────────────── */
window._chartTheme = {
    dark: {
        bg: 'transparent',
        textColor: '#a0cfff',
        titleColor: '#e8f4ff',
        axisLineColor: 'rgba(64,196,255,0.2)',
        splitLineColor: 'rgba(64,196,255,0.08)',
        legendColor: '#a0cfff',
        tooltip: { bg: 'rgba(10,20,60,0.92)', border: 'rgba(64,196,255,0.3)', textColor: '#e8f4ff' },
        palette: ['#40c4ff','#4cd964','#ff9500','#ff4d4f','#7b61ff','#00e5c0','#ffcc00','#ff6b9d']
    },
    light: {
        bg: 'transparent',
        textColor: '#4a6880',
        titleColor: '#1a2a4a',
        axisLineColor: 'rgba(100,150,220,0.3)',
        splitLineColor: 'rgba(100,150,220,0.12)',
        legendColor: '#4a6880',
        tooltip: { bg: 'rgba(255,255,255,0.96)', border: '#d0dff0', textColor: '#1a2a4a' },
        palette: ['#1677ff','#52c41a','#fa8c16','#ff4d4f','#722ed1','#13c2c2','#fadb14','#eb2f96']
    }
};

window.getChartTheme = function() {
    return isLight() ? window._chartTheme.light : window._chartTheme.dark;
};

window.applyChartTheme = function(chart, options) {
    if (!chart || !chart.setOption) return;
    var theme = window.getChartTheme();
    var patch = {
        backgroundColor: theme.bg,
        color: theme.palette,
        legend: { textStyle: { color: theme.legendColor, fontSize: 12 } },
        tooltip: {
            backgroundColor: theme.tooltip.bg,
            borderColor: theme.tooltip.border,
            textStyle: { color: theme.tooltip.textColor, fontSize: 13 }
        },
        xAxis: options && options.noAxis ? undefined : {
            axisLine: { lineStyle: { color: theme.axisLineColor } },
            axisLabel: { color: theme.textColor, fontSize: 12 },
            splitLine: { lineStyle: { color: theme.splitLineColor } }
        },
        yAxis: options && options.noAxis ? undefined : {
            axisLine: { lineStyle: { color: theme.axisLineColor } },
            axisLabel: { color: theme.textColor, fontSize: 12 },
            splitLine: { lineStyle: { color: theme.splitLineColor } }
        }
    };
    chart.setOption(patch, false);
};

/* ─────────────────── CSS 动画 ─────────────────── */
var style = document.createElement('style');
style.textContent = '@keyframes fadeOut{to{opacity:0;transform:translateY(-8px)}}';
document.head.appendChild(style);

/* ─────────────────── 13. ECharts 图表全局主题注入 ─────────────────── */
// 在 ECharts 加载完成后，拦截 echarts.init，自动应用统一主题
function patchECharts() {
    if (!window.echarts) return;
    if (window.echarts.__themePatched) return;
    window.echarts.__themePatched = true;

    var origInit = window.echarts.init;
    window.echarts.init = function(dom, theme, opts) {
        var chart = origInit.call(window.echarts, dom, theme, opts);
        var origSetOption = chart.setOption.bind(chart);

        chart.setOption = function(option, notMerge, lazyUpdate) {
            // 注入统一主题配置
            var t = window.getChartTheme ? window.getChartTheme() : window._chartTheme.dark;
            
            // 全局字体
            if (!option.textStyle) option.textStyle = {};
            option.textStyle.color = option.textStyle.color || t.textColor;
            option.textStyle.fontSize = option.textStyle.fontSize || 12;

            // tooltip
            if (option.tooltip) {
                option.tooltip.backgroundColor = option.tooltip.backgroundColor || t.tooltip.bg;
                option.tooltip.borderColor = option.tooltip.borderColor || t.tooltip.border;
                if (!option.tooltip.textStyle) option.tooltip.textStyle = {};
                option.tooltip.textStyle.color = option.tooltip.textStyle.color || t.tooltip.textColor;
                option.tooltip.textStyle.fontSize = option.tooltip.textStyle.fontSize || 13;
            }

            // legend
            if (option.legend) {
                var legends = Array.isArray(option.legend) ? option.legend : [option.legend];
                legends.forEach(function(lg) {
                    if (!lg.textStyle) lg.textStyle = {};
                    lg.textStyle.color = lg.textStyle.color || t.legendColor;
                    lg.textStyle.fontSize = lg.textStyle.fontSize || 12;
                });
            }

            // xAxis
            if (option.xAxis) {
                var xAxes = Array.isArray(option.xAxis) ? option.xAxis : [option.xAxis];
                xAxes.forEach(function(ax) {
                    if (!ax.axisLabel) ax.axisLabel = {};
                    ax.axisLabel.color = ax.axisLabel.color || t.textColor;
                    ax.axisLabel.fontSize = ax.axisLabel.fontSize || 11;
                    if (!ax.axisLine) ax.axisLine = {};
                    if (!ax.axisLine.lineStyle) ax.axisLine.lineStyle = {};
                    ax.axisLine.lineStyle.color = ax.axisLine.lineStyle.color || t.axisLineColor;
                    if (!ax.splitLine) ax.splitLine = {};
                    if (!ax.splitLine.lineStyle) ax.splitLine.lineStyle = {};
                    ax.splitLine.lineStyle.color = ax.splitLine.lineStyle.color || t.splitLineColor;
                    if (ax.nameTextStyle) ax.nameTextStyle.color = ax.nameTextStyle.color || t.textColor;
                });
            }

            // yAxis
            if (option.yAxis) {
                var yAxes = Array.isArray(option.yAxis) ? option.yAxis : [option.yAxis];
                yAxes.forEach(function(ax) {
                    if (!ax.axisLabel) ax.axisLabel = {};
                    ax.axisLabel.color = ax.axisLabel.color || t.textColor;
                    ax.axisLabel.fontSize = ax.axisLabel.fontSize || 11;
                    if (!ax.axisLine) ax.axisLine = {};
                    if (!ax.axisLine.lineStyle) ax.axisLine.lineStyle = {};
                    ax.axisLine.lineStyle.color = ax.axisLine.lineStyle.color || t.axisLineColor;
                    if (!ax.splitLine) ax.splitLine = {};
                    if (!ax.splitLine.lineStyle) ax.splitLine.lineStyle = {};
                    ax.splitLine.lineStyle.color = ax.splitLine.lineStyle.color || t.splitLineColor;
                    if (ax.nameTextStyle) ax.nameTextStyle.color = ax.nameTextStyle.color || t.textColor;
                });
            }

            return origSetOption(option, notMerge, lazyUpdate);
        };

        return chart;
    };
}

// 轮询等待 echarts 加载
var _echartsCheckCount = 0;
function waitForECharts() {
    if (window.echarts) {
        patchECharts();
    } else if (_echartsCheckCount < 20) {
        _echartsCheckCount++;
        setTimeout(waitForECharts, 300);
    }
}
setTimeout(waitForECharts, 200);

// 主题切换时重绘所有图表
var origToggle = window.toggleTheme;
window.toggleTheme = function() {
    if (origToggle) origToggle.apply(this, arguments);
    // 延迟后对所有图表重绘
    setTimeout(function() {
        if (!window.echarts) return;
        document.querySelectorAll('[_echarts_instance_]').forEach(function(dom) {
            try {
                var chart = window.echarts.getInstanceByDom(dom);
                if (chart) {
                    var t = window.getChartTheme();
                    // 重设axis颜色
                    chart.setOption({
                        legend: { textStyle: { color: t.legendColor } },
                        tooltip: { backgroundColor: t.tooltip.bg, borderColor: t.tooltip.border, textStyle: { color: t.tooltip.textColor } }
                    }, false);
                }
            } catch(e) {}
        });
    }, 150);
};

})(window);



/* ═══════════════════════════════════════════════════════════════
   PHASE 2: 扩展功能模块 (v20260315)
   ═══════════════════════════════════════════════════════════════ */

/* ─── A. 彭州市真实边界数据 ─── */
window._pengzhouBoundary = [[[[31.0318, 104.10644], [31.03069, 104.10587], [31.03276, 104.0949], [31.03591, 104.09155], [31.03755, 104.08679], [31.0394, 104.08729], [31.04219, 104.08639], [31.04464, 104.08449], [31.04782, 104.08105], [31.05011, 104.07783], [31.05241, 104.07905], [31.05755, 104.07882], [31.05939, 104.07647], [31.06197, 104.07138], [31.06413, 104.06056], [31.06624, 104.05912], [31.07476, 104.05999], [31.07436, 104.05828], [31.07818, 104.05761], [31.07959, 104.05065], [31.0861, 104.0433], [31.08881, 104.04278], [31.0907, 104.04382], [31.09244, 104.03868], [31.09387, 104.03411], [31.09624, 104.03142], [31.09773, 104.0289], [31.10007, 104.03007], [31.0998, 104.02711], [31.1023, 104.02903], [31.10356, 104.02765], [31.10485, 104.02544], [31.10552, 104.02023], [31.10867, 104.02384], [31.11109, 104.03027], [31.11213, 104.02746], [31.11374, 104.02402], [31.11343, 104.02139], [31.11483, 104.01729], [31.11829, 104.02127], [31.12302, 104.02479], [31.12588, 104.02315], [31.12789, 104.02406], [31.12863, 104.02234], [31.13022, 104.02281], [31.13379, 104.02324], [31.13747, 104.0204], [31.14268, 104.02287], [31.14737, 104.01717], [31.14791, 104.01413], [31.14994, 104.01292], [31.15083, 104.01108], [31.15602, 104.00956], [31.15867, 104.00956], [31.16051, 104.01091], [31.16391, 104.0101], [31.16554, 104.01498], [31.16987, 104.01252], [31.17376, 104.01024], [31.1785, 104.00817], [31.17743, 104.00661], [31.17726, 104.00204], [31.1813, 103.99916], [31.18638, 103.99713], [31.18941, 103.99409], [31.19102, 103.98851], [31.19415, 103.98807], [31.19778, 103.98412], [31.20094, 103.98298], [31.20066, 103.97989], [31.20175, 103.97659], [31.20332, 103.97472], [31.20488, 103.97386], [31.20731, 103.97143], [31.20642, 103.96777], [31.20819, 103.96352], [31.206, 103.96156], [31.20297, 103.9573], [31.20426, 103.95449], [31.20349, 103.94861], [31.20696, 103.94649], [31.21043, 103.9485], [31.21511, 103.95032], [31.21539, 103.94659], [31.21635, 103.94376], [31.21794, 103.94135], [31.2183, 103.94578], [31.22127, 103.94327], [31.22318, 103.93929], [31.22317, 103.93502], [31.22997, 103.92998], [31.23443, 103.9291], [31.24361, 103.93462], [31.24756, 103.93207], [31.25362, 103.92678], [31.25826, 103.92597], [31.26114, 103.92903], [31.26516, 103.93296], [31.27151, 103.92914], [31.27499, 103.92306], [31.27898, 103.91417], [31.28362, 103.91236], [31.28613, 103.90727], [31.29364, 103.90278], [31.29787, 103.8961], [31.30377, 103.89222], [31.31598, 103.89193], [31.3228, 103.89297], [31.32927, 103.89736], [31.3396, 103.9082], [31.34386, 103.91427], [31.34582, 103.92107], [31.34755, 103.92534], [31.35243, 103.92775], [31.35741, 103.928], [31.36038, 103.92538], [31.36439, 103.91969], [31.36954, 103.92081], [31.37297, 103.9184], [31.37971, 103.91038], [31.38356, 103.90788], [31.38973, 103.90732], [31.39641, 103.90273], [31.40377, 103.8976], [31.40888, 103.89475], [31.41524, 103.89745], [31.41993, 103.89456], [31.42101, 103.88996], [31.41573, 103.88369], [31.41498, 103.87602], [31.41899, 103.87068], [31.42152, 103.86694], [31.41733, 103.86075], [31.41382, 103.85343], [31.41291, 103.84457], [31.41441, 103.83703], [31.41136, 103.83164], [31.41245, 103.82864], [31.42265, 103.81747], [31.43193, 103.81604], [31.43474, 103.81089], [31.43507, 103.80631], [31.43107, 103.80411], [31.43004, 103.80053], [31.43176, 103.79733], [31.43156, 103.79237], [31.43409, 103.78896], [31.43406, 103.78574], [31.43121, 103.77809], [31.42616, 103.76961], [31.42382, 103.76015], [31.41916, 103.75029], [31.41437, 103.74508], [31.40494, 103.74308], [31.40035, 103.74143], [31.3981, 103.73757], [31.39718, 103.73383], [31.39452, 103.72709], [31.3927, 103.71741], [31.39194, 103.71279], [31.39148, 103.70142], [31.38655, 103.69957], [31.38229, 103.69624], [31.37797, 103.69179], [31.37358, 103.68431], [31.36915, 103.68012], [31.36251, 103.68404], [31.36007, 103.68713], [31.35305, 103.69151], [31.34877, 103.69695], [31.33821, 103.69453], [31.33251, 103.69864], [31.32438, 103.70061], [31.32052, 103.70845], [31.31778, 103.711], [31.31103, 103.71371], [31.3069, 103.71253], [31.30281, 103.70874], [31.30024, 103.70955], [31.29613, 103.71378], [31.28937, 103.71575], [31.28435, 103.72172], [31.27994, 103.72253], [31.26785, 103.72119], [31.25928, 103.71349], [31.25444, 103.7073], [31.24983, 103.70338], [31.24297, 103.69566], [31.23598, 103.69397], [31.22851, 103.68866], [31.21903, 103.68477], [31.20992, 103.68461], [31.2029, 103.69032], [31.19535, 103.69563], [31.18883, 103.69673], [31.18633, 103.70149], [31.1856, 103.70603], [31.17567, 103.71469], [31.16836, 103.72072], [31.16557, 103.72774], [31.16393, 103.73362], [31.16424, 103.73955], [31.16041, 103.73989], [31.14932, 103.73494], [31.14578, 103.73503], [31.13926, 103.74085], [31.1333, 103.74109], [31.12705, 103.74184], [31.11954, 103.74494], [31.1131, 103.74966], [31.1119, 103.75519], [31.10633, 103.75825], [31.10203, 103.76367], [31.09126, 103.7629], [31.08519, 103.76709], [31.08386, 103.77166], [31.08433, 103.78401], [31.08, 103.78366], [31.07949, 103.77135], [31.07632, 103.75991], [31.07607, 103.75431], [31.07075, 103.74395], [31.07166, 103.73053], [31.07124, 103.72768], [31.06599, 103.72795], [31.05959, 103.74014], [31.05364, 103.74075], [31.03497, 103.74708], [31.03359, 103.75051], [31.02546, 103.75382], [31.01489, 103.75026], [31.0083, 103.75663], [31.00389, 103.76131], [31.00013, 103.76645], [30.98991, 103.77162], [30.98269, 103.77426], [30.98103, 103.77669], [30.98, 103.78259], [30.97475, 103.78028], [30.95785, 103.77906], [30.95543, 103.78506], [30.9503, 103.78948], [30.94878, 103.7933], [30.94547, 103.79561], [30.94302, 103.80146], [30.9448, 103.80484], [30.94494, 103.81239], [30.94308, 103.81699], [30.94089, 103.82114], [30.93865, 103.82441], [30.93457, 103.8263], [30.92938, 103.82871], [30.9278, 103.8339], [30.92725, 103.84017], [30.92917, 103.84295], [30.93445, 103.8441], [30.93671, 103.84758], [30.93564, 103.85436], [30.93759, 103.85938], [30.93764, 103.86289], [30.93447, 103.868], [30.93079, 103.87413], [30.92998, 103.885], [30.92708, 103.88943], [30.92645, 103.89805], [30.92651, 103.90505], [30.92266, 103.91082], [30.92205, 103.91735], [30.92288, 103.92658], [30.92172, 103.94082], [30.91791, 103.94704], [30.9156, 103.95229], [30.91921, 103.9537], [30.91896, 103.95642], [30.9169, 103.96169], [30.92003, 103.96438], [30.92827, 103.96457], [30.93127, 103.96991], [30.93629, 103.97621], [30.93727, 103.98323], [30.93933, 103.99049], [30.93803, 103.99391], [30.93737, 103.99784], [30.94059, 104.00059], [30.9407, 104.00336], [30.94134, 104.00972], [30.94353, 104.01261], [30.94654, 104.0142], [30.95189, 104.02331], [30.95782, 104.02927], [30.95806, 104.03674], [30.96007, 104.04109], [30.96443, 104.04488], [30.9648, 104.04884], [30.96248, 104.0505], [30.9585, 104.05125], [30.95219, 104.06094], [30.94842, 104.06342], [30.94312, 104.06354], [30.93513, 104.0679], [30.93022, 104.06774], [30.92522, 104.06733], [30.92088, 104.07148], [30.91639, 104.07842], [30.91489, 104.09131], [30.91286, 104.10149], [30.91393, 104.11667], [30.91154, 104.12104], [30.91378, 104.12955], [30.91473, 104.1379], [30.91577, 104.14445], [30.91309, 104.15002], [30.92212, 104.14666], [30.92503, 104.14256], [30.92733, 104.1448], [30.92946, 104.15275], [30.9352, 104.15301], [30.94012, 104.15249], [30.94239, 104.15421], [30.9431, 104.15825], [30.94896, 104.1604], [30.95035, 104.16255], [30.95128, 104.16635], [30.95478, 104.16639], [30.95997, 104.16381], [30.96307, 104.16767], [30.96816, 104.16849], [30.97461, 104.16963], [30.97857, 104.16833], [30.98287, 104.16568], [30.98617, 104.16945], [30.99182, 104.16805], [30.99783, 104.16188], [31.00105, 104.15947], [31.00502, 104.15446], [31.0027, 104.15145], [31.00365, 104.14791], [31.0056, 104.14934], [31.00604, 104.14532], [31.00873, 104.14433], [31.01158, 104.14342], [31.0124, 104.13789], [31.01633, 104.12589], [31.01815, 104.12279], [31.0173, 104.11964], [31.02003, 104.11597], [31.02508, 104.11727], [31.02664, 104.12204], [31.02989, 104.12179], [31.02999, 104.11072], [31.03214, 104.10949], [31.0334, 104.10908], [31.03228, 104.10702], [31.03162, 104.10843], [31.03164, 104.10658], [31.0318, 104.10644]]], [[[31.032, 104.10714], [31.0318, 104.10644], [31.032, 104.10714]]]];

/* ─── B. 绘制彭州市边界 ─── */
window.drawPengzhouBoundary = function(map) {
    if (!map || !window.L || !window._pengzhouBoundary) return null;
    var layers = [];
    window._pengzhouBoundary.forEach(function(poly) {
        var layer = L.polygon(poly, {
            color: '#40c4ff', weight: 2, opacity: 0.9,
            fillColor: '#40c4ff', fillOpacity: 0.04, dashArray: '6,4'
        }).addTo(map);
        layers.push(layer);
    });
    return layers;
};

/* ─── C. 全屏地图（增强版，代入所有图层+边界）─── */
window.openFullscreenMap = function(mapId, mapObj) {
    var m = createModalBase('__mapFullscreen');
    var lt = isLight();
    m.box.style.cssText = 'width:96vw;height:92vh;max-width:96vw;max-height:92vh;background:' + (lt ? '#f0f4f8' : '#0a1024') + ';border:1px solid ' + (lt ? '#d0dff0' : 'rgba(43,104,224,0.4)') + ';border-radius:12px;overflow:hidden;display:flex;flex-direction:column';

    var header = document.createElement('div');
    header.style.cssText = 'padding:10px 16px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid ' + (lt ? '#d0dff0' : 'rgba(43,104,224,0.3)') + ';flex-shrink:0;background:' + (lt ? 'rgba(240,244,248,0.9)' : 'rgba(10,16,40,0.9)');
    header.innerHTML = '<span style="font-weight:600;font-size:15px;color:' + (lt ? '#1a2a4a' : '#e8f4ff') + '">\uD83D\uDCCD 彭州市砂石监管 · 地图全屏视图</span>';

    var ctrlGroup = document.createElement('div');
    ctrlGroup.style.cssText = 'display:flex;align-items:center;gap:8px';

    var centerBtn2 = document.createElement('button');
    centerBtn2.className = 'btn btn-default';
    centerBtn2.innerHTML = '<i class="fas fa-crosshairs"></i> 居中';
    centerBtn2.style.cssText = 'padding:5px 10px;font-size:12px;height:30px';

    var closeSpan = document.createElement('span');
    closeSpan.style.cssText = 'cursor:pointer;font-size:22px;color:' + (lt ? '#6080a0' : '#a0cfff') + ';line-height:1;padding:0 4px;border-radius:4px';
    closeSpan.innerHTML = '&times;';
    closeSpan.onclick = function() {
        m.overlay.remove();
        setTimeout(function() {
            var om = mapObj || window['_mapObj_' + mapId];
            if (om && om.invalidateSize) om.invalidateSize();
        }, 100);
    };

    ctrlGroup.appendChild(centerBtn2);
    ctrlGroup.appendChild(closeSpan);
    header.appendChild(ctrlGroup);

    var mapContainer = document.createElement('div');
    mapContainer.style.cssText = 'flex:1;width:100%;min-height:0;position:relative';
    var innerMap = document.createElement('div');
    innerMap.id = '__fullscreenMap';
    innerMap.style.cssText = 'width:100%;height:100%';
    mapContainer.appendChild(innerMap);

    m.box.appendChild(header);
    m.box.appendChild(mapContainer);

    setTimeout(function() {
        if (typeof L === 'undefined') {
            innerMap.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:200px;color:' + (lt ? '#4a6880' : '#a0cfff') + ';font-size:14px">地图库加载中，请稍候...</div>';
            return;
        }
        var origMapObj = mapObj || window['_mapObj_' + mapId] || window._mapTrajectoryObj || window._mapLocationObj || window._mapTripObj;
        var center = [31.05, 103.95], zoom = 11;
        if (origMapObj && origMapObj.getCenter) {
            try { center = [origMapObj.getCenter().lat, origMapObj.getCenter().lng]; zoom = origMapObj.getZoom(); } catch(e2) {}
        }
        var fsMap = L.map('__fullscreenMap', { center: center, zoom: zoom, zoomControl: true });
        L.tileLayer('https://webrd0{s}.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}', {
            subdomains: ['1','2','3','4'], maxZoom: 18, attribution: '&copy; \u9ad8\u5fb7\u5730\u56fe'
        }).addTo(fsMap);
        window.drawPengzhouBoundary(fsMap);
        if (origMapObj) {
            try {
                origMapObj.eachLayer(function(layer) {
                    if (layer instanceof L.TileLayer) return;
                    try {
                        if (layer instanceof L.Marker) {
                            var lm = L.marker(layer.getLatLng(), {});
                            if (layer._popup) lm.bindPopup(layer._popup._content);
                            lm.addTo(fsMap);
                        } else if (layer instanceof L.CircleMarker) {
                            var lc = L.circleMarker(layer.getLatLng(), layer.options);
                            if (layer._popup) lc.bindPopup(layer._popup._content);
                            lc.addTo(fsMap);
                        } else if (layer instanceof L.Polyline && !(layer instanceof L.Polygon)) {
                            L.polyline(layer.getLatLngs(), layer.options).addTo(fsMap);
                        }
                    } catch(le) {}
                });
            } catch(e3) {}
        }
        centerBtn2.onclick = function() { fsMap.setView([31.05, 103.95], 11); };
        setTimeout(function() { fsMap.invalidateSize(); }, 200);
    }, 300);
};

/* ─── D. 地图控制（居中+全屏，增强版）─── */
window.initMapControls = function(mapId, mapObj) {
    var el = document.getElementById(mapId);
    if (!el) return;
    var container = el.parentElement;
    if (!container) return;
    var old = container.querySelector('.map-ctrl-btns');
    if (old) old.remove();

    var ctrlDiv = document.createElement('div');
    ctrlDiv.className = 'map-ctrl-btns';
    ctrlDiv.style.cssText = 'position:absolute;top:10px;right:10px;z-index:1000;display:flex;flex-direction:column;gap:6px';

    var lt = isLight();
    var btnBase = 'width:34px;height:34px;border-radius:6px;border:1px solid ' + (lt ? 'rgba(100,150,220,0.4)' : 'rgba(64,196,255,0.3)') + ';background:' + (lt ? 'rgba(255,255,255,0.92)' : 'rgba(13,22,51,0.85)') + ';backdrop-filter:blur(8px);cursor:pointer;font-size:15px;display:flex;align-items:center;justify-content:center;color:' + (lt ? '#4a6880' : '#a0cfff') + ';transition:all 0.2s;padding:0;';

    var centerBtn = document.createElement('button');
    centerBtn.style.cssText = btnBase;
    centerBtn.title = '\u5c45\u4e2d\u5b9a\u4f4d\u81f3\u5f6d\u5dde\u5e02';
    centerBtn.innerHTML = '\u2295';
    centerBtn.onclick = function() {
        var mo = mapObj || window['_mapObj_' + mapId] || window._mapTrajectoryObj || window._mapLocationObj || window._mapTripObj;
        if (mo && mo.setView) { mo.setView([31.05, 103.95], 11); showToast('\u5df2\u5c45\u4e2d\u81f3\u5f6d\u5dde\u5e02', 'info'); }
    };

    var fsBtn = document.createElement('button');
    fsBtn.style.cssText = btnBase;
    fsBtn.title = '\u5168\u5c4f\u67e5\u770b\u5730\u56fe';
    fsBtn.innerHTML = '\u26f6';
    fsBtn.onclick = function() { window.openFullscreenMap(mapId, mapObj); };

    ctrlDiv.appendChild(centerBtn);
    ctrlDiv.appendChild(fsBtn);
    if (!['absolute','fixed','relative'].includes(container.style.position)) container.style.position = 'relative';
    container.appendChild(ctrlDiv);
    if (mapObj) window['_mapObj_' + mapId] = mapObj;

    setTimeout(function() {
        var mo = mapObj || window['_mapObj_' + mapId];
        if (mo && mo.addLayer && !mo._pengzhouBoundaryDrawn) {
            mo._pengzhouBoundaryDrawn = true;
            window.drawPengzhouBoundary(mo);
        }
    }, 500);
};

/* ─── E. KPI卡片点击筛选 ─── */
(function initKpiClickFilter() {
    function applyFilter(filterVal) {
        var applied = false;
        document.querySelectorAll('.data-table tbody').forEach(function(tbody) {
            tbody.querySelectorAll('tr').forEach(function(tr) {
                if (!filterVal) { tr.style.display = ''; applied = true; }
                else { tr.style.display = tr.textContent.toLowerCase().includes(filterVal.toLowerCase()) ? '' : 'none'; applied = true; }
            });
        });
        if (applied) showToast(filterVal ? ('\u5df2\u7b5b\u9009\uff1a' + filterVal) : '\u5df2\u663e\u793a\u5168\u90e8', 'info');
    }
    function bindKpi() {
        document.querySelectorAll('.kpi-card').forEach(function(card) {
            if (card.dataset.kpiInited) return;
            card.dataset.kpiInited = '1';
            card.style.cursor = 'pointer';
            card.style.transition = 'all 0.2s';
            card.addEventListener('mouseenter', function() { card.style.transform = 'translateY(-2px)'; card.style.boxShadow = '0 4px 20px rgba(64,196,255,0.25)'; });
            card.addEventListener('mouseleave', function() { card.style.transform = ''; card.style.boxShadow = ''; });
            card.addEventListener('click', function() {
                var link = card.dataset.link;
                if (link) { window.location.href = link; return; }
                var filterAttr = card.dataset.filter;
                if (filterAttr) { applyFilter(filterAttr); return; }
                var label = (card.querySelector('.kpi-label') || {}).textContent || '';
                var valEl = card.querySelector('.kpi-value');
                var color = valEl ? (valEl.style.color || '') : '';
                var colorMap = { '#ff4d4f': '\u5f02\u5e38', '#ff9500': '\u8b66\u793a', '#4cd964': '\u6b63\u5e38' };
                if (colorMap[color]) { applyFilter(colorMap[color]); return; }
                if (label.includes('\u5f02\u5e38') || label.includes('\u8fdd\u89c4')) { applyFilter('\u5f02\u5e38'); return; }
                if (label.includes('\u4e34\u671f') || label.includes('\u8b66\u793a')) { applyFilter('\u4e34\u671f'); return; }
                applyFilter(null);
                showToast('\u5c55\u793a\u5168\u90e8\u6570\u636e', 'info');
            });
        });
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', function() { setTimeout(bindKpi, 200); });
    else setTimeout(bindKpi, 200);
    window._bindKpiCards = bindKpi;
})();

/* ─── F. 操作按钮弹窗（处置/凭证/打印/告警/详情/编辑）─── */
(function initOperationBtns() {
    function showDisposalModal(row) {
        createFormModal('\u26a0\ufe0f \u8fdd\u89c4\u5904\u7f6e', [
            {name:'type', label:'\u5904\u7f6e\u7c7b\u578b', type:'select', options:['\u8bf7\u9009\u62e9','\u8b66\u544a\u901a\u62a5','\u6682\u505c\u4e1a\u52a1','\u7f5a\u6b3e\u5904\u7406','\u540a\u9500\u8d44\u8d28','\u79fb\u4ea4\u516c\u5b89'], required:true},
            {name:'reason', label:'\u5904\u7f6e\u539f\u56e0', type:'textarea', placeholder:'\u8bf7\u8be6\u7ec6\u63cf\u8ff0...', required:true},
            {name:'operator', label:'\u5904\u7f6e\u4eba\u5458', type:'text', placeholder:'\u8bf7\u8f93\u5165\u5904\u7f6e\u4eba\u59d3\u540d', required:true},
            {name:'deadline', label:'\u6574\u6539\u671f\u9650', type:'date'}
        ], function() {
            showToast('\u5904\u7f6e\u8bb0\u5f55\u5df2\u63d0\u4ea4', 'success');
            var s = row && row.querySelector('.status-tag,.status-badge');
            if (s) { s.textContent = '\u5904\u7f6e\u4e2d'; s.style.color = '#ff9500'; }
        });
    }

    function showVoucherModal(row) {
        var lt = isLight();
        var m = createModalBase('__voucher');
        m.box.style.cssText = m.box.style.cssText + ';min-width:520px';
        var hdr = document.createElement('div'); hdr.style.cssText = m.theme.header;
        hdr.innerHTML = '<span style="' + m.theme.title + '">\uD83D\uDCCB \u4e1a\u52a1\u51ed\u8bc1</span>';
        var cl = document.createElement('span'); cl.style.cssText = m.theme.close; cl.innerHTML = '&times;'; cl.onclick = function() { m.overlay.remove(); }; hdr.appendChild(cl);
        var body = document.createElement('div'); body.style.cssText = m.theme.body;
        var cells = row ? Array.from(row.querySelectorAll('td')) : [];
        var vno = 'PZ' + Date.now().toString().slice(-8);
        body.innerHTML = '<div style="border:1px solid ' + (lt?'#d0dff0':'rgba(64,196,255,0.2)') + ';border-radius:8px;padding:20px;text-align:center"><div style="font-size:17px;font-weight:700;color:' + (lt?'#1a2a4a':'#e8f4ff') + ';margin-bottom:6px">\u5f6d\u5dde\u5e02\u7802\u77f3\u8d44\u6e90\u667a\u6167\u76d1\u7ba1\u5e73\u53f0</div><div style="color:' + (lt?'#6080a0':'#a0cfff') + ';font-size:13px;margin-bottom:14px">\u4e1a\u52a1\u64cd\u4f5c\u51ed\u8bc1</div><div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;text-align:left;font-size:13px;margin-bottom:14px"><div><span style="color:' + (lt?'#6080a0':'#80a0c0') + '">\u51ed\u8bc1\u7f16\u53f7\uff1a</span><b>' + vno + '</b></div><div><span style="color:' + (lt?'#6080a0':'#80a0c0') + '">\u751f\u6210\u65f6\u95f4\uff1a</span>' + new Date().toLocaleString('zh-CN') + '</div><div><span style="color:' + (lt?'#6080a0':'#80a0c0') + '">\u4e1a\u52a1\u4e3b\u4f53\uff1a</span>' + (cells[0]?cells[0].textContent.trim():'-') + '</div><div><span style="color:' + (lt?'#6080a0':'#80a0c0') + '">\u4e1a\u52a1\u5185\u5bb9\uff1a</span>' + (cells[1]?cells[1].textContent.trim():'-') + '</div></div><div style="font-size:12px;color:' + (lt?'#6080a0':'#80a0c0') + '">\u672c\u51ed\u8bc1\u7531\u7cfb\u7edf\u81ea\u52a8\u751f\u6210\uff0c\u5177\u6709\u6cd5\u5f8b\u6548\u529b</div></div>';
        var ftr = document.createElement('div'); ftr.style.cssText = m.theme.footer;
        var pb = document.createElement('button'); pb.className = 'btn btn-primary'; pb.innerHTML = '<i class="fas fa-print"></i> \u6253\u5370'; pb.onclick = function() { window.print(); };
        var cb2 = document.createElement('button'); cb2.className = 'btn btn-default'; cb2.textContent = '\u5173\u95ed'; cb2.onclick = function() { m.overlay.remove(); };
        ftr.appendChild(pb); ftr.appendChild(cb2);
        m.box.appendChild(hdr); m.box.appendChild(body); m.box.appendChild(ftr);
    }

    function showAlertSubmit(row) {
        var cells = row ? Array.from(row.querySelectorAll('td')) : [];
        var subj = cells[0] ? cells[0].textContent.trim() : '\u672a\u77e5\u4e3b\u4f53';
        createFormModal('\uD83D\uDEA8 \u751f\u6210\u544a\u8b66', [
            {name:'title', label:'\u544a\u8b66\u6807\u9898', type:'text', value: subj + ' - \u8fdd\u89c4\u544a\u8b66', required:true},
            {name:'level', label:'\u544a\u8b66\u7ea7\u522b', type:'select', options:['\u8bf7\u9009\u62e9','\u4e00\u822c','\u8b66\u793a','\u4e25\u91cd','\u7d27\u6025'], required:true},
            {name:'type', label:'\u544a\u8b66\u7c7b\u578b', type:'select', options:['\u8bf7\u9009\u62e9','\u8d8a\u754c\u884c\u4e3a','\u8d85\u901f\u8fdd\u89c4','\u8d85\u8f7d\u8fdd\u89c4','\u8f68\u8ff9\u5f02\u5e38','\u8bc1\u4ef6\u903e\u671f','\u5176\u4ed6'], required:true},
            {name:'desc', label:'\u544a\u8b66\u63cf\u8ff0', type:'textarea', placeholder:'\u8bf7\u8be6\u7ec6\u63cf\u8ff0...', required:true}
        ], function() {
            showToast('\u544a\u8b66\u5df2\u63d0\u4ea4\uff0c\u6b63\u5728\u8df3\u8f6c\u5230\u544a\u8b66\u4e2d\u5fc3...', 'warning', 3000);
            setTimeout(function() { window.location.href = 'alert_center.html'; }, 1500);
        });
    }

    function showDetailModal(row) {
        var lt = isLight();
        var m = createModalBase('__detail');
        m.box.style.cssText = m.box.style.cssText + ';min-width:520px;max-width:680px';
        var hdr = document.createElement('div'); hdr.style.cssText = m.theme.header;
        hdr.innerHTML = '<span style="' + m.theme.title + '">\uD83D\uDCC4 \u8be6\u7ec6\u4fe1\u606f</span>';
        var cl = document.createElement('span'); cl.style.cssText = m.theme.close; cl.innerHTML = '&times;'; cl.onclick = function() { m.overlay.remove(); }; hdr.appendChild(cl);
        var body = document.createElement('div'); body.style.cssText = m.theme.body;
        if (row) {
            var table = row.closest('table');
            var headers = table ? Array.from(table.querySelectorAll('thead th')).map(function(th) { return th.textContent.trim(); }) : [];
            var cells = Array.from(row.querySelectorAll('td'));
            var html = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">';
            cells.forEach(function(td, i) {
                var lbl = headers[i] || ('\u5b57\u6bb5' + (i+1));
                if (lbl === '\u64cd\u4f5c' || lbl === '' || !td.textContent.trim()) return;
                html += '<div style="background:' + (lt?'rgba(240,244,248,0.8)':'rgba(10,20,50,0.6)') + ';border-radius:6px;padding:10px 14px"><div style="font-size:11px;color:' + (lt?'#6080a0':'#80a0c0') + ';margin-bottom:3px">' + lbl + '</div><div style="font-size:13px;font-weight:500;color:' + (lt?'#1a2a4a':'#e8f4ff') + '">' + td.textContent.trim() + '</div></div>';
            });
            html += '</div>';
            body.innerHTML = html;
        }
        var ftr = document.createElement('div'); ftr.style.cssText = m.theme.footer;
        var cb2 = document.createElement('button'); cb2.className = 'btn btn-default'; cb2.textContent = '\u5173\u95ed'; cb2.onclick = function() { m.overlay.remove(); };
        ftr.appendChild(cb2);
        m.box.appendChild(hdr); m.box.appendChild(body); m.box.appendChild(ftr);
    }

    function bindOpBtns() {
        document.querySelectorAll('table td button, .data-table td button').forEach(function(btn) {
            if (btn.dataset.opInited) return;
            btn.dataset.opInited = '1';
            var txt = btn.textContent.trim();
            var action = btn.dataset.action || '';
            var row = btn.closest('tr');
            if (txt.includes('\u5904\u7f6e') || action === 'handle' || action === 'dispose') {
                btn.onclick = function(e) { e.stopPropagation(); showDisposalModal(row); }; return;
            }
            if (txt.includes('\u51ed\u8bc1') || action === 'voucher') {
                btn.onclick = function(e) { e.stopPropagation(); showVoucherModal(row); }; return;
            }
            if (txt.includes('\u6253\u5370') || action === 'print') {
                btn.onclick = function(e) { e.stopPropagation(); window.print(); }; return;
            }
            if (txt.includes('\u544a\u8b66') || action === 'alert') {
                btn.onclick = function(e) { e.stopPropagation(); showAlertSubmit(row); }; return;
            }
            if (txt.includes('\u8be6\u60c5') || txt.includes('\u67e5\u770b') || action === 'view' || action === 'detail') {
                if (!btn.dataset.deleteInited && !btn.getAttribute('onclick')) {
                    btn.onclick = function(e) { e.stopPropagation(); showDetailModal(row); }; return;
                }
            }
            if (txt.includes('\u7f16\u8f91') || action === 'edit') {
                btn.onclick = function(e) {
                    e.stopPropagation();
                    if (!row) return;
                    var tbl = row.closest('table');
                    var hdrs = tbl ? Array.from(tbl.querySelectorAll('thead th')).map(function(th){return th.textContent.trim();}) : [];
                    var tds = Array.from(row.querySelectorAll('td'));
                    var flds = [];
                    tds.forEach(function(td, i) {
                        var h = hdrs[i] || '';
                        if (!h || h === '\u64cd\u4f5c' || td.querySelector('button,.status-tag')) return;
                        flds.push({name:'f'+i, label:h, type:'text', value:td.textContent.trim()});
                    });
                    createFormModal('\u270f\ufe0f \u7f16\u8f91\u8bb0\u5f55', flds, function(data) {
                        var vals = Object.values(data); var j = 0;
                        tds.forEach(function(td, i) {
                            var h = hdrs[i] || '';
                            if (!h || h === '\u64cd\u4f5c' || td.querySelector('button,.status-tag')) return;
                            if (j < vals.length) { td.textContent = vals[j]; j++; }
                        });
                        showToast('\u7f16\u8f91\u4fdd\u5b58\u6210\u529f', 'success');
                    });
                };
            }
        });
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', function() { setTimeout(bindOpBtns, 400); });
    else setTimeout(bindOpBtns, 400);
    window._rebindActionBtns = bindOpBtns;
})();

/* ─── G. 标题栏：全套按钮注入（主题切换 + 清除缓存 + 用户名 + 退出 + 时间）─── */
(function initPageHeaderBtns() {

    /* 注入一次性 CSS（仅在 common.css 已处理的基础上补充 :hover 规则）*/
    function injectCSS() {
        if (document.getElementById('__phdr_css')) return;
        var s = document.createElement('style');
        s.id = '__phdr_css';
        s.textContent = [
            /* 基础：所有注入按钮 */
            '.phdr-btn{',
            '  height:32px;padding:0 13px;border-radius:6px;',
            '  background:rgba(255,255,255,0.10);',
            '  border:1px solid rgba(255,255,255,0.22);',
            '  color:rgba(255,255,255,0.85);',
            '  cursor:pointer;font-size:13px;',
            '  transition:all 0.25s;',
            '  display:inline-flex;align-items:center;gap:6px;',
            '  white-space:nowrap;flex-shrink:0;font-family:inherit;',
            '  box-sizing:border-box;',
            '}',
            /* 通用 hover */
            '.phdr-btn:hover{',
            '  background:rgba(255,255,255,0.20);',
            '  border-color:rgba(255,255,255,0.45);',
            '  color:#fff;',
            '}',
            /* 主题切换：方形图标 */
            '.phdr-btn.phdr-theme{width:32px;padding:0;justify-content:center;font-size:14px;}',
            /* 退出按钮：hover 变红，默认与其他完全一致 */
            '.phdr-btn.phdr-logout:hover{',
            '  background:rgba(255,77,79,0.22);',
            '  border-color:rgba(255,100,102,0.55);',
            '  color:#ff6b6b;',
            '  box-shadow:0 0 10px rgba(255,77,79,0.25);',
            '}',
            /* 时间文字 */
            '.phdr-time{',
            '  font-size:12px;color:rgba(255,255,255,0.55);',
            '  letter-spacing:0.5px;white-space:nowrap;',
            '  padding-left:4px;flex-shrink:0;',
            '}'
        ].join('');
        document.head.appendChild(s);
    }

    /* 获取登录用户名 */
    function getUsername() {
        return sessionStorage.getItem('login_name') || localStorage.getItem('ss_username') || '管理员';
    }

    /* 更新主题切换按钮图标（纯内容，不改 style）*/
    function updateThemeBtn(btn) {
        var lt = document.documentElement.classList.contains('light-theme');
        btn.innerHTML = lt
            ? '<i class="fas fa-moon" style="pointer-events:none"></i>'
            : '<i class="fas fa-sun" style="pointer-events:none"></i>';
        btn.title = lt ? '\u5207\u6362\u6df1\u8272\u6a21\u5f0f' : '\u5207\u6362\u6d45\u8272\u6a21\u5f0f';
    }

    /* 构建主题切换按钮 */
    function buildThemeToggle() {
        var btn = document.createElement('button');
        btn.className = 'phdr-btn phdr-theme';
        btn.onclick = function() { if (window.toggleTheme) window.toggleTheme(); };
        updateThemeBtn(btn);
        return btn;
    }

    /* 构建清除缓存按钮 */
    function buildClearCache() {
        var btn = document.createElement('button');
        btn.className = 'phdr-btn phdr-cache';
        btn.title = '\u4e00\u952e\u6e05\u9664\u6d4f\u89c8\u5668\u7f13\u5b58';
        btn.innerHTML = '<i class="fas fa-trash-alt" style="pointer-events:none"></i> \u6e05\u9664\u7f13\u5b58';
        btn.onclick = function() { if (window.clearCacheNow) window.clearCacheNow(); };
        return btn;
    }

    /* 构建用户名标签 */
    function buildUserTag() {
        var tag = document.createElement('div');
        tag.className = 'phdr-btn phdr-user';
        tag.title = '\u70b9\u51fb\u8fdb\u5165\u7528\u6237\u7ba1\u7406';
        tag.innerHTML = '<i class="fas fa-user-circle" style="pointer-events:none"></i> ' + getUsername();
        tag.onclick = function() { window.location.href = '../pages/user_management.html'; };
        return tag;
    }

    /* 构建退出按钮 */
    function buildLogout() {
        var btn = document.createElement('button');
        btn.className = 'phdr-btn phdr-logout';
        btn.innerHTML = '<i class="fas fa-sign-out-alt" style="pointer-events:none"></i> \u9000\u51fa';
        btn.onclick = function() {
            createConfirmModal('\u786e\u8ba4\u9000\u51fa', '\u786e\u5b9a\u8981\u9000\u51fa\u5f53\u524d\u767b\u5f55\u5417\uff1f', function() {
                window.location.href = '../login.html';
            });
        };
        return btn;
    }

    /* 构建时间显示 */
    function buildTimeDiv() {
        var div = document.createElement('div');
        div.id = 'pageTime';
        div.className = 'phdr-time';
        return div;
    }

    /* ── 标准化左侧面包屑 ──
       标准格式：
         <div class="page-header-left">
           <a class="back-btn" ...>...</a>
           <div class="page-breadcrumb">青川县庄子上工业园区智慧园区平台 / <span>分类 / 页面名</span></div>
         </div>
       处理逻辑：
         1. 读取现有 .page-breadcrumb 的纯文本（去 HTML 标签）
         2. 把 "青川县庄子上工业园区智慧园区平台 / " 后面的部分用 <span> 包裹
         3. 删除 .page-header-left 内除 .back-btn 和 .page-breadcrumb 以外的所有子元素
    */
    function normalizeHeaderLeft() {
        var hl = document.querySelector('.page-header-left');
        if (!hl) return;

        // 1. 找到（或找出）面包屑文字
        var bc = hl.querySelector('.page-breadcrumb');
        var rawText = '';
        if (bc) {
            // 取纯文本，去掉多余空白
            rawText = bc.textContent.replace(/\s+/g, ' ').trim();
        }

        // 2. 删除 .page-header-left 内除 .back-btn 之外的所有子节点
        Array.from(hl.childNodes).forEach(function(node) {
            var isBackBtn = node.nodeType === 1 && node.classList && node.classList.contains('back-btn');
            if (!isBackBtn) hl.removeChild(node);
        });

        // 3. 重建标准面包屑
        var newBc = document.createElement('div');
        newBc.className = 'page-breadcrumb';

        // 分割：固定前缀 + 后半段
        var prefix = '青川县庄子上工业园区智慧园区平台 / ';
        var rest = '';
        if (rawText.indexOf(prefix) === 0) {
            rest = rawText.slice(prefix.length).trim();
        } else if (rawText.indexOf('青川县庄子上工业园区智慧园区平台') === 0) {
            // 容错：前缀格式略有差异
            var slashIdx = rawText.indexOf(' / ');
            rest = slashIdx >= 0 ? rawText.slice(slashIdx + 3).trim() : rawText;
        } else {
            rest = rawText;
        }

        newBc.innerHTML = '青川县庄子上工业园区智慧园区平台 / <span>' + rest + '</span>';
        hl.appendChild(newBc);
    }

    /* 注入标题栏全套按钮 */
    function initHeader() {
        injectCSS();

        // 先标准化左侧
        normalizeHeaderLeft();

        var hr = document.querySelector('.page-header-right');
        if (!hr) return;

        // 清除所有旧按钮/时间元素
        hr.querySelectorAll(
            '.phdr-btn, .phdr-time, ' +
            '.theme-toggle-btn, .page-theme-toggle-btn, .page-logout-btn, ' +
            '.page-header-btn, .page-header-user, .page-header-time, ' +
            '#pageTime, #clock'
        ).forEach(function(b) { b.remove(); });

        // 注入：主题切换 → 清除缓存 → 用户名 → 退出 → 时间
        hr.appendChild(buildThemeToggle());
        hr.appendChild(buildClearCache());
        hr.appendChild(buildUserTag());
        hr.appendChild(buildLogout());
        hr.appendChild(buildTimeDiv());

        hr.dataset.headerInited = '1';
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initHeader);
    } else {
        initHeader();
    }

    /* 拦截 toggleTheme：切换后同步更新主题按钮图标 */
    var _origToggle = window.toggleTheme;
    window.toggleTheme = function() {
        if (_origToggle) _origToggle.apply(this, arguments);
        requestAnimationFrame(function() {
            document.querySelectorAll('.phdr-btn.phdr-theme').forEach(updateThemeBtn);
        });
    };
})();

/* ─── G2. 清除缓存功能（子页面版，强制覆盖各页面旧版本）─── */
window.clearCacheNow = async function() {
    showToast('\uD83E\uDDF9 \u6b63\u5728\u6e05\u9664\u6d4f\u89c8\u5668\u7f13\u5b58...', 'info', 2000);
    // localStorage（保留主题和登录态）
    var keep = ['ss_theme', 'ss_username', 'ss_token'];
    var toRemove = [];
    for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && keep.indexOf(k) === -1) toRemove.push(k);
    }
    toRemove.forEach(function(k) { localStorage.removeItem(k); });
    // sessionStorage
    sessionStorage.clear();
    // Cache API
    var cleared = 0;
    if (window.caches) {
        try {
            var keys = await caches.keys();
            await Promise.all(keys.map(function(k) { return caches.delete(k); }));
            cleared = keys.length;
        } catch(e) {}
    }
    // Service Worker
    if (navigator.serviceWorker) {
        try {
            var regs = await navigator.serviceWorker.getRegistrations();
            await Promise.all(regs.map(function(r) { return r.unregister(); }));
        } catch(e) {}
    }
    setTimeout(function() {
        showToast('\u2705 \u7f13\u5b58\u5df2\u5168\u90e8\u6e05\u9664\uff01\u5171\u6e05\u9664 ' + (toRemove.length + cleared) + ' \u9879', 'success', 3000);
    }, 600);
};

/* ─── H. Checkbox 全选/反选 ─── */
(function() {
    function bindCb() {
        document.querySelectorAll('#__cb_all').forEach(function(cb) {
            if (cb.dataset.cbInited) return; cb.dataset.cbInited='1';
            cb.addEventListener('change', function() {
                var tbl = cb.closest('table');
                if (!tbl) return;
                tbl.querySelectorAll('.row-cb').forEach(function(rc) {
                    rc.checked = cb.checked;
                    rc.closest('tr').style.background = cb.checked ? (isLight()?'rgba(22,119,255,0.06)':'rgba(64,196,255,0.08)') : '';
                });
            });
        });
        document.querySelectorAll('.row-cb').forEach(function(rc) {
            if (rc.dataset.cbInited) return; rc.dataset.cbInited='1';
            rc.addEventListener('change', function() {
                rc.closest('tr').style.background = rc.checked ? (isLight()?'rgba(22,119,255,0.06)':'rgba(64,196,255,0.08)') : '';
                var tbl = rc.closest('table'); if (!tbl) return;
                var all = tbl.querySelectorAll('.row-cb');
                var cbAll = tbl.querySelector('#__cb_all');
                if (cbAll) { cbAll.checked = Array.from(all).every(function(c){return c.checked;}); cbAll.indeterminate = !cbAll.checked && Array.from(all).some(function(c){return c.checked;}); }
            });
        });
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', function(){setTimeout(bindCb,400);});
    else setTimeout(bindCb, 400);
    window._rebindCheckboxes = bindCb;
})();

/* ─── I. 地图监控（自动绑定+边界）─── */
(function() {
    // mapId → window变量名 映射表
    var MAP_OBJ_MAP = {
        'mapLocation':   '_mapLocationObj',
        'mapTrip':       '_mapTripObj',
        'mapTrajectory': '_mapTrajectoryObj',
        'mapBoundary':   '_mapBoundaryObj',
        'mapFixed':      '_mapFixedObj',
        'mapInTransit':  '_mapInTransitObj'
    };
    var done = {};
    function chk() {
        Object.keys(MAP_OBJ_MAP).forEach(function(mid) {
            if (done[mid]) return;
            var el = document.getElementById(mid);
            if (el && el.classList.contains('leaflet-container')) {
                done[mid] = true;
                var mo = window[MAP_OBJ_MAP[mid]];
                initMapControls(mid, mo);
                if (mo && !mo._pengzhouBoundaryDrawn) {
                    mo._pengzhouBoundaryDrawn = true;
                    setTimeout(function(){ window.drawPengzhouBoundary(mo); }, 300);
                }
            }
        });
    }
    setInterval(chk, 1000); setTimeout(chk, 2000);
})();

/* ─── J. 统一时间显示 ─── */
(function() {
    function tick() {
        var n=new Date(), p=function(s){return String(s).padStart(2,'0');};
        var s=n.getFullYear()+'-'+p(n.getMonth()+1)+'-'+p(n.getDate())+' '+p(n.getHours())+':'+p(n.getMinutes())+':'+p(n.getSeconds());
        document.querySelectorAll('#pageTime,#clock,#headerTime').forEach(function(el){el.textContent=s;});
    }
    tick(); setInterval(tick, 1000);
})();
