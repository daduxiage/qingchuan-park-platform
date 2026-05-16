// ===== 平台事件日志追踪（服务端版） =====
window._PageTracker = (function(){
    var API = '/api/log';

    function postLog(entry) {
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('POST', API, true);
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.setRequestHeader('ngrok-skip-browser-warning', '1');
            xhr.timeout = 3000;
            xhr.send(JSON.stringify(entry));
        } catch(e) {}
    }

    function log(eventType, detail) {
        postLog({
            user: sessionStorage.getItem('login_name') || '未知',
            ua: navigator.userAgent.substring(0, 120),
            event: eventType,
            detail: detail || '',
            page: window.location.pathname
        });
    }

    function logLogin(user) { log('登录平台', '用户 ' + user + ' 登录平台'); }
    function logLogout(user) { log('登出平台', '用户 ' + user + ' 登出平台'); }
    function logPageView(pageName) { log('页面访问', '访问页面: ' + pageName); }
    function logModuleClick(moduleName) { log('模块点击', '点击模块: ' + moduleName); }

    function getAllEvents(callback, retryCount) {
        retryCount = retryCount || 0;
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', API, true);
            xhr.setRequestHeader('ngrok-skip-browser-warning', '1');
            xhr.timeout = 5000;
            xhr.onload = function() {
                if (xhr.status === 200) {
                    try { callback(JSON.parse(xhr.responseText)); }
                    catch(e) { callback([]); }
                } else if (retryCount < 2) {
                    setTimeout(function() { getAllEvents(callback, retryCount + 1); }, 1000);
                } else { callback([]); }
            };
            xhr.onerror = function() {
                if (retryCount < 2) setTimeout(function() { getAllEvents(callback, retryCount + 1); }, 1000);
                else callback([]);
            };
            xhr.ontimeout = function() {
                if (retryCount < 2) setTimeout(function() { getAllEvents(callback, retryCount + 1); }, 1000);
                else callback([]);
            };
            xhr.send();
        } catch(e) { callback([]); }
    }

    return {
        log: log,
        logLogin: logLogin,
        logLogout: logLogout,
        logPageView: logPageView,
        logModuleClick: logModuleClick,
        getAllEvents: getAllEvents
    };
})();

// 自动记录页面访问
(function(){
    var path = window.location.pathname;
    var name = path.split('/').pop().replace('.html','') || '首页';
    _PageTracker.logPageView(name);
})();
