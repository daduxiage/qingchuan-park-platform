/**
 * park_loading.js — 园区页面统一加载动画控制
 * 自动隐藏加载遮罩（配合 CSS 动画 bar 1.5s + 0.5s 缓冲 = 2s）
 */
(function(){
  var timer = null;
  window.hideLoading = function(){
    if (timer) clearTimeout(timer);
    timer = setTimeout(function(){
      var lo = document.getElementById('loadingOverlay');
      if (lo) lo.classList.add('hide');
    }, 100);
  };
  // 兜底：2s 后自动隐藏
  setTimeout(function(){
    var lo = document.getElementById('loadingOverlay');
    if (lo && !lo.classList.contains('hide')) lo.classList.add('hide');
  }, 2200);
})();
