/* =============================================
   V19.4 Finance Charts - Chart.js 4.4.0 JS互操作
   供 Blazor 页面通过 IJSRuntime 调用
   V19.5: 修复 Blazor Server 渲染时序——增加 ready() 就绪机制
   ============================================= */

window.financeCharts = {
  /** 等待 Chart.js 就绪 (修复 Blazor Server 预渲染→交互式过渡时序)
   *  轮询 window.Chart, 最大等待 10 秒, 就绪后 resolve */
  ready() {
    return new Promise(function(resolve) {
      if (window.Chart) { resolve(true); return; }
      var attempts = 0;
      var maxAttempts = 100;
      var timer = setInterval(function() {
        attempts++;
        if (window.Chart) { clearInterval(timer); resolve(true); return; }
        if (attempts >= maxAttempts) { clearInterval(timer); resolve(false); }
      }, 100);
    });
  },

  /** 窗口resize时重绘所有Chart实例 (修复平板/移动端图表空白) */
  _resizeAll() {
    if (!window.Chart) return;
    Object.values(Chart.instances || {}).forEach(function(c) {
      try { c.resize(); } catch(e) {}
    });
  },

  /** 安全销毁已有图表，防 "canvas already in use" */
  _destroyChart(canvasId) {
    var existing = Chart.getChart(canvasId);
    if (existing) existing.destroy();
  },

  /** 渲染柱状图+折线图叠加 (月度营收趋势)
   *  canvasId: canvas元素 id
   *  labels: 横轴标签数组
   *  barData: 柱状图数据数组
   *  barLabel: 柱状图图例
   *  lineData: 折线图数据数组
   *  lineLabel: 折线图图例
   */
  async renderBarLine(canvasId, labels, barData, barLabel, lineData, lineLabel) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: barLabel || '营收(万元)',
            data: barData,
            backgroundColor: 'rgba(76,175,80,0.5)',
            borderColor: '#4CAF50',
            borderWidth: 2,
            borderRadius: 6,
            yAxisID: 'y'
          },
          {
            label: lineLabel || '利润(万元)',
            data: lineData,
            type: 'line',
            borderColor: '#2196F3',
            backgroundColor: 'transparent',
            borderWidth: 3,
            pointBackgroundColor: '#2196F3',
            pointRadius: 4,
            tension: 0.3,
            yAxisID: 'y'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e0e0e0', usePointStyle: true } } },
        scales: {
          x: { ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { ticks: { color: '#999', callback: function(v) { return v + 'w'; } }, grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  },

  /** 渲染环形图 (品类销售占比)
   *  canvasId: canvas元素 id
   *  labels: 扇形标签数组
   *  data: 数据数组
   */
  async renderDoughnut(canvasId, labels, data) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: data,
          backgroundColor: ['#4CAF50','#2196F3','#FF9800','#9C27B0','#F44336','#00BCD4','#FFEB3B','#795548'],
          borderColor: '#1e1e32',
          borderWidth: 3
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { color: '#e0e0e0', padding: 15, usePointStyle: true } }
        }
      }
    });
  },

  /** 渲染折线图
   *  canvasId: canvas元素 id
   *  labels: 横轴标签
   *  datasets: [{label, data, color}]
   */
  async renderLine(canvasId, labels, datasets) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    var ds = datasets.map(function(d) {
      return {
        label: d.label || '',
        data: d.data || [],
        borderColor: d.color || '#00bcd4',
        backgroundColor: (d.color || '#00bcd4').replace(')', ',0.1)').replace('rgb', 'rgba'),
        fill: d.fill !== false,
        tension: 0.3
      };
    });
    new Chart(ctx, {
      type: 'line',
      data: { labels: labels, datasets: ds },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e0e0e0' } } },
        scales: {
          x: { ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  },

  /** 渲染饼图 */
  async renderPie(canvasId, labels, data) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    new Chart(ctx, {
      type: 'pie',
      data: {
        labels: labels,
        datasets: [{
          data: data,
          backgroundColor: ['#4CAF50','#2196F3','#FF9800','#9C27B0','#F44336','#00BCD4','#FFEB3B','#795548'],
          borderColor: '#1e1e32',
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { color: '#e0e0e0', padding: 15 } }
        }
      }
    });
  },

  /** 渲染堆叠柱状图
   *  canvasId: canvas元素 id
   *  labels: 横轴标签
   *  datasets: [{label, data, color}]
   */
  async renderStackedBar(canvasId, labels, datasets) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    var ds = datasets.map(function(d) {
      return {
        label: d.label || '',
        data: d.data || [],
        backgroundColor: d.color || '#4CAF50',
        borderRadius: 4
      };
    });
    new Chart(ctx, {
      type: 'bar',
      data: { labels: labels, datasets: ds },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e0e0e0', usePointStyle: true } } },
        scales: {
          x: { stacked: true, ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { stacked: true, ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  },

  /** 渲染横向柱状图
   *  canvasId: canvas元素 id
   *  labels: 纵轴标签
   *  data: 数据数组
   *  color: 柱色
   */
  async renderHorizontalBar(canvasId, labels, data, color) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: '',
          data: data,
          backgroundColor: color || 'rgba(0,188,212,0.6)',
          borderColor: color || '#00bcd4',
          borderWidth: 1,
          borderRadius: 4
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { ticks: { color: '#e0e0e0' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  },

  /** 渲染仪表盘(环形半圆) 用于KPI展示
   *  canvasId: canvas元素 id
   *  value: 当前值 (0-100)
   *  label: 标签文字
   */
  async renderGauge(canvasId, value, label) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    var v = Math.min(100, Math.max(0, value));
    new Chart(ctx, {
      type: 'doughnut',
      data: {
        datasets: [{
          data: [v, 100 - v],
          backgroundColor: [v >= 70 ? '#4CAF50' : v >= 40 ? '#FF9800' : '#F44336', 'rgba(255,255,255,0.08)'],
          borderWidth: 0,
          circumference: 270,
          rotation: 225
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '75%',
        plugins: {
          legend: { display: false },
          tooltip: { enabled: false }
        }
      }
    });
  },

  /** 渲染多序列柱状图(并排)
   *  canvasId: canvas元素 id
   *  labels: 横轴标签
   *  datasets: [{label, data, color}]
   */
  async renderMultiBar(canvasId, labels, datasets) {
    await this.ready();
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    this._destroyChart(canvasId);
    var ds = datasets.map(function(d, i) {
      var colors = ['#4CAF50','#2196F3','#FF9800','#9C27B0','#F44336','#00BCD4'];
      return {
        label: d.label || '',
        data: d.data || [],
        backgroundColor: d.color || colors[i % colors.length],
        borderRadius: 4
      };
    });
    new Chart(ctx, {
      type: 'bar',
      data: { labels: labels, datasets: ds },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e0e0e0', usePointStyle: true } } },
        scales: {
          x: { ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { ticks: { color: '#999' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  }
};

// 窗口resize时自动重绘所有图表，修复平板/移动端图表空白问题
window.addEventListener('resize', function() {
  if (window.financeCharts && window.financeCharts._resizeAll) {
    window.financeCharts._resizeAll();
  }
});
