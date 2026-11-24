/* 点呼支援レポート — MyGeotab Add-in
 * 前日の運転データを本人の過去7日平均と比較し、定型コメント付きの点呼用一覧を作る。
 * ファイル名(tenkoReport.html)と geotab.addin.tenkoReport の名前は一致させること。
 */
(function () {
  'use strict';

  window.geotab = window.geotab || {};
  geotab.addin = geotab.addin || {};

  // ===== 設定（運用に合わせて変更） =====================================
  var CONFIG = {
    baselineDays: 7,          // 比較対象：前日より前の何日間か
    minDistanceKm: 20,        // 前日の走行距離がこれ未満なら「参考」扱い
    thresholds: {
      upRatio: 1.5,           // 前日の発生率が平均の1.5倍以上 → 増加
      downRatio: 0.5,         // 前日の発生率が平均の0.5倍以下 → 減少
      minCount: 2             // 増減と判定する最低回数（偶発的な1回を除外）
    },
    // ルールIDは MyGeotab の「ルール」設定に合わせて確認・変更してください
    categories: [
      { key: 'brake', label: '急ブレーキ', ruleIds: ['RuleHarshBrakingId'] },
      { key: 'accel', label: '急加速',     ruleIds: ['RuleJackrabbitStartsId'] },
      { key: 'speed', label: '速度超過',   ruleIds: ['RulePostedSpeedingId'] }
    ],
    comments: {
      brake: {
        up:   '急ブレーキが普段より増えています。車間距離と早めの減速を意識しましょう。',
        down: '急ブレーキが普段より減っています。この調子で続けましょう。'
      },
      accel: {
        up:   '急加速が普段より増えています。発進はゆっくり、穏やかなアクセル操作を。',
        down: '急加速が普段より減っています。穏やかな発進ができています。'
      },
      speed: {
        up:   '速度超過が普段より増えています。時間に余裕を持ち、制限速度を守りましょう。',
        down: '速度管理が改善しています。今週も継続しましょう。'
      }
    },
    goodComment:    '前日は急ブレーキ・急加速・速度超過がありませんでした。引き続き安全運転を。',
    steadyComment:  '普段と大きな変化はありません。引き続き安全運転を。',
    noBaseComment:  '比較できる過去データがありません。今日から傾向を確認していきます。',
    lowDataComment: '（前日の走行距離が短いため、比較は参考程度です）'
  };

  var STATUS = {
    warn:   { label: '注意',     order: 0 },
    better: { label: '改善',     order: 1 },
    good:   { label: '良好',     order: 2 },
    flat:   { label: '変化なし', order: 3 },
    nobase: { label: '比較なし', order: 4 },
    low:    { label: '参考',     order: 5 }
  };

  var UNKNOWN_DRIVER = { UnknownDriverId: 1, NoDriverId: 1 };

  // ===== ユーティリティ ===================================================
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function md(d) { return (d.getMonth() + 1) + '/' + d.getDate(); }
  function parseYmd(s) { var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
  function idOf(x) { return !x ? null : (typeof x === 'string' ? x : x.id); }
  function fmt1(n) { return n == null ? '–' : (Math.round(n * 10) / 10).toFixed(1); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function $(id) { return document.getElementById(id); }

  var store = {
    get: function (k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* 保存不可でも動作は継続 */ } }
  };

  // ===== データ取得 =======================================================
  function multiCall(api, calls) {
    return new Promise(function (resolve, reject) { api.multiCall(calls, resolve, reject); });
  }

  function loadRaw(api, state, target) {
    var dayStart = target;
    var dayEnd = addDays(target, 1);
    var baseStart = addDays(target, -CONFIG.baselineDays);
    var groups = state && typeof state.getGroupFilter === 'function' ? state.getGroupFilter() : null;
    var deviceSearch = groups && groups.length ? { groups: groups } : undefined;
    var range = { fromDate: baseStart.toISOString(), toDate: dayEnd.toISOString() };

    var ruleIds = [];
    CONFIG.categories.forEach(function (c) {
      c.ruleIds.forEach(function (id) { if (ruleIds.indexOf(id) < 0) ruleIds.push(id); });
    });

    var calls = [
      ['Get', { typeName: 'Trip', search: { fromDate: range.fromDate, toDate: range.toDate, deviceSearch: deviceSearch } }],
      ['Get', { typeName: 'Device', search: deviceSearch || {} }],
      ['Get', { typeName: 'User', search: { isDriver: true } }]
    ];
    ruleIds.forEach(function (id) {
      calls.push(['Get', { typeName: 'ExceptionEvent', search: {
        fromDate: range.fromDate, toDate: range.toDate, ruleSearch: { id: id }, deviceSearch: deviceSearch
      } }]);
    });

    return multiCall(api, calls).then(function (res) {
      var byRule = {};
      ruleIds.forEach(function (id, i) { byRule[id] = res[3 + i] || []; });
      return {
        trips: res[0] || [], devices: res[1] || [], users: res[2] || [],
        exceptionsByRule: byRule, dayStart: dayStart, dayEnd: dayEnd, baseStart: baseStart
      };
    });
  }

  // ===== 集計・判定 =======================================================
  function aggregate(raw, unit) {
    var devName = {}, userName = {};
    raw.devices.forEach(function (d) { devName[d.id] = d.name; });
    raw.users.forEach(function (u) {
      userName[u.id] = ((u.lastName || '') + ' ' + (u.firstName || '')).trim() || u.name;
    });

    var ents = {};
    var unknownKm = 0;
    var t0 = raw.baseStart.getTime(), t1 = raw.dayStart.getTime(), t2 = raw.dayEnd.getTime();

    function keyOf(devId, drvId) {
      if (unit === 'device') return devId;
      return drvId && !UNKNOWN_DRIVER[drvId] ? drvId : null;
    }
    function ent(key) {
      if (!ents[key]) {
        var c = {};
        CONFIG.categories.forEach(function (cat) { c[cat.key] = { day: 0, base: 0 }; });
        ents[key] = {
          id: key,
          name: unit === 'driver' ? (userName[key] || key) : (devName[key] || key),
          dayKm: 0, baseKm: 0, c: c
        };
      }
      return ents[key];
    }
    function slot(time) {
      var t = new Date(time).getTime();
      if (isNaN(t) || t < t0 || t >= t2) return null;
      return t >= t1 ? 'day' : 'base';
    }

    raw.trips.forEach(function (tr) {
      var s = slot(tr.start);
      if (!s) return;
      var key = keyOf(idOf(tr.device), idOf(tr.driver));
      var km = tr.distance || 0;
      if (!key) { if (s === 'day') unknownKm += km; return; }
      ent(key)[s + 'Km'] += km;
    });

    CONFIG.categories.forEach(function (cat) {
      cat.ruleIds.forEach(function (rid) {
        (raw.exceptionsByRule[rid] || []).forEach(function (ev) {
          var s = slot(ev.activeFrom);
          if (!s) return;
          var key = keyOf(idOf(ev.device), idOf(ev.driver));
          if (!key) return;
          ent(key).c[cat.key][s]++;
        });
      });
    });

    var rows = Object.keys(ents).map(function (k) { return ents[k]; })
      .filter(function (e) { return e.dayKm > 0; })   // 前日に運行した人（車両）が点呼対象
      .map(evaluate)
      .sort(function (a, b) {
        return STATUS[a.status].order - STATUS[b.status].order || String(a.name).localeCompare(String(b.name), 'ja');
      });

    return { rows: rows, unknownKm: unknownKm };
  }

  function evaluate(e) {
    var T = CONFIG.thresholds;
    e.cats = CONFIG.categories.map(function (cat) {
      var c = e.c[cat.key];
      var dayRate = e.dayKm > 0 ? c.day / e.dayKm * 100 : 0;
      var baseRate = e.baseKm > 0 ? c.base / e.baseKm * 100 : null;
      var trend = 'flat';
      if (baseRate === null) trend = 'nobase';
      else if (c.day >= T.minCount && (baseRate === 0 || dayRate >= baseRate * T.upRatio)) trend = 'up';
      // 減少：普段なら前日の距離で1回以上は出ていたはず（期待値>=1）なのに、平均の半分以下だった
      else if (c.base >= T.minCount && baseRate * e.dayKm / 100 >= 1 && dayRate <= baseRate * T.downRatio) trend = 'down';
      return { key: cat.key, label: cat.label, day: c.day, base: c.base, dayRate: dayRate, baseRate: baseRate, trend: trend };
    });

    var ups = e.cats.filter(function (c) { return c.trend === 'up'; });
    var downs = e.cats.filter(function (c) { return c.trend === 'down'; });
    var allZero = e.cats.every(function (c) { return c.day === 0; });
    var noBase = e.baseKm <= 0;
    e.lowData = e.dayKm < CONFIG.minDistanceKm;

    if (e.lowData) e.status = 'low';
    else if (ups.length) e.status = 'warn';
    else if (downs.length) e.status = 'better';
    else if (noBase) e.status = 'nobase';
    else if (allZero) e.status = 'good';
    else e.status = 'flat';

    e.changes = ups.concat(downs).map(function (c) {
      return { trend: c.trend, text: c.label + '：' + fmt1(c.baseRate) + ' → ' + fmt1(c.dayRate) + '回/100km' };
    });

    var lines = ups.map(function (c) { return CONFIG.comments[c.key].up; })
      .concat(downs.map(function (c) { return CONFIG.comments[c.key].down; }));
    if (!lines.length) lines.push(noBase ? CONFIG.noBaseComment : allZero ? CONFIG.goodComment : CONFIG.steadyComment);
    if (e.lowData) lines.push(CONFIG.lowDataComment);
    e.templateComment = lines.join('\n');
    return e;
  }

  // ===== 画面 =============================================================
  function createApp() {
    var api, state, current = null, loading = false;

    function storeKey(row) { return 'tenkoReport:' + $('tk-date').value + ':' + $('tk-unit').value + ':' + row.id; }
    function noteOf(row) {
      var saved = store.get(storeKey(row)) || {};
      return {
        comment: typeof saved.comment === 'string' ? saved.comment : row.templateComment,
        edited: typeof saved.comment === 'string' && saved.comment !== row.templateComment,
        checked: !!saved.checked
      };
    }
    function saveNote(row, patch) {
      var k = storeKey(row);
      var cur = store.get(k) || {};
      Object.keys(patch).forEach(function (p) { cur[p] = patch[p]; });
      store.set(k, cur);
    }

    function showMsg(text, isErr) {
      var el = $('tk-msg');
      el.hidden = !text;
      el.textContent = text || '';
      el.className = 'tk-msg' + (isErr ? ' err' : '');
    }

    function cellHtml(c) {
      var arrow = c.trend === 'up' ? '<span class="tk-arrow">▲</span>' : c.trend === 'down' ? '<span class="tk-arrow">▼</span>' : '';
      return '<td class="tk-cell ' + c.trend + '">' +
        '<div class="cnt">' + c.day + '<small>回</small>' + arrow + '</div>' +
        '<div class="rate">' + fmt1(c.dayRate) + ' /100km</div>' +
        '<div class="rate">平均 ' + fmt1(c.baseRate) + '</div></td>';
    }

    function render() {
      if (!current) return;
      var filter = $('tk-filter').value;
      var rows = current.rows;
      var counts = { warn: 0, better: 0, checked: 0 };
      rows.forEach(function (r) {
        if (r.status === 'warn') counts.warn++;
        if (r.status === 'better') counts.better++;
        if (noteOf(r).checked) counts.checked++;
      });
      $('k-total').textContent = rows.length;
      $('k-warn').textContent = counts.warn;
      $('k-better').textContent = counts.better;
      $('k-checked').textContent = counts.checked + ' / ' + rows.length;

      var shown = rows.filter(function (r) {
        if (filter === 'changed') return r.status === 'warn' || r.status === 'better';
        if (filter === 'unchecked') return !noteOf(r).checked;
        return true;
      });

      var html = shown.map(function (r) {
        var n = noteOf(r);
        var idx = rows.indexOf(r);
        var changes = r.changes.length
          ? '<ul class="tk-changes">' + r.changes.map(function (c) { return '<li class="' + c.trend + '">' + esc(c.text) + '</li>'; }).join('') + '</ul>'
          : '<span class="tk-none">—</span>';
        return '<tr class="st-' + r.status + (n.checked ? ' is-checked' : '') + '" data-idx="' + idx + '">' +
          '<td><div class="tk-name">' + esc(r.name) + '</div>' +
          '<div class="tk-sub">前日 ' + Math.round(r.dayKm) + ' km</div>' +
          '<div class="tk-sub">' + CONFIG.baselineDays + '日計 ' + Math.round(r.baseKm) + ' km</div></td>' +
          r.cats.map(cellHtml).join('') +
          '<td><span class="tk-badge ' + r.status + '">' + STATUS[r.status].label + '</span></td>' +
          '<td>' + changes + '</td>' +
          '<td><textarea class="tk-comment' + (n.edited ? ' edited' : '') + '" aria-label="' + esc(r.name) + 'の点呼コメント">' + esc(n.comment) + '</textarea>' +
          '<div class="print-only">' + esc(n.comment) + '</div>' +
          '<button type="button" class="tk-reset"' + (n.edited ? '' : ' hidden') + '>定型コメントに戻す</button></td>' +
          '<td><input type="checkbox" class="tk-check" aria-label="' + esc(r.name) + 'を確認済にする"' + (n.checked ? ' checked' : '') + '></td>' +
          '</tr>';
      }).join('');

      $('tk-body').innerHTML = html || '<tr><td colspan="8" class="tk-none">表示する対象がありません。</td></tr>';
      Array.prototype.forEach.call(document.querySelectorAll('#tenko .tk-comment'), autoGrow);
    }

    function autoGrow(ta) {
      ta.style.height = 'auto';
      ta.style.height = (ta.scrollHeight + 2) + 'px';
    }

    function periodText(raw, unknownKm) {
      var t = 'ルール判定の対象期間　前日：' + md(raw.dayStart) +
        '　／　比較：' + md(raw.baseStart) + '〜' + md(addDays(raw.dayStart, -1)) + '（' + CONFIG.baselineDays + '日間）';
      if ($('tk-unit').value === 'driver' && unknownKm > 0) {
        t += '　／　ドライバー未識別の前日走行 ' + Math.round(unknownKm) + ' km は集計外';
      }
      return t;
    }

    function load() {
      if (loading || !api) return;
      loading = true;
      $('tk-reload').disabled = true;
      showMsg('データを取得しています…');
      var target = parseYmd($('tk-date').value);
      loadRaw(api, state, target).then(function (raw) {
        current = aggregate(raw, $('tk-unit').value);
        $('tk-period').textContent = periodText(raw, current.unknownKm);
        showMsg(current.rows.length ? '' : '前日に運行した対象がありません。日付・グループ・集計単位を確認してください。');
        render();
      }).catch(function (err) {
        current = null;
        $('tk-body').innerHTML = '';
        showMsg('データ取得に失敗しました：' + (err && err.message ? err.message : err), true);
      }).then(function () {
        loading = false;
        $('tk-reload').disabled = false;
      });
    }

    function rowFromEvent(ev) {
      var tr = ev.target.closest('tr[data-idx]');
      return tr && current ? { tr: tr, row: current.rows[+tr.dataset.idx] } : null;
    }

    function exportCsv() {
      if (!current) return;
      var head = ['対象', '前日走行km', CONFIG.baselineDays + '日計km'];
      CONFIG.categories.forEach(function (c) {
        head.push(c.label + '(前日回数)', c.label + '(前日/100km)', c.label + '(平均/100km)');
      });
      head.push('判定', '確認した変化', '点呼コメント', '確認済');
      var lines = [head];
      current.rows.forEach(function (r) {
        var n = noteOf(r);
        var line = [r.name, Math.round(r.dayKm), Math.round(r.baseKm)];
        r.cats.forEach(function (c) { line.push(c.day, fmt1(c.dayRate), fmt1(c.baseRate)); });
        line.push(STATUS[r.status].label, r.changes.map(function (c) { return c.text; }).join(' / '), n.comment, n.checked ? '済' : '');
        lines.push(line);
      });
      var csv = lines.map(function (l) {
        return l.map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(',');
      }).join('\r\n');
      var blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = '点呼一覧_' + $('tk-date').value + '.csv';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 0);
    }

    function bindUi() {
      CONFIG.categories.forEach(function (c, i) { var th = $('th-cat' + i); if (th) th.textContent = c.label; });
      var T = CONFIG.thresholds;
      $('tk-rule').textContent = '※判定条件：走行100kmあたりの発生回数を本人の過去' + CONFIG.baselineDays + '日間と比較。' +
        '「注意」＝前日' + T.minCount + '回以上かつ平均の' + T.upRatio + '倍以上、' +
        '「改善」＝平均の' + T.downRatio + '倍以下（普段なら前日の距離で1回以上出ている人のみ）、' +
        '前日' + CONFIG.minDistanceKm + 'km未満は「参考」。';
      $('tk-date').value = ymd(addDays(new Date(), -1));
      $('tk-date').max = ymd(new Date());
      $('tk-reload').addEventListener('click', load);
      $('tk-date').addEventListener('change', load);
      $('tk-unit').addEventListener('change', load);
      $('tk-filter').addEventListener('change', render);
      $('tk-print').addEventListener('click', function () { window.print(); });
      $('tk-csv').addEventListener('click', exportCsv);

      var body = $('tk-body');
      body.addEventListener('input', function (ev) {
        if (!ev.target.classList.contains('tk-comment')) return;
        var x = rowFromEvent(ev); if (!x) return;
        var val = ev.target.value;
        autoGrow(ev.target);
        saveNote(x.row, { comment: val });
        var edited = val !== x.row.templateComment;
        ev.target.classList.toggle('edited', edited);
        x.tr.querySelector('.tk-reset').hidden = !edited;
        x.tr.querySelector('.print-only').textContent = val;
      });
      body.addEventListener('click', function (ev) {
        if (!ev.target.classList.contains('tk-reset')) return;
        var x = rowFromEvent(ev); if (!x) return;
        saveNote(x.row, { comment: null });
        render();
      });
      body.addEventListener('change', function (ev) {
        if (!ev.target.classList.contains('tk-check')) return;
        var x = rowFromEvent(ev); if (!x) return;
        saveNote(x.row, { checked: ev.target.checked });
        render();
      });
    }

    return {
      initialize: function (freshApi, freshState, callback) {
        api = freshApi; state = freshState;
        bindUi();
        callback();
      },
      focus: function (freshApi, freshState) {
        api = freshApi; state = freshState;
        load();
      },
      blur: function () { }
    };
  }

  geotab.addin.tenkoReport = function () { return createApp(); };

  // ===== デモモード（ローカルで開いた時 / ?demo=1）=========================
  function mulberry32(a) {
    return function () {
      a |= 0; a = a + 0x6D2B79F5 | 0;
      var t = Math.imul(a ^ a >>> 15, 1 | a);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }
  function hash(s) { var h = 2166136261; for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return h; }
  function poisson(rng, lambda) { var L = Math.exp(-lambda), k = 0, p = 1; do { k++; p *= rng(); } while (p > L); return k - 1; }

  function makeDemoApi() {
    // base: 平常時の100kmあたり発生率, day: 前日だけ変える値, km: 1日の走行距離
    var people = [
      { last: '山田', first: '太郎', plate: '品川 100 あ 12-01', base: { brake: 0.4, accel: 0.3, speed: 0.5 }, day: { brake: 3.0 } },
      { last: '佐藤', first: '次郎', plate: '品川 100 あ 12-02', base: { brake: 0.5, accel: 0.4, speed: 2.2 }, day: { speed: 0 } },
      { last: '鈴木', first: '花子', plate: '品川 100 あ 12-03', base: { brake: 0.1, accel: 0.1, speed: 0.1 }, day: { brake: 0, accel: 0, speed: 0 } },
      { last: '高橋', first: '健',   plate: '品川 100 あ 12-04', base: { brake: 0.6, accel: 0.5, speed: 0.6 }, day: { accel: 2.6, speed: 2.0 } },
      { last: '田中', first: '誠',   plate: '品川 100 あ 12-05', base: { brake: 0.7, accel: 0.6, speed: 0.8 }, day: { brake: 0.8, accel: 0.6, speed: 0.8 } },
      { last: '伊藤', first: '美咲', plate: '品川 100 あ 12-06', base: { brake: 0.5, accel: 0.5, speed: 0.5 }, day: {}, dayKm: 12 },
      { last: '渡辺', first: '翔',   plate: '品川 100 あ 12-07', base: { brake: 0.9, accel: 0.4, speed: 0.6 }, day: { brake: 1.0, accel: 0.5, speed: 0.6 } },
      { last: '中村', first: '大輔', plate: '品川 100 あ 12-08', base: { brake: 1.8, accel: 0.6, speed: 0.4 }, day: { brake: 0, accel: 0.6 } }
    ];
    var ruleOf = { brake: 'RuleHarshBrakingId', accel: 'RuleJackrabbitStartsId', speed: 'RulePostedSpeedingId' };

    function build(from, to) {
      var target = addDays(new Date(to.getTime() - 1), 0);
      target.setHours(0, 0, 0, 0);
      var out = { trips: [], ex: { RuleHarshBrakingId: [], RuleJackrabbitStartsId: [], RulePostedSpeedingId: [] } };
      for (var d = new Date(from); d < to; d = addDays(d, 1)) {
        var isTarget = d.getTime() === target.getTime();
        people.forEach(function (p, i) {
          var rng = mulberry32(hash(ymd(d) + ':' + i));
          var km = isTarget && p.dayKm ? p.dayKm : Math.round(110 + rng() * 120);
          var drv = { id: 'demo-u' + i }, dev = { id: 'demo-d' + i };
          var t1 = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 7, 30);
          var t2 = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 13, 10);
          out.trips.push({ device: dev, driver: drv, start: t1.toISOString(), distance: km * 0.55 });
          out.trips.push({ device: dev, driver: drv, start: t2.toISOString(), distance: km * 0.45 });
          Object.keys(ruleOf).forEach(function (k) {
            var rate = isTarget && p.day[k] != null ? p.day[k] : p.base[k];
            var n = isTarget && p.day[k] != null ? Math.round(rate * km / 100) : poisson(rng, rate * km / 100);
            for (var j = 0; j < n; j++) {
              var at = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 8 + Math.floor(rng() * 9), Math.floor(rng() * 60));
              out.ex[ruleOf[k]].push({ rule: { id: ruleOf[k] }, device: dev, driver: drv, activeFrom: at.toISOString() });
            }
          });
        });
      }
      return out;
    }

    return {
      multiCall: function (calls, ok) {
        var s0 = calls[0][1].search;
        var data = build(new Date(s0.fromDate), new Date(s0.toDate));
        var res = calls.map(function (c) {
          var p = c[1];
          if (p.typeName === 'Trip') return data.trips;
          if (p.typeName === 'Device') return people.map(function (x, i) { return { id: 'demo-d' + i, name: x.plate }; });
          if (p.typeName === 'User') return people.map(function (x, i) { return { id: 'demo-u' + i, lastName: x.last, firstName: x.first, name: 'demo' + i }; });
          if (p.typeName === 'ExceptionEvent') return data.ex[p.search.ruleSearch.id] || [];
          return [];
        });
        setTimeout(function () { ok(res); }, 250);
      }
    };
  }

  var isDemo = location.protocol === 'file:' || /[?&]demo=1\b/.test(location.search);
  if (isDemo) {
    document.addEventListener('DOMContentLoaded', function () {
      $('tk-demo').hidden = false;
      var app = geotab.addin.tenkoReport();
      var demoApi = makeDemoApi();
      var demoState = { getGroupFilter: function () { return []; } };
      app.initialize(demoApi, demoState, function () { app.focus(demoApi, demoState); });
    });
  }
})();
