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
    // hint：ルールのプルダウンに候補として出す、ルール名のキーワード
    categories: [
      { key: 'brake', label: '急ブレーキ', ruleIds: ['RuleHarshBrakingId'],     hint: /ブレーキ|制動|減速|brak|decel/i },
      { key: 'accel', label: '急加速',     ruleIds: ['RuleJackrabbitStartsId'], hint: /加速|発進|accel|jackrabbit/i },
      { key: 'speed', label: '速度超過',   ruleIds: ['RulePostedSpeedingId'],   hint: /速度|スピード|speed/i },
      // モービルアイ（group: 'me'）。ruleIds を空にしておくと、名前が「ME…」で hint に合うルールを自動で全部使う
      // （例：ME車線逸脱警報-右 と ME車線逸脱警報-左 を合算）
      { key: 'fcw', label: 'ME追突警報',   short: '追突',     group: 'me', ruleIds: [], hint: /追突|前方衝突|forward\s*collision|fcw/i },
      { key: 'hmw', label: 'ME車間警報',   short: '車間',     group: 'me', ruleIds: [], hint: /車間|headway|hmw/i },
      { key: 'ldw', label: 'ME車線逸脱',   short: '車線逸脱', group: 'me', ruleIds: [], hint: /車線|lane|ldw/i },
      { key: 'pcw', label: 'ME歩行者警報', short: '歩行者',   group: 'me', ruleIds: [], hint: /歩行者|pedestrian|pcw/i }
    ],
    // モービルアイのルールの見分け方（ルール名が「ME」で始まる、または Mobileye を含む）
    meRule: /^\s*ME(?![a-z])|[Mm]obileye|モービルアイ/,
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
      },
      fcw: {
        up:   '追突警報が普段より増えています。前の車との距離を保ち、早めのブレーキを心がけましょう。',
        down: '追突警報が普段より減っています。この調子で続けましょう。'
      },
      hmw: {
        up:   '車間警報が普段より増えています。前の車との車間を十分にとりましょう。',
        down: '車間の取り方が改善しています。今週も継続しましょう。'
      },
      ldw: {
        up:   '車線逸脱警報が普段より増えています。疲れや脇見がないか確認し、車線の中央を保ちましょう。',
        down: '車線逸脱警報が普段より減っています。この調子で続けましょう。'
      },
      pcw: {
        up:   '歩行者警報が普段より増えています。交差点や生活道路では歩行者・自転車に特に注意しましょう。',
        down: '歩行者警報が普段より減っています。引き続き周囲に注意しましょう。'
      }
    },
    goodComment:    '前日は注意が必要な運転の記録がありませんでした。引き続き安全運転を。',
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

  // 画面のグループ選択を優先。未選択なら MyGeotab 上部のグループフィルターに従う
  function loadRaw(api, state, target, groupId, catsIn) {
    var dayStart = target;
    var dayEnd = addDays(target, 1);
    var baseStart = addDays(target, -CONFIG.baselineDays);
    var groups = groupId ? [{ id: groupId }]
      : (state && typeof state.getGroupFilter === 'function' ? state.getGroupFilter() : null);
    var groupFiltered = !!(groups && groups.length);
    // 注意：Trip / ExceptionEvent の検索は deviceSearch のグループ指定が効かない（車両IDのみ有効）。
    // そのためグループは Device の検索だけに使い、トリップとイベントは集計時に車両で絞り込む
    var deviceSearch = groupFiltered ? { groups: groups } : undefined;
    var range = { fromDate: baseStart.toISOString(), toDate: dayEnd.toISOString() };

    var cats = catsIn || activeCategories(groupId);
    var ruleIds = [];
    cats.forEach(function (c) {
      c.ruleIds.forEach(function (id) { if (ruleIds.indexOf(id) < 0) ruleIds.push(id); });
    });

    var calls = [
      ['Get', { typeName: 'Trip', search: { fromDate: range.fromDate, toDate: range.toDate } }],
      ['Get', { typeName: 'Device', search: deviceSearch || {} }],
      ['Get', { typeName: 'User', search: { isDriver: true } }],
      ['Get', { typeName: 'Rule' }]
    ];
    ruleIds.forEach(function (id) {
      calls.push(['Get', { typeName: 'ExceptionEvent', search: {
        fromDate: range.fromDate, toDate: range.toDate, ruleSearch: { id: id }
      } }]);
    });

    return multiCall(api, calls).then(function (res) {
      var byRule = {};
      ruleIds.forEach(function (id, i) { byRule[id] = res[4 + i] || []; });
      return {
        trips: res[0] || [], devices: res[1] || [], users: res[2] || [], rules: res[3] || [],
        categories: cats, exceptionsByRule: byRule, dayStart: dayStart, dayEnd: dayEnd, baseStart: baseStart,
        groupFiltered: groupFiltered
      };
    });
  }

  // 画面で選んだルール（このブラウザに保存）があれば CONFIG より優先する
  // 選択はグループごとに保存（例：三井物産のグループでだけ専用ルールを使う）
  var RULE_STORE_KEY = 'tenkoReport:rules';
  function ruleStoreKey(groupId) { return RULE_STORE_KEY + ':' + (groupId || 'all'); }
  function activeCategories(groupId) {
    var ov = store.get(ruleStoreKey(groupId)) || {};
    return CONFIG.categories.map(function (c) {
      return { key: c.key, label: c.label, group: c.group, ruleIds: ov[c.key] ? [ov[c.key]] : c.ruleIds.slice(), overridden: !!ov[c.key] };
    });
  }

  // ドライバーの表示名。
  // 漢字・かなの名前は MyGeotab 上で「名」欄に姓、「姓」欄に名が入っていることが多いため、
  // MyGeotab の画面と同じ「firstName lastName」の順（例：沢津橋 貴之）で表示する。
  // ローマ字の名前は「lastName firstName」（例：Maekawa Tomoya）のまま。
  var CJK = /[぀-ヿ㐀-鿿豈-﫿]/;
  function displayName(u) {
    var first = (u.firstName || '').trim(), last = (u.lastName || '').trim();
    var parts = CJK.test(first + last) ? [first, last] : [last, first];
    return parts.filter(Boolean).join(' ') || u.name || u.id;
  }

  // ===== 集計・判定 =======================================================
  function aggregate(raw, unit) {
    var devName = {}, userName = {};
    raw.devices.forEach(function (d) { devName[d.id] = d.name; });
    raw.users.forEach(function (u) {
      userName[u.id] = displayName(u);
    });

    var ents = {};
    var unknownKm = 0;
    var t0 = raw.baseStart.getTime(), t1 = raw.dayStart.getTime(), t2 = raw.dayEnd.getTime();

    function keyOf(devId, drvId) {
      if (unit === 'device') return devId;
      return drvId && !UNKNOWN_DRIVER[drvId] ? drvId : null;
    }
    var nDays = CONFIG.baselineDays + 1;   // 比較の7日間＋前日（グラフ用の日別データ）
    var zeros = function () { var a = []; for (var i = 0; i < nDays; i++) a.push(0); return a; };
    function dayIndex(time) {
      var i = Math.floor((new Date(time).getTime() - t0) / 86400000);
      return Math.max(0, Math.min(nDays - 1, i));
    }
    function ent(key) {
      if (!ents[key]) {
        var c = {};
        CONFIG.categories.forEach(function (cat) { c[cat.key] = { day: 0, base: 0, byDay: zeros() }; });
        ents[key] = {
          id: key,
          name: unit === 'driver' ? (userName[key] || key) : (devName[key] || key),
          dayKm: 0, baseKm: 0, c: c, kmByDay: zeros(),
          devs: {}, drvs: {}   // 前日にこの対象と組んだ車両／ドライバー（絞り込み用）
        };
        ents[key][unit === 'driver' ? 'drvs' : 'devs'][key] = 1;
      }
      return ents[key];
    }
    function slot(time) {
      var t = new Date(time).getTime();
      if (isNaN(t) || t < t0 || t >= t2) return null;
      return t >= t1 ? 'day' : 'base';
    }

    var tripsByDev = {};
    var driversSeen = {};
    // グループ指定時は、そのグループの車両（Device の検索結果）のデータだけを集計する
    var allowedDev = null;
    if (raw.groupFiltered) {
      allowedDev = {};
      raw.devices.forEach(function (d) { allowedDev[d.id] = 1; });
    }
    var inScope = function (devId) { return !allowedDev || !!allowedDev[devId]; };

    raw.trips.forEach(function (tr) {
      var devId = idOf(tr.device), drvId = idOf(tr.driver);
      if (!inScope(devId)) return;
      var known = drvId && !UNKNOWN_DRIVER[drvId];
      if (known) {
        (tripsByDev[devId] = tripsByDev[devId] || []).push({
          s: new Date(tr.start).getTime(), e: new Date(tr.stop || tr.start).getTime(), drv: drvId
        });
      }
      var s = slot(tr.start);
      if (!s) return;
      if (known) driversSeen[drvId] = 1;
      var key = keyOf(devId, drvId);
      var km = tr.distance || 0;
      if (!key) { if (s === 'day') unknownKm += km; return; }
      var e = ent(key);
      e[s + 'Km'] += km;
      e.kmByDay[dayIndex(tr.start)] += km;
      if (s === 'day') { e.devs[devId] = 1; if (known) e.drvs[drvId] = 1; }
    });

    // 例外イベントにドライバーが付いていない時は、同じ車両・同じ時刻のトリップのドライバーで補う
    function driverFromTrip(devId, time) {
      var t = new Date(time).getTime();
      var list = tripsByDev[devId] || [];
      for (var i = 0; i < list.length; i++) { if (t >= list[i].s && t <= list[i].e) return list[i].drv; }
      return null;
    }

    var diag = raw.categories.map(function (cat) {
      var total = 0, assigned = 0, known = 0;
      cat.ruleIds.forEach(function (rid) {
        (raw.exceptionsByRule[rid] || []).forEach(function (ev) {
          var s = slot(ev.activeFrom);
          if (!s) return;
          var devId = idOf(ev.device), drvId = idOf(ev.driver);
          if (!inScope(devId)) return;
          total++;
          if (!drvId || UNKNOWN_DRIVER[drvId]) drvId = driverFromTrip(devId, ev.activeFrom);
          if (drvId) known++;   // 集計単位に関係なく、ドライバーが分かった件数
          var key = keyOf(devId, drvId);
          if (!key) return;
          assigned++;
          var cc = ent(key).c[cat.key];
          cc[s]++;
          cc.byDay[dayIndex(ev.activeFrom)]++;
        });
      });
      return { key: cat.key, label: cat.label, group: cat.group, auto: cat.auto, ruleIds: cat.ruleIds, overridden: cat.overridden,
        total: total, assigned: assigned, known: known };
    });

    var rows = Object.keys(ents).map(function (k) { return ents[k]; })
      .filter(function (e) { return e.dayKm > 0; })   // 前日に運行した人（車両）が点呼対象
      .map(evaluate)
      .sort(function (a, b) {
        return STATUS[a.status].order - STATUS[b.status].order || String(a.name).localeCompare(String(b.name), 'ja');
      });

    var byName = function (a, b) { return String(a.name).localeCompare(String(b.name), 'ja'); };
    return {
      rows: rows, unknownKm: unknownKm, diag: diag, rules: raw.rules, baseStart: raw.baseStart, dayStart: raw.dayStart,
      vehicleOpts: raw.devices.map(function (d) { return { id: d.id, name: d.name || d.id }; }).sort(byName),
      driverOpts: Object.keys(driversSeen).map(function (id) { return { id: id, name: userName[id] || id }; }).sort(byName),
      users: raw.users.map(function (u) { return { id: u.id, name: userName[u.id] || u.id, activeTo: u.activeTo, companyGroups: u.companyGroups || [] }; })
    };
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
      return { key: cat.key, label: cat.label, short: cat.short || cat.label, group: cat.group,
        day: c.day, base: c.base, dayRate: dayRate, baseRate: baseRate, trend: trend, byDay: c.byDay };
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

    // ルールが割り当たっていない項目（モービルアイ未導入など）
    function unsetKeys() {
      var u = {};
      (current && current.diag || []).forEach(function (d) { if (!d.ruleIds.length) u[d.key] = 1; });
      return u;
    }

    // モービルアイの4項目は1つの列にまとめて表示（列が増えすぎないように）
    function meCellHtml(r) {
      var unset = unsetKeys();
      var me = r.cats.filter(function (c) { return c.group === 'me'; });
      if (!me.length) return '<td></td>';
      return '<td class="tk-me">' + me.map(function (c) {
        if (unset[c.key]) return '<div class="tk-me-line na">' + esc(c.short) + '<span>—</span></div>';
        var arrow = c.trend === 'up' ? ' ▲' : c.trend === 'down' ? ' ▼' : '';
        return '<div class="tk-me-line ' + c.trend + '" title="' + esc(c.label + '：' + fmt1(c.dayRate) + '/100km（平均 ' + fmt1(c.baseRate) + '）') + '">' +
          esc(c.short) + '<span><b>' + c.day + '</b>回' + arrow + '</span></div>';
      }).join('') + '</td>';
    }

    function cellHtml(c) {
      var arrow = c.trend === 'up' ? '<span class="tk-arrow">▲</span>' : c.trend === 'down' ? '<span class="tk-arrow">▼</span>' : '';
      return '<td class="tk-cell ' + c.trend + '">' +
        '<div class="cnt">' + c.day + '<small>回</small>' + arrow + '</div>' +
        '<div class="rate">' + fmt1(c.dayRate) + ' /100km</div>' +
        '<div class="rate">平均 ' + fmt1(c.baseRate) + '</div></td>';
    }

    // 車両・ドライバーのプルダウンで絞った行（KPI・一覧・CSVの対象）
    function filteredRows() {
      var fDev = $('tk-fdev').value, fDrv = $('tk-fdrv').value;
      return current.rows.filter(function (r) {
        return (!fDev || r.devs[fDev]) && (!fDrv || r.drvs[fDrv]);
      });
    }

    function fillSelect(sel, allLabel, opts) {
      var keep = sel.value;
      sel.innerHTML = '<option value="">' + esc(allLabel) + '</option>' + opts.map(function (o) {
        return '<option value="' + esc(o.id) + '">' + esc(o.label || o.name) + '</option>';
      }).join('');
      sel.value = opts.some(function (o) { return o.id === keep; }) ? keep : '';
    }

    function populateFilters() {
      fillSelect($('tk-fdev'), 'すべての車両', current.vehicleOpts);
      fillSelect($('tk-fdrv'), 'すべてのドライバー', driverOptionsForGroup());
    }

    // 選んだグループ（配下の階層を含む）に所属するドライバーだけを候補にする。
    // グループ未選択、または所属するドライバーが1人もいない時は、期間内に運行したドライバーを候補にする
    function driverOptionsForGroup() {
      var gid = $('tk-group').value;
      if (!gid || !groupById[gid]) return current.driverOpts;
      var sub = {};
      (function walk(id) {
        if (sub[id] || !groupById[id]) return;
        sub[id] = 1;
        (groupById[id].children || []).forEach(function (c) { walk(c.id); });
      })(gid);
      var now = Date.now();
      var inGroup = current.users.filter(function (u) {
        if (u.activeTo && new Date(u.activeTo).getTime() < now) return false;   // 無効化済みのユーザーは除く
        return u.companyGroups.some(function (g) { return sub[idOf(g)]; });
      }).sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'ja'); });
      return inGroup.length ? inGroup : current.driverOpts;
    }

    var GROUP_STORE_KEY = 'tenkoReport:group';
    var groupsLoaded = false;
    var groupById = {};

    // MyGeotab が最初から持つシステムのグループ（Asset Type、燃料種別など）は
    // ID が「Group…Id」の形。御社で作ったグループはランダムなID（例: b27A5）なので、これで見分ける
    var SYSTEM_GROUP = /^Group[A-Za-z]+Id$/;

    // 御社で作ったグループだけを、会社グループ配下のツリー順に並べ、階層を字下げで表す
    function groupOptions(groups) {
      var byId = {};
      groups.forEach(function (g) { byId[g.id] = g; });
      var byName = function (a, b) { return String(a.name).localeCompare(String(b.name), 'ja'); };
      if (!byId.GroupCompanyId) {
        return groups.filter(function (g) { return !SYSTEM_GROUP.test(g.id); }).sort(byName)
          .map(function (g) { return { id: g.id, label: g.name || g.id }; });
      }
      var out = [], seen = {};
      (function walk(id, depth) {
        if (seen[id] || !byId[id]) return;
        seen[id] = 1;
        var isRoot = id === 'GroupCompanyId';
        if (!isRoot && SYSTEM_GROUP.test(id)) return;   // システムのグループは配下ごと除外
        if (!isRoot) out.push({ id: id, label: Array(depth + 1).join('　') + (byId[id].name || id) });
        (byId[id].children || []).map(function (c) { return byId[c.id]; }).filter(Boolean).sort(byName)
          .forEach(function (k) { walk(k.id, isRoot ? 0 : depth + 1); });
      })('GroupCompanyId', 0);
      return out;
    }

    var rulesCache = null;

    // 集計に使う項目とルール。ルールが未指定（モービルアイなど）の項目は、
    // 名前が「ME…」で項目のキーワードに合うルールを自動で全部使う（例：車線逸脱の右・左を合算）
    function categoriesFor(groupId) {
      return activeCategories(groupId).map(function (c) {
        var cfg = CONFIG.categories.filter(function (x) { return x.key === c.key; })[0];
        if (!c.overridden && !c.ruleIds.length && rulesCache && cfg.hint) {
          c.ruleIds = rulesCache.filter(function (r) {
            var n = r.name || '';
            return CONFIG.meRule.test(n) && cfg.hint.test(n) && ruleFitsGroup(r);
          }).map(function (r) { return r.id; });
          c.auto = true;
        }
        return c;
      });
    }

    function loadGroups() {
      if (groupsLoaded) return Promise.resolve();
      return multiCall(api, [['Get', { typeName: 'Group' }], ['Get', { typeName: 'Rule' }]]).then(function (res) {
        groupsLoaded = true;
        rulesCache = res[1] || [];
        (res[0] || []).forEach(function (g) { groupById[g.id] = g; });
        fillSelect($('tk-group'), 'すべて', groupOptions(res[0] || []));
        var saved = store.get(GROUP_STORE_KEY);
        if (saved && Array.prototype.some.call($('tk-group').options, function (o) { return o.value === saved; })) {
          $('tk-group').value = saved;
        }
      }).catch(function () { /* グループが取れなくても一覧は表示する */ });
    }

    function render() {
      if (!current) return;
      var filter = $('tk-filter').value;
      var all = current.rows;
      var rows = filteredRows();
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
        var idx = all.indexOf(r);
        var changes = r.changes.length
          ? '<ul class="tk-changes">' + r.changes.map(function (c) { return '<li class="' + c.trend + '">' + esc(c.text) + '</li>'; }).join('') + '</ul>'
          : '<span class="tk-none">—</span>';
        return '<tr class="st-' + r.status + (n.checked ? ' is-checked' : '') + '" data-idx="' + idx + '">' +
          '<td><button type="button" class="tk-name tk-name-link" title="過去' + (CONFIG.baselineDays + 1) + '日間のグラフを表示">' + esc(r.name) + '</button>' +
          '<div class="tk-sub">前日 ' + Math.round(r.dayKm) + ' km</div>' +
          '<div class="tk-sub">' + CONFIG.baselineDays + '日計 ' + Math.round(r.baseKm) + ' km</div></td>' +
          r.cats.filter(function (c) { return c.group !== 'me'; }).map(cellHtml).join('') + meCellHtml(r) +
          '<td><span class="tk-badge ' + r.status + '">' + STATUS[r.status].label + '</span></td>' +
          '<td>' + changes + '</td>' +
          '<td><textarea class="tk-comment' + (n.edited ? ' edited' : '') + '" aria-label="' + esc(r.name) + 'の点呼コメント">' + esc(n.comment) + '</textarea>' +
          '<div class="print-only">' + esc(n.comment) + '</div>' +
          '<button type="button" class="tk-reset"' + (n.edited ? '' : ' hidden') + '>定型コメントに戻す</button></td>' +
          '<td><input type="checkbox" class="tk-check" aria-label="' + esc(r.name) + 'を確認済にする"' + (n.checked ? ' checked' : '') + '></td>' +
          '</tr>';
      }).join('');

      $('tk-body').innerHTML = html || '<tr><td colspan="9" class="tk-none">表示する対象がありません。</td></tr>';
      Array.prototype.forEach.call(document.querySelectorAll('#tenko .tk-comment'), autoGrow);
    }

    // ===== 過去8日間のグラフ（ポップアップ）=================================
    var WEEK = ['日', '月', '火', '水', '木', '金', '土'];
    var CHART = { past: '#86b6ef', target: '#2a78d6', grid: '#e6e9ef', axis: '#c5cfdc', ink: '#1c2533', muted: '#5f6b7a' };

    // 目盛りが整数のきりのいい値になるよう最大値と間隔を決める
    function niceScale(maxV, integer) {
      if (maxV <= 0) return { max: integer ? 1 : 10, step: integer ? 1 : 5 };
      if (integer && maxV <= 4) return { max: maxV, step: 1 };
      var raw = maxV / 4, mag = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10));
      var step = [1, 2, 5, 10].map(function (m) { return m * mag; }).filter(function (s) { return s >= raw; })[0];
      if (integer) step = Math.max(1, Math.round(step));
      return { max: Math.ceil(maxV / step) * step, step: step };
    }

    function barPath(x, y, w, h) {   // 上端だけ角丸（4px）、下端は基準線に接地
      var r = Math.min(4, h, w / 2);
      return 'M' + x + ',' + (y + h) + 'V' + (y + r) + 'Q' + x + ',' + y + ' ' + (x + r) + ',' + y +
        'H' + (x + w - r) + 'Q' + (x + w) + ',' + y + ' ' + (x + w) + ',' + (y + r) + 'V' + (y + h) + 'Z';
    }

    // 1指標×8日の棒グラフ。前日だけ濃い青、比較の7日間は薄い青
    function weekChart(values, days, unit, integer) {
      var W = 320, H = 168, ml = 34, mr = 8, mt = 18, mb = 30;
      var pw = W - ml - mr, ph = H - mt - mb, n = values.length, band = pw / n, bw = Math.max(6, band - 8);
      var sc = niceScale(Math.max.apply(null, values), integer);
      var Y = function (v) { return mt + ph - v / sc.max * ph; };
      var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" class="tk-chart" role="img">';
      for (var t = 0; t <= sc.max + 1e-9; t += sc.step) {
        s += '<line x1="' + ml + '" x2="' + (W - mr) + '" y1="' + Y(t) + '" y2="' + Y(t) + '" stroke="' + (t === 0 ? CHART.axis : CHART.grid) + '" stroke-width="1"/>' +
          '<text x="' + (ml - 6) + '" y="' + (Y(t) + 3.5) + '" text-anchor="end" font-size="10" fill="' + CHART.muted + '">' + (Math.round(t * 10) / 10) + '</text>';
      }
      values.forEach(function (v, i) {
        var x = ml + i * band + (band - bw) / 2, last = i === n - 1;
        var h = Y(0) - Y(v);
        if (v > 0) s += '<path d="' + barPath(x, Y(v), bw, Math.max(h, 1)) + '" fill="' + (last ? CHART.target : CHART.past) + '"/>';
        if (last) {   // 値ラベルは前日だけ（全部に付けると読みにくい）
          s += '<text x="' + (x + bw / 2) + '" y="' + (Y(v) - 5) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + CHART.ink + '">' +
            (integer ? v : Math.round(v)) + '</text>';
        }
        s += '<text x="' + (ml + i * band + band / 2) + '" y="' + (H - mb + 14) + '" text-anchor="middle" font-size="10" ' +
          (last ? 'font-weight="700" fill="' + CHART.ink + '"' : 'fill="' + CHART.muted + '"') + '>' + (last ? '前日' : days[i].short) + '</text>' +
          '<text x="' + (ml + i * band + band / 2) + '" y="' + (H - mb + 26) + '" text-anchor="middle" font-size="9" fill="' + CHART.muted + '">' + days[i].wd + '</text>';
        // マウスを乗せた時の当たり判定（棒より広く、プロットの高さいっぱい）
        s += '<rect class="tk-hit" x="' + (ml + i * band) + '" y="' + mt + '" width="' + band + '" height="' + ph + '" fill="transparent" ' +
          'data-tip="' + esc(days[i].label + '：' + (integer ? v : Math.round(v)) + unit) + '"/>';
      });
      return s + '</svg>';
    }

    function openChart(row) {
      var n = CONFIG.baselineDays + 1;
      var days = [];
      for (var i = 0; i < n; i++) {
        var d = addDays(current.baseStart, i);
        days.push({ short: md(d), wd: WEEK[d.getDay()], label: md(d) + '（' + WEEK[d.getDay()] + '）' + (i === n - 1 ? '・前日' : '') });
      }
      var panel = function (title, sub, svg) {
        return '<section class="tk-mpanel"><h3>' + esc(title) + '</h3><p class="tk-msub">' + sub + '</p>' + svg + '</section>';
      };
      var html = '<div class="tk-mgrid">' +
        panel('走行距離', '前日 ' + Math.round(row.dayKm) + ' km ／ 7日計 ' + Math.round(row.baseKm) + ' km',
          weekChart(row.kmByDay, days, ' km', false));
      var unset = unsetKeys();
      var cats = row.cats.filter(function (c) { return !unset[c.key]; });   // ルール未設定の項目は出さない
      cats.forEach(function (c) {
        var arrow = c.trend === 'up' ? ' <span class="tk-up">▲ 増加</span>' : c.trend === 'down' ? ' <span class="tk-down">▼ 減少</span>' : '';
        html += panel(c.label + '（回数）', '前日 ' + c.day + '回（' + fmt1(c.dayRate) + '/100km）・平均 ' + fmt1(c.baseRate) + '/100km' + arrow,
          weekChart(c.byDay, days, '回', true));
      });
      html += '</div>';

      // 表で見る（色だけに頼らない・数値を確認できるように）
      html += '<table class="tk-mtable"><thead><tr><th>日付</th><th>走行距離</th>' +
        cats.map(function (c) { return '<th>' + esc(c.label) + '</th>'; }).join('') + '</tr></thead><tbody>';
      days.forEach(function (d, i) {
        html += '<tr' + (i === n - 1 ? ' class="is-target"' : '') + '><td>' + esc(d.label) + '</td><td>' + Math.round(row.kmByDay[i]) + ' km</td>' +
          cats.map(function (c) { return '<td>' + c.byDay[i] + '回</td>'; }).join('') + '</tr>';
      });
      html += '</tbody></table>' +
        '<p class="tk-msub">棒グラフは日ごとの回数です。判定（注意・改善）は走行100kmあたりの回数で比べています。</p>';

      $('tk-modal-title').textContent = row.name + '　過去' + n + '日間の推移';
      $('tk-modal-body').innerHTML = html;
      $('tk-modal').hidden = false;
      $('tk-modal-close').focus();
    }

    function closeChart() {
      $('tk-modal').hidden = true;
      $('tk-tip').hidden = true;
    }

    var RULE_HINT = /ブレーキ|制動|加速|発進|速度|スピード|brak|accel|jackrabbit|speed/i;

    // ルール名の末尾が「 - 会社名」（例：急ブレーキ - 三井物産）のルールは、その会社専用とみなす。
    // 選んだグループ、またはその上の階層のグループ名に会社名が含まれる時だけ候補に出す
    var COMPANY_SUFFIX = /\s[-－–—]\s*([^-－–—]+)$/;
    function selectedGroupNames() {
      var gid = $('tk-group').value, names = [];
      if (!gid) return names;
      var parentOf = {};
      Object.keys(groupById).forEach(function (pid) {
        (groupById[pid].children || []).forEach(function (c) { parentOf[c.id] = pid; });
      });
      for (var id = gid, n = 0; id && groupById[id] && n < 20; id = parentOf[id], n++) {
        if (groupById[id].name) names.push(groupById[id].name);
      }
      return names;
    }
    function ruleFitsGroup(rule) {
      var m = COMPANY_SUFFIX.exec(rule.name || '');
      if (!m) return true;   // 会社名の付かないルールはどのグループでも表示
      var company = m[1].trim();
      return selectedGroupNames().some(function (n) { return n.indexOf(company) >= 0; });
    }

    function renderDiag() {
      var box = $('tk-diag');
      if (!current) { box.hidden = true; return; }
      var note = $('tk-diag-note');
      var allRules = (current.rules || []).slice().sort(function (a, b) {
        return String(a.name).localeCompare(String(b.name), 'ja');
      });
      var nameOf = {};
      allRules.forEach(function (r) { nameOf[r.id] = r.name; });
      var unit = $('tk-unit').value;
      var days = CONFIG.baselineDays + 1;

      // ルールが見つからない項目（＝設定の見直しが必要）
      var missing = current.diag.filter(function (d) {
        return d.ruleIds.some(function (x) { return !nameOf[x]; });
      });
      var total = 0, known = 0;
      current.diag.forEach(function (d) { total += d.total; known += d.known; });
      var weakDriverId = total > 0 && known / total < 0.5;

      // ---- お知らせ（必要な時だけ、「判定に使うルール」欄の右側に短く表示）----
      var notes = [];
      if (missing.length) {
        notes.push('<p>⚠ <strong>「' + esc(missing.map(function (d) { return d.label; }).join('」「')) +
          '」のルールが見つかりません。</strong>左で選んでください</p>');
      }
      if (weakDriverId) {
        var head = '<p>⚠ <strong>ドライバー識別が不足</strong>（' + total + '件中 ' + known + '件のみ判明）';
        notes.push(unit === 'device'
          ? head + 'のため、車両別で表示中</p>' +
            '<button type="button" class="tk-btn tk-primary tk-btn-sm" data-unit="driver">ドライバー別で表示</button>'
          : head + '。車両別がおすすめです</p>' +
            '<button type="button" class="tk-btn tk-primary tk-btn-sm" data-unit="device">車両別で表示</button>');
      }
      note.hidden = true;   // 以前の一覧上のお知らせ欄は使わない
      var noteHtml = notes.length ? '<div class="tk-rulebar-note">' + notes.join('') + '</div>' : '';

      // ---- 判定に使うルール（急ブレーキ・急加速・速度超過のプルダウン）----
      var ruleLabel = function (d) {
        var id = d.ruleIds[0];
        var catCfg = CONFIG.categories.filter(function (c) { return c.key === d.key; })[0];
        var def = catCfg.ruleIds;
        var autoIds = d.auto ? d.ruleIds : [];
        var list = allRules.filter(function (r) {
          if (d.overridden && r.id === id) return true;   // 選択中のルールは常に残す
          return ruleFitsGroup(r) && (showAllRules || (catCfg.hint || RULE_HINT).test(r.name || ''));
        });
        // 先頭の選択肢：標準ルール（急ブレーキ等）／自動で選んだルール（モービルアイ）
        var first = def.length
          ? '標準：' + def.map(function (x) { return nameOf[x] || x + '（見つかりません）'; }).join('、')
          : (autoIds.length ? '自動：' + autoIds.map(function (x) { return nameOf[x] || x; }).join('＋') : '（該当するルールなし）');
        var opts = '<option value="">' + esc(first) + '</option>';
        if (d.overridden && !nameOf[id]) opts += '<option value="' + esc(id) + '" selected>' + esc(id) + '（見つかりません）</option>';
        opts += list.filter(function (r) { return def.indexOf(r.id) < 0; }).map(function (r) {
          return '<option value="' + esc(r.id) + '"' + (d.overridden && r.id === id ? ' selected' : '') + '>' + esc(r.name) + '</option>';
        }).join('');
        var count = d.ruleIds.length ? days + '日間 ' + d.total + '件（うちドライバー判明 ' + d.known + '件）' : 'ルール未設定';
        return '<label>' + esc(d.label) +
          '<select class="tk-rule tk-wide" data-key="' + d.key + '"' + (autoIds.length > 1 ? ' title="' + esc(first) + '"' : '') + '>' + opts + '</select>' +
          '<span class="tk-rule-count' + (d.total === 0 ? ' zero' : '') + '">' + esc(count) + '</span></label>';
      };
      var mainHtml = current.diag.filter(function (d) { return d.group !== 'me'; }).map(ruleLabel).join('');
      var meDiag = current.diag.filter(function (d) { return d.group === 'me'; });

      box.hidden = false;
      $('tk-diag-body').innerHTML =
        '<div class="tk-rulebar-row">' + mainHtml +
        '<label class="tk-allrules"><span><input type="checkbox" id="tk-allrules"' + (showAllRules ? ' checked' : '') + '> すべてのルールを候補に表示</span></label>' +
        noteHtml + '</div>' +
        (meDiag.length ? '<div class="tk-rulebar-row tk-rulebar-me"><div class="tk-rulebar-sub">モービルアイ</div>' + meDiag.map(ruleLabel).join('') + '</div>' : '');
    }
    var showAllRules = false;

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
      loadRaw(api, state, target, $('tk-group').value, categoriesFor($('tk-group').value)).then(function (raw) {
        current = aggregate(raw, $('tk-unit').value);
        $('tk-period').textContent = periodText(raw, current.unknownKm);
        showMsg(current.rows.length ? '' : '前日に運行した対象がありません。日付・グループ・集計単位を確認してください。');
        populateFilters();
        render();
        renderDiag();
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
      filteredRows().forEach(function (r) {
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
      $('tk-group').addEventListener('change', function () {
        store.set(GROUP_STORE_KEY, $('tk-group').value);
        $('tk-fdev').value = '';
        $('tk-fdrv').value = '';
        load();
      });
      $('tk-fdev').addEventListener('change', render);
      $('tk-fdrv').addEventListener('change', render);
      $('tk-print').addEventListener('click', function () { window.print(); });
      $('tk-csv').addEventListener('click', exportCsv);

      $('tk-diag-body').addEventListener('click', function (ev) {
        var unitTo = ev.target.dataset && ev.target.dataset.unit;
        if (!unitTo) return;
        $('tk-unit').value = unitTo;
        load();
      });
      $('tk-diag-body').addEventListener('change', function (ev) {
        if (ev.target.id === 'tk-allrules') { showAllRules = ev.target.checked; renderDiag(); return; }
        if (!ev.target.classList.contains('tk-rule')) return;
        var key = ruleStoreKey($('tk-group').value);
        var ov = store.get(key) || {};
        if (ev.target.value) ov[ev.target.dataset.key] = ev.target.value;
        else delete ov[ev.target.dataset.key];
        store.set(key, ov);
        load();
      });

      // 名前クリックでグラフのポップアップ
      $('tk-body').addEventListener('click', function (ev) {
        if (!ev.target.closest('.tk-name-link')) return;
        var x = rowFromEvent(ev); if (!x) return;
        openChart(x.row);
      });
      $('tk-modal-close').addEventListener('click', closeChart);
      $('tk-modal').addEventListener('click', function (ev) { if (ev.target.id === 'tk-modal') closeChart(); });
      document.addEventListener('keydown', function (ev) { if (ev.key === 'Escape' && !$('tk-modal').hidden) closeChart(); });
      $('tk-modal-body').addEventListener('mousemove', function (ev) {
        var tip = $('tk-tip');
        if (!ev.target.classList || !ev.target.classList.contains('tk-hit')) { tip.hidden = true; return; }
        var box = $('tk-modal-box').getBoundingClientRect();
        tip.textContent = ev.target.getAttribute('data-tip');
        tip.hidden = false;
        tip.style.left = (ev.clientX - box.left + 12) + 'px';
        tip.style.top = (ev.clientY - box.top + $('tk-modal-box').scrollTop - 30) + 'px';
      });
      $('tk-modal-body').addEventListener('mouseleave', function () { $('tk-tip').hidden = true; });

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
        loadGroups().then(load);
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
      { last: '高橋', first: '健',   plate: '品川 100 あ 12-04', base: { brake: 0.6, accel: 0.5, speed: 0.6 }, day: { accel: 2.6, speed: 2.0, hmw: 2.5 } },
      { last: '田中', first: '誠',   plate: '品川 100 あ 12-05', base: { brake: 0.7, accel: 0.6, speed: 0.8 }, day: { brake: 0.8, accel: 0.6, speed: 0.8 } },
      { last: '伊藤', first: '美咲', plate: '品川 100 あ 12-06', base: { brake: 0.5, accel: 0.5, speed: 0.5 }, day: {}, dayKm: 12 },
      { last: '渡辺', first: '翔',   plate: '品川 100 あ 12-07', base: { brake: 0.9, accel: 0.4, speed: 0.6 }, day: { brake: 1.0, accel: 0.5, speed: 0.6 } },
      { last: '中村', first: '大輔', plate: '品川 100 あ 12-08', base: { brake: 1.8, accel: 0.6, speed: 0.4 }, day: { brake: 0, accel: 0.6 } }
    ];
    var ruleOf = { brake: 'RuleHarshBrakingId', accel: 'RuleJackrabbitStartsId', speed: 'RulePostedSpeedingId',
      fcw: 'demo-me-fcw', hmw: 'demo-me-hmw', ldw: 'demo-me-ldw-r', pcw: 'demo-me-pcw' };
    var ME_BASE = { fcw: 0.1, hmw: 0.4, ldw: 0.3, pcw: 0.05 };   // モービルアイの平常時の発生率（デモ）

    function build(from, to) {
      var target = addDays(new Date(to.getTime() - 1), 0);
      target.setHours(0, 0, 0, 0);
      var out = { trips: [], ex: {} };
      for (var d = new Date(from); d < to; d = addDays(d, 1)) {
        var isTarget = d.getTime() === target.getTime();
        people.forEach(function (p, i) {
          var rng = mulberry32(hash(ymd(d) + ':' + i));
          var km = isTarget && p.dayKm ? p.dayKm : Math.round(110 + rng() * 120);
          // ?demo=nodriver：ドライバー識別が無い状態を再現
          var drv = { id: /nodriver/.test(location.search) && i % 4 ? 'UnknownDriverId' : 'demo-u' + i }, dev = { id: 'demo-d' + i };
          var t1 = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 7, 30);
          var t2 = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 13, 10);
          out.trips.push({ device: dev, driver: drv, start: t1.toISOString(), distance: km * 0.55 });
          out.trips.push({ device: dev, driver: drv, start: t2.toISOString(), distance: km * 0.45 });
          Object.keys(ruleOf).forEach(function (k) {
            var base = p.base[k] != null ? p.base[k] : ME_BASE[k];
            var rate = isTarget && p.day[k] != null ? p.day[k] : base;
            var n = isTarget && p.day[k] != null ? Math.round(rate * km / 100) : poisson(rng, rate * km / 100);
            for (var j = 0; j < n; j++) {
              var at = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 8 + Math.floor(rng() * 9), Math.floor(rng() * 60));
              // 車線逸脱は右・左の2ルールに分けて記録（実データと同じ形）
              var rid = k === 'ldw' && j % 2 ? 'demo-me-ldw-l' : ruleOf[k];
              (out.ex[rid] = out.ex[rid] || []).push({ rule: { id: rid }, device: dev, driver: drv, activeFrom: at.toISOString() });
            }
          });
        });
      }
      return out;
    }

    return {
      multiCall: function (calls, ok) {
        // デモのグループ：東京営業所＝車両0〜3、名古屋営業所＝車両4〜7
        var inGroup = function (devId, groups) {
          if (!groups || !groups.length) return true;
          var i = +String(devId).replace('demo-d', '');
          return groups.some(function (g) {
            if (g.id === 'demo-g-east' || g.id === 'demo-g-tokyo') return i < 4;
            if (g.id === 'demo-g-west' || g.id === 'demo-g-nagoya') return i >= 4;
            return true;
          });
        };
        var tripCall = calls.filter(function (c) { return c[1].typeName === 'Trip'; })[0];
        var data = tripCall ? build(new Date(tripCall[1].search.fromDate), new Date(tripCall[1].search.toDate)) : null;
        var groupsOf = function (p) { return p.search && (p.search.groups || (p.search.deviceSearch && p.search.deviceSearch.groups)); };
        var res = calls.map(function (c) {
          var p = c[1], gs = groupsOf(p);
          var pick = function (list) { return list.filter(function (x) { return inGroup(idOf(x.device), gs); }); };
          if (p.typeName === 'Group') return [
            { id: 'GroupCompanyId', name: 'Company Group', children: [{ id: 'demo-g-east' }, { id: 'demo-g-west' }, { id: 'GroupAssetTypeId' }] },
            { id: 'GroupAssetTypeId', name: 'Asset Type', children: [{ id: 'GroupVehicleId' }] },
            { id: 'GroupVehicleId', name: 'Vehicle', children: [] },
            { id: 'demo-g-east', name: '関東エリア', children: [{ id: 'demo-g-tokyo' }] },
            { id: 'demo-g-tokyo', name: '東京営業所', children: [] },
            { id: 'demo-g-west', name: '中部エリア', children: [{ id: 'demo-g-nagoya' }] },
            { id: 'demo-g-nagoya', name: '名古屋営業所', children: [] }
          ];
          if (p.typeName === 'Trip') return pick(data.trips);
          if (p.typeName === 'Device') return people.map(function (x, i) { return { id: 'demo-d' + i, name: x.plate }; })
            .filter(function (d) { return inGroup(d.id, gs); });
          if (p.typeName === 'User') return people.map(function (x, i) {
            // 実データと同じく、漢字の名前は firstName に姓・lastName に名が入っている想定
            return { id: 'demo-u' + i, firstName: x.last, lastName: x.first, name: 'demo' + i,
              companyGroups: [{ id: i < 4 ? 'demo-g-tokyo' : 'demo-g-nagoya' }] };
          }).concat([{ id: 'demo-u-extra', firstName: '小林', lastName: '亮', name: 'demo-extra', companyGroups: [{ id: 'demo-g-tokyo' }] }]);
          if (p.typeName === 'ExceptionEvent') return pick(data.ex[p.search.ruleSearch.id] || []);
          if (p.typeName === 'Rule') return [
            { id: 'RuleHarshBrakingId', name: '急ブレーキ' }, { id: 'RuleJackrabbitStartsId', name: '急発進' },
            { id: 'RulePostedSpeedingId', name: '速度超過（制限速度）' }, { id: 'RuleIdlingId', name: 'アイドリング' },
            { id: 'demo-custom-speed', name: '速度超過 80km/h（カスタム）' },
            { id: 'demo-tokyo-brake', name: '急ブレーキ - 東京営業所' },
            { id: 'demo-hb-new', name: 'Harsh Braking (New)' }, { id: 'demo-ha-new', name: 'Harsh Acceleration (New)' },
            { id: 'demo-sp-new', name: 'Speeding (New)' },
            { id: 'demo-me-fcw', name: 'ME追突警報' }, { id: 'demo-me-ufcw', name: 'ME低速時追突警報' },
            { id: 'demo-me-hmw', name: 'ME車間警報' }, { id: 'demo-me-pcw', name: 'ME歩行者警報' },
            { id: 'demo-me-ldw-r', name: 'ME車線逸脱警報-右' }, { id: 'demo-me-ldw-l', name: 'ME車線逸脱警報-左' },
            { id: 'demo-me-tamper', name: 'MEタンパーアラート' }
          ];
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
