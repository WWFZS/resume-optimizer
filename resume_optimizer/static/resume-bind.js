/* ════════════════════════════════════════════════════════════════════
 * resume-bind.js — 真实 Word 模板盒子 ⇄ 统一简历 JSON（S.resume）槽位映射
 * ──────────────────────────────────────────────────────────────────
 * 「换模板内容跟随」的核心：S.resume 永远是唯一真相，Word 模板 = 几何骨架
 * + 槽位绑定。box.id 是位置寻址、跨模板不稳定，所以绑定一律按
 * 「占位符语义 > role 声明 + role 内 y 序」计算，绝不按 id 跨模板搬运。
 *
 *   buildBindings(structure) → { boxId: {mode:'basic'|'role', role?, ordinal?} }
 *   seedEdits(structure, resume) → { boxId: {paragraphs:[…]} }
 *       预填：basic 盒按 {{占位符}} 与「标签：值」双路替换（只换有值的）；
 *       板块盒按 role 内序号灌第 N 条记录；序号超出的板块盒留白
 *       （有数据填充、无数据留白，绝不把模板样例当用户内容导出）。
 *   applyBoxEdit(structure, resume, boxId, text) → bool
 *       回写：解析「标签：值」→ basic；「头行 + 正文」→ 结构化记录字段。
 *       解析保守、字段合并式写入，整段文字永不丢失。
 *   coverage(structure, resume) → { role: {records, boxes} }（放不下提示）
 *
 * 语义来源：盒子 {{占位符}} 中文标签（最可靠，实测 role 几何标注存在误标，
 * 如工作经历盒被标成 education 但 placeholders=["工作经历"]）+ role 声明。
 * 与归档 editor2 transferToTemplate 的 role/stdKey 机制同源，中枢换 resume。
 * 纯函数模块：浏览器挂 window.ResumeBind，node 挂 module.exports（自检）。
 * ════════════════════════════════════════════════════════════════════ */
(function (root) {
  'use strict';

  /* ── 板块词表（与 editor-v3.js 的 LIST_KINDS / 渲染字段对齐）── */
  var LIST_ROLES = ['education', 'internship', 'work', 'project', 'campus', 'skill', 'honor'];
  var TEXT_ROLES = ['self', 'custom', 'hobby'];

  /* 基本信息中文标签 → basic 字段（刻意保守，未命中不动作） */
  var BASIC_LABELS = {
    '姓名': 'name', '名字': 'name',
    '求职意向': 'intention', '意向职位': 'intention', '应聘岗位': 'intention', '目标岗位': 'intention', '期望职位': 'intention',
    '电话': 'phone', '手机': 'phone', '手机号码': 'phone', '联系电话': 'phone',
    '邮箱': 'email', '电子邮箱': 'email', 'Email': 'email', 'E-mail': 'email',
    '微信': 'wechat',
    '现居': 'city', '所在城市': 'city', '现居城市': 'city', '居住地': 'city',
    '性别': 'gender', '年龄': 'age', '生日': 'birth_date', '出生日期': 'birth_date', '出生年月': 'birth_date',
    '籍贯': 'hometown', '民族': 'nation', '身高': 'height', '体重': 'weight',
    '政治面貌': 'political_status', '婚姻状况': 'marriage',
    '学历': 'education_degree', '毕业院校': 'school', '专业': 'major',
    '期望薪资': 'salary', '到岗时间': 'arrive_time', '工作年限': 'work_years'
  };
  /* 占位符/role 别名 → 标准语义（null = 装饰性标签，不绑定） */
  var SEMANTIC_ALIAS = {
    '意向岗位': 'intention', '求职意向岗位': 'intention', '联系方式': 'phone', '手机号': 'phone',
    '邮箱地址': 'email', '微信号': 'wechat',
    '教育背景': 'education', '教育经历': 'education', '工作经历': 'work', '工作经验': 'work',
    '实习经历': 'internship', '实习经验': 'internship', '项目经历': 'project', '项目经验': 'project',
    '校园经历': 'campus', '校园经验': 'campus', '社团活动': 'campus',
    '技能特长': 'skill', '专业技能': 'skill', '特长技能': 'skill', '技能': 'skill',
    '荣誉证书': 'honor', '奖项证书': 'honor', '荣誉奖项': 'honor', '证书': 'honor',
    '个人优势': 'self', '自我评价': 'self', '自我介绍': 'self',
    '兴趣爱好': 'hobby', '爱好': 'hobby',
    '基本信息': null, '个人信息': null
  };
  var ROLE_SET = {};
  LIST_ROLES.concat(TEXT_ROLES).forEach(function (r) { ROLE_SET[r] = true; });

  var DEGREE_WORDS = ['博士', '硕士', '研究生', '本科', '大专', '中专', '高中', 'MBA', 'EMBA'];

  /* ── 小工具 ── */
  function clean(s) { return String(s == null ? '' : s).replace(/\u00a0/g, ' ').trim(); }
  function para(text, align) {
    return { paragraphs: [{ align: align || 'left', line: null,
      runs: [{ text: String(text == null ? '' : text) }] }] };
  }
  function basicKey(label) {
    var key = clean(label);
    if (BASIC_LABELS[key]) return BASIC_LABELS[key];
    if (SEMANTIC_ALIAS[key]) key = SEMANTIC_ALIAS[key];
    if (BASIC_LABELS[key]) return BASIC_LABELS[key];
    var labels = Object.keys(BASIC_LABELS);
    for (var i = 0; i < labels.length; i++) {
      if (key.length <= labels[i].length + 4 && key.indexOf(labels[i]) >= 0) return BASIC_LABELS[labels[i]];
    }
    return null;
  }
  /* 一个标签/占位符 → 语义：{role} 或 {basic} 或 null */
  function semantic(label) {
    var key = clean(label);
    if (SEMANTIC_ALIAS[key] === null) return null;
    var r = SEMANTIC_ALIAS[key];
    if (r && ROLE_SET[r]) return { role: r };
    if (r && r !== key) {
      var bk = BASIC_LABELS[r];
      if (bk) return { basic: bk };
    }
    if (ROLE_SET[key]) return { role: key };
    var b = basicKey(key);
    return b ? { basic: b } : null;
  }

  /* ══════════════════ 绑定表：seed 与 writeback 共用的唯一依据 ══════════════════ */
  function buildBindings(structure) {
    var map = {};
    var buckets = {};
    var boxes = (structure && structure.boxes) || [];
    boxes.forEach(function (b) {
      if (b.kind === 'title') return;
      var phRole = null, phBasic = [];
      (b.placeholders || []).forEach(function (p) {
        var s = semantic(p);
        if (!s) return;
        if (s.role && !phRole) phRole = s.role;
        if (s.basic && phBasic.indexOf(s.basic) < 0) phBasic.push(s.basic);
      });
      var declared = clean(b.role);
      /* 占位符语义 > role 声明（实测 role 有误标，占位符是模板作者写的） */
      var role = phRole || (ROLE_SET[declared] ? declared : null);
      var isField = b.kind === 'field' || b.kind === 'photo';
      if (role) {
        (buckets[role] = buckets[role] || []).push(b);
      } else if (isField || phBasic.length > 0 ||
                 (declared === 'basic' && /\{\{|[：:]/.test(String(b.text || '')))) {
        map[b.id] = { mode: 'basic' };
      }
      /* 其余（自由装饰文本、无语义可辨的 free 盒）：不绑定，保留模板原文 */
    });
    Object.keys(buckets).forEach(function (role) {
      var list = buckets[role].sort(function (a, b2) { return (a.y || 0) - (b2.y || 0); });
      if (TEXT_ROLES.indexOf(role) >= 0) {
        /* 文本类板块：优先绑带语义占位符的盒；无占位符且同 role 多盒 → 全部
           不绑（实测「软件技能」盒被误标 self，宁可不填也不灌错位置） */
        var withPh = list.filter(function (b) {
          return (b.placeholders || []).some(function (p) { var s = semantic(p); return s && s.role === role; });
        });
        var chosen = withPh.length ? withPh : (list.length === 1 ? list : []);
        chosen.forEach(function (b) { map[b.id] = { mode: 'role', role: role, ordinal: 0 }; });
        return;
      }
      list.forEach(function (b, i) { map[b.id] = { mode: 'role', role: role, ordinal: i }; });
    });
    return map;
  }

  /* ── 记录读取与文本化 ── */
  function periodOf(rec) {
    var s = clean(rec.start), e = rec.current ? '至今' : clean(rec.end);
    if (s && e) return s + ' - ' + e;
    return s || e || clean(rec.period || '');
  }
  function recordLines(rec, role) {
    if (!rec || typeof rec !== 'object') {
      var raw = clean(rec);
      return raw ? [raw] : null;
    }
    var head = [], lines = [];
    if (role === 'education') {
      if (rec.school) head.push(clean(rec.school));
      var who = [clean(rec.major), rec.degree ? '（' + clean(rec.degree) + '）' : ''].join('');
      if (who) head.push(who);
    } else if (role === 'work' || role === 'internship') {
      if (rec.company) head.push(clean(rec.company));
      if (rec.position) head.push(clean(rec.position));
    } else if (role === 'project' || role === 'campus') {
      if (rec.name) head.push(clean(rec.name));
      if (rec.role) head.push(clean(rec.role));
    } else if (role === 'skill') {
      var sk = clean(rec.name);
      if (sk) { if (rec.level) sk += '（' + clean(rec.level) + '）'; head.push(sk); }
    } else if (role === 'honor') {
      var ho = clean(rec.name);
      if (ho) {
        if (rec.issuer) ho += '（' + clean(rec.issuer) + '）';
        if (clean(rec.time)) ho = clean(rec.time) + '  ' + ho;
        head.push(ho);
      }
    }
    var per = periodOf(rec);
    if (per) head.push(per);
    head = head.filter(Boolean);
    if (head.length) lines.push(head.join('    '));
    var body = clean(rec.content);
    if (body) lines.push(body);
    return lines.length ? lines : null;
  }
  function moduleValue(resume, role) {
    var m = (resume.modules || {})[role];
    if (role === 'self' || role === 'custom') return clean(m);
    if (role === 'hobby') return Array.isArray(m) ? m.map(clean).filter(Boolean).join('、') : clean(m);
    return Array.isArray(m) ? m : [];
  }
  function recordCount(resume, role) {
    var val = moduleValue(resume, role);
    if (typeof val === 'string') return val ? 1 : 0;
    var n = 0;
    val.forEach(function (r) {
      if (typeof r !== 'object' || recordLines(r, role)) n++;
    });
    return n;
  }

  /* ══════════════════ seedEdits：resume → 盒子编辑增量 ══════════════════ */
  function seedEdits(structure, resume) {
    var edits = {};
    if (!structure || !resume) return edits;
    var bindings = buildBindings(structure);
    var boxes = structure.boxes || [];
    var basic = resume.basic || {};

    boxes.forEach(function (b) {
      var bind = bindings[b.id];
      if (!bind) return;
      var align = b.paras && b.paras[0] && b.paras[0].align;

      if (bind.mode === 'basic') {
        var text = String(b.text || '');
        var changed = false;
        /* 「标签：值」行：有值替换、无值清空 —— 用户简历是唯一真相，
           模板样例值（生日 1994.04.06 之类）绝不冒充用户内容导出 */
        var lines = text.split('\n');
        lines = lines.map(function (ln) {
          var m = /^([^：:]{1,8})[：:]\s*(.*)$/.exec(ln);
          if (!m) return ln;
          var key = basicKey(m[1]);
          if (!key) return ln;
          changed = true;
          return m[1] + '：' + clean(basic[key]);
        });
        text = lines.join('\n');
        /* {{占位符}}：有值替换、无值剥除（本盒已被改写时不留 {{}} 残片） */
        text = text.replace(/\{\{([^}]+)\}\}/g, function (whole, ph) {
          var s = semantic(ph);
          var key = s && s.basic;
          if (!key) return '';
          var v = clean(basic[key]);
          changed = true;
          return v;
        });
        if (changed) edits[b.id] = para(text, align);
        return;
      }

      /* mode === 'role'：第 N 盒 ← 第 N 条记录；序号超出 → 留白（不导出样例） */
      var cnt = recordCount(resume, bind.role);
      if (bind.ordinal >= cnt) {
        if (String(b.text || '').replace(/\{\{[^}]*\}\}/g, '').trim()) {
          edits[b.id] = para('', align);
        }
        return;
      }
      var val = moduleValue(resume, bind.role);
      var outLines = null;
      if (typeof val === 'string') outLines = val ? [val] : null;
      else outLines = recordLines(val[bind.ordinal], bind.role);
      if (outLines && outLines.length) {
        edits[b.id] = { paragraphs: outLines.map(function (t) {
          return { align: align || 'left', line: null, runs: [{ text: t }] };
        }) };
      }
    });
    return edits;
  }

  /* ══════════════════ 文本 → 结构化回写 ══════════════════ */
  var DATE_RANGE_RE = /(\d{4})\s*[年.\/\-]\s*(\d{1,2})?\s*月?\s*(?:-|—|~|至|到)\s*((?:\d{4}\s*[年.\/\-]?\s*\d{0,2}\s*月?)|至今|今)/;
  var YEAR_HEAD_RE = /^\d{4}\s*年[^，。,]{0,6}$/;
  function parseRecordText(text, role, prev) {
    var raw = String(text || '').replace(/\u00a0/g, ' ');
    var lines = raw.split('\n').map(clean).filter(Boolean);
    var rec = prev ? JSON.parse(JSON.stringify(prev)) : {};
    if (!lines.length) return rec;   /* 清空盒子不动记录（用户可能只是删着玩） */
    var head = lines[0], body = lines.slice(1).join('\n');
    var m = DATE_RANGE_RE.exec(head), start = '', end = '', current = false;
    if (m) {
      start = m[1] + (m[2] ? '-' + ('0' + m[2]).slice(-2) : '');
      if (/至今|今/.test(m[3])) current = true;
      else {
        var em = /(\d{4})\s*[年.\/\-]?\s*(\d{1,2})?/.exec(m[3]);
        if (em) end = em[1] + (em[2] ? '-' + ('0' + em[2]).slice(-2) : '');
      }
      head = clean(head.replace(m[0], ''));
    }
    var parts = head.split(/\s{2,}|[|｜]/).map(clean).filter(Boolean);
    /* 模板常见「2014年寒假  公司  职位」：首段是年+假期的，不当公司名 */
    if (parts.length >= 2 && YEAR_HEAD_RE.test(parts[0])) {
      if (!start) start = parts[0];
      parts = parts.slice(1);
    }
    if (role === 'education') {
      if (parts[0]) rec.school = parts[0];
      if (parts[1]) {
        var deg = null;
        DEGREE_WORDS.forEach(function (d) { if (!deg && parts[1].indexOf(d) >= 0) deg = d; });
        if (deg) {
          rec.degree = deg;
          rec.major = clean(parts[1].split(deg).join('').replace(/[（()）\s]+/g, ' '));
        } else rec.major = parts[1];
      }
    } else if (role === 'work' || role === 'internship') {
      if (parts[0]) rec.company = parts[0];
      if (parts[1]) rec.position = parts[1];
    } else if (role === 'project' || role === 'campus') {
      if (parts[0]) rec.name = parts[0];
      if (parts[1]) rec.role = parts[1];
    } else if (role === 'skill') {
      if (parts[0]) {
        var lm = /^(.*?)[（(]([^（）()]+)[)）]$/.exec(parts[0]);
        if (lm) { rec.name = lm[1]; rec.level = lm[2]; } else rec.name = parts[0];
      }
    } else if (role === 'honor') {
      var tm = /^(\d{4}[年.\/\-]?\d{0,2}月?\s{1,})(.*)$/.exec(head);
      var nameSrc = tm ? clean(tm[2]) : (parts[0] || '');
      if (tm) rec.time = clean(tm[1]);
      var nm = /^(.*?)[（(]([^（）()]+)[)）]$/.exec(nameSrc);
      if (nameSrc) { rec.name = nameSrc ? (nm ? clean(nm[1]) : nameSrc) : ''; if (nm) rec.issuer = nm[2]; }
    }
    if (start) rec.start = start;
    if (end) rec.end = end;
    if (m) rec.current = current;
    if (body) rec.content = body;
    else if (!parts.length && !m) rec.content = clean(raw);
    return rec;
  }

  function applyBoxEdit(structure, resume, boxId, text) {
    var bindings = buildBindings(structure);
    var bind = bindings[boxId];
    if (!bind) return false;
    resume.modules = resume.modules || {};
    resume.basic = resume.basic || {};
    var t = String(text == null ? '' : text).replace(/\u00a0/g, ' ');

    if (bind.mode === 'basic') {
      var hitAny = false;
      t.split('\n').forEach(function (ln) {
        var m = /^([^：:]{1,8})[：:]\s*(.*)$/.exec(clean(ln));
        if (!m) return;
        var key = basicKey(m[1]);
        if (!key) return;
        hitAny = true;
        resume.basic[key] = clean(m[2]);
      });
      if (!hitAny) {
        var one = clean(t);
        var box = ((structure.boxes || []).filter(function (b) { return String(b.id) === String(boxId); })[0] || {});
        var ph = (box.placeholders || []).join(' ');
        if (one && t.indexOf('\n') < 0 && /姓名|名字/.test(ph + ' ' + String(box.text || ''))) {
          resume.basic.name = one; hitAny = true;
        }
      }
      return hitAny;
    }

    var role = bind.role, mods = resume.modules;
    if (role === 'self' || role === 'custom') { mods[role] = clean(t); return true; }
    if (role === 'hobby') {
      mods.hobby = t.split(/[\n、,，]/).map(clean).filter(Boolean);
      return true;
    }
    /* LIST 角色：role 内序号定位记录；数组不足时补齐（永不产生空洞） */
    var arr = Array.isArray(mods[role]) ? mods[role] : (mods[role] = []);
    while (arr.length <= bind.ordinal) arr.push({});
    var target = arr[bind.ordinal];
    var parsed = parseRecordText(t, role, target);
    Object.keys(parsed).forEach(function (k) { target[k] = parsed[k]; });
    target.index = bind.ordinal;
    return true;
  }

  /* ══════════════════ coverage：放不下 / 没地方填的提示依据 ══════════════════ */
  function coverage(structure, resume) {
    var out = {};
    if (!structure || !resume) return out;
    var bindings = buildBindings(structure);
    var boxCount = {};
    Object.keys(bindings).forEach(function (id) {
      var bind = bindings[id];
      if (bind.mode !== 'role') return;
      if (TEXT_ROLES.indexOf(bind.role) >= 0) { boxCount[bind.role] = 1; return; }
      boxCount[bind.role] = Math.max(boxCount[bind.role] || 0, bind.ordinal + 1);
    });
    LIST_ROLES.concat(TEXT_ROLES).forEach(function (role) {
      var recs = recordCount(resume, role);
      var boxes = boxCount[role] || 0;
      if (recs || boxes) out[role] = { records: recs, boxes: boxes };
    });
    return out;
  }

  var ResumeBind = {
    buildBindings: buildBindings,
    seedEdits: seedEdits,
    applyBoxEdit: applyBoxEdit,
    coverage: coverage,
    _basicKey: basicKey, _semantic: semantic, _parseRecordText: parseRecordText
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = ResumeBind;
  if (root) root.ResumeBind = ResumeBind;
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
