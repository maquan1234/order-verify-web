'use strict';
/*
 * verify_core.js — 审单核对核心逻辑（浏览器/Node 通用，无 DOM / 无 XLSX 依赖）
 * 由 verifier.py 1:1 翻译：货品映射 → sheet0/sheet1 匹配 → 备注解析 → verify_orders → 导出数据
 */

// =================== 工具 ===================
const CN_NUM_MAP = { '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };

function parseCnNumber(text) {
  if (Object.prototype.hasOwnProperty.call(CN_NUM_MAP, text)) return CN_NUM_MAP[text];
  if (text.length === 2 && text[0] === '十' && CN_NUM_MAP[text[1]] !== undefined) return 10 + CN_NUM_MAP[text[1]];
  return null;
}

function safeGet(row, idx, def = '') {
  if (!Array.isArray(row)) return def;
  if (idx < row.length) {
    const val = row[idx];
    return (val !== null && val !== undefined) ? val : def;
  }
  return def;
}

function toStr(v) {
  if (v === null || v === undefined) return '';
  return String(v);
}

// =================== 货品信息 ===================
const SUFFIX_MAP = [['-9', 9], ['-6', 6], ['-3', 3], ['-2', 2], ['-1', 1]];

function buildSingleProductInfo(prod) {
  const nameStr = toStr(prod.name);
  const commonStr = toStr(prod.common);
  const barcodeStr = toStr(prod.barcode).trim();

  let boxCount = 1;
  let baseBarcode = barcodeStr;
  let boxCountFromSuffix = false;
  for (const [suffix, count] of SUFFIX_MAP) {
    if (barcodeStr.endsWith(suffix)) {
      boxCount = count;
      baseBarcode = barcodeStr.slice(0, barcodeStr.length - suffix.length);
      boxCountFromSuffix = true;
      break;
    }
  }

  let innerCount = boxCountFromSuffix ? boxCount : 1;
  if (!boxCountFromSuffix && nameStr) {
    let m = nameStr.match(/[*×](\d+)(?![\.\d])/);
    if (m) {
      innerCount = parseInt(m[1], 10);
    } else {
      m = nameStr.match(/(\d+)\s*(支装|只装)/);
      if (m) innerCount = parseInt(m[1], 10);
    }
  }

  const commonNames = commonStr ? commonStr.split('/').map(c => c.trim()).filter(Boolean) : [];

  let packageType = '瓶装';
  let bagsPerPack = 1;
  if (nameStr.includes('袋装') || /(\d+)袋/.test(nameStr)) {
    packageType = '袋装';
    const m = nameStr.match(/(\d+)\s*袋/);
    if (m) bagsPerPack = parseInt(m[1], 10);
  }

  return {
    name: nameStr,
    code: toStr(prod.code).trim(),
    barcode: barcodeStr,
    common_names: commonNames,
    box_count: boxCount,
    box_count_from_suffix: boxCountFromSuffix,
    inner_count: innerCount,
    base_barcode: baseBarcode,
    package_type: packageType,
    bags_per_pack: bagsPerPack,
  };
}

function buildBarcodeToInfo(products) {
  const m = {};
  for (const prod of products) {
    const b = toStr(prod.barcode).trim();
    if (!b) continue;
    m[b] = buildSingleProductInfo(prod);
  }
  return m;
}

function buildCodeToInfo(products) {
  const m = {};
  for (const prod of products) {
    const c = toStr(prod.code).trim();
    if (!c) continue;
    m[c] = buildSingleProductInfo(prod);
  }
  return m;
}

// =================== 产品分类映射 ===================
function getProductCategory(barcode, barcodeToInfo, code, codeToInfo) {
  let info = barcodeToInfo[barcode];
  if (!info && code && codeToInfo) info = codeToInfo[code];
  if (!info) return '未知';
  const name = info.name;
  const common = info.common_names.join('/');

  if (name.includes('EGCG饮3.0') || common.includes('升级肽3.0')) return '升级肽3.0';
  if (name.includes('EGCG饮2.0') && name.includes('枇杷')) return '升级肽2.0枇杷味';
  if (name.includes('EGCG饮2.0') || common.includes('升级肽2.0')) return '升级肽2.0';
  if (name.includes('EGCG饮1.0') || common.includes('升级肽1.0')) return '升级肽1.0';
  if (name.includes('胶原蛋白肽维C饮品2.0') || name.includes('维C饮2.0')) return '经典肽2.0';
  if (name.includes('胶原蛋白肽饮3.0') || common.includes('经典肽3.0')) return '经典肽3.0';
  if (name.includes('维C饮1.0') || common.includes('经典肽1.0')) return '经典肽1.0';
  if (name.includes('高端肽') || common.includes('高端肽')) return '高端肽';
  if (name.includes('PQQ') || common.includes('PQQ')) return 'PQQ饮';
  if (name.includes('虾青素白松茸') || common.includes('虾青素Pro')) return '虾青素Pro';
  if (name.includes('虾青素凝胶糖果-6粒') || common.includes('虾青素6粒') || common.includes('虾青素试吃')) return '虾青素6粒';
  if (name.includes('虾青素凝胶糖果') && !name.includes('6粒')) return '虾青素1.0';
  if (name.includes('超级零') && name.includes('虾青素')) return '虾青素饮';
  if (name.includes('富铁软糖') || common.includes('软糖')) {
    if (name.includes('降糖') || common.includes('减糖') || common.includes('降糖版')) {
      if (name.includes('6粒')) return '富铁软糖降糖版6粒';
      return '富铁软糖降糖版';
    }
    if (common.includes('常糖')) return '富铁软糖常糖版';
    if (common.includes('7mg')) return '富铁软糖7mg';
    return '富铁软糖';
  }
  if (name.includes('白芸豆阿拉伯糖') || common.includes('白芸豆')) return '白芸豆粉剂';
  if (name.includes('亢唐') || name.includes('抗糖') || common.includes('抗糖') || common.includes('亢糖')) return '抗糖';
  if (name.includes('白密码') || name.includes('白番茄') || common.includes('白番茄饮') || common.includes('美白饮')) return '白密码';
  if (name.includes('美白淡斑精华') || common.includes('美白精华') || common.includes('淡斑精华')) {
    if (common.includes('5只') || name.includes('5支装')) return '美白精华5支装';
    return '美白精华';
  }
  if (name.includes('洁面') || common.includes('洁面') || common.includes('洗面奶')) return '洁面';
  if (name.includes('修复') || common.includes('修复贴')) return '修复贴';
  if (name.includes('复合益生菌固体饮料3g*7') || (common.includes('益生菌') && !common.includes('5只') && !common.includes('冻干粉'))) return '益生菌';
  if (name.includes('复合益生菌-5支') || common.includes('益生菌5只')) return '益生菌5支装';
  if (name.includes('益生菌冻干粉') || common.includes('益生菌冻干粉')) {
    if (common.includes('3袋') || name.includes('3g*3')) return '益生菌冻干粉3袋';
    return '益生菌冻干粉';
  }
  if (name.includes('红参铁') || name.includes('气血饮') || common.includes('元气饮') || common.includes('气血饮') || common.includes('富铁饮') || common.includes('红参铁')) return '元气饮';
  if (name.includes('褪黑素') || common.includes('褪黑素')) return '褪黑素';
  if (name.includes('钙维生素D3') || common.includes('女钙') || common.includes('成人钙') || common.includes('女性钙')) {
    if (common.includes('2只')) return '女钙2支装';
    return '女钙';
  }
  if (name.includes('双钙营养包') || common.includes('儿童钙')) {
    if (common.includes('单只')) return '儿童钙单支';
    return '儿童钙';
  }
  if (name.includes('小博士钙铁锌') || common.includes('高高钙')) {
    if (common.includes('2只')) return '高高钙2支装';
    return '高高钙';
  }
  if (name.includes('多维片') || common.includes('多维')) return '多维';
  if (name.includes('牛初乳') || common.includes('牛初乳')) return '牛初乳';
  if (name.includes('磷虾油') || common.includes('磷虾油')) return '磷虾油';
  if (name.includes('水润面膜') || common.includes('水润面膜')) return '水润面膜';
  if (name.includes('开瓶器') || common.includes('开瓶器')) {
    if (name.includes('30ml') || common.includes('30ml') || common.includes('紫色') || common.includes('经典肽开瓶器')) return '30ml开瓶器';
    if (name.includes('50ml') || common.includes('粉色') || common.includes('升级肽开瓶器')) return '50ml开瓶器';
    if (common.includes('大开瓶器')) return '大开瓶器';
    return '开瓶器';
  }
  if (name.includes('手提袋') || common.includes('手提袋') || common.includes('礼品袋')) return '手提袋';
  if (name.includes('书包') || common.includes('书包')) return '书包';
  if (name.includes('身高贴') || common.includes('身高贴')) return '身高贴';
  if (name.includes('梳子') || common.includes('梳子') || name.includes('按摩梳')) return '梳子';
  if (name.includes('行李箱') || common.includes('行李箱')) return '行李箱';
  if (name.includes('虾青素盲盒') || common.includes('虾青素盲盒')) return '虾青素盲盒';
  if (name.includes('小熊') || common.includes('小熊盲盒')) return '小熊盲盒';
  if (name.includes('相伴礼盒') || common.includes('相伴礼盒')) return '相伴礼盒';
  return '其他:' + name;
}

// =================== 备注解析 ===================
const PRODUCT_KEYWORDS = [
  ['50ml开瓶器', ['50ml开瓶器', '50毫升开瓶器', '粉色开瓶器', '升级肽开瓶器']],
  ['30ml开瓶器', ['30ml开瓶器', '30毫升开瓶器', '紫色开瓶器', '经典肽开瓶器']],
  ['大开瓶器', ['大开瓶器']],
  ['开瓶器', ['开瓶器']],
  ['升级肽3.0', ['升级肽3.0', '升级肽3', 'egcg3.0', 'egcg3', '升级3.0', '升级3']],
  ['升级肽2.0', ['升级肽2.0', '升级肽2', 'egcg2.0', 'egcg2']],
  ['升级肽1.0', ['升级肽1.0', '升级肽1', 'egcg1.0', 'egcg1']],
  ['升级肽', ['升级肽', 'egcg']],
  ['经典肽3.0', ['经典肽3.0', '经典肽3', '3.0经典肽', '3.0版本经典肽', '经典肽3.0版本']],
  ['经典肽2.0', ['经典肽2.0', '经典肽2']],
  ['经典肽1.0', ['经典肽1.0', '经典肽1']],
  ['经典肽', ['经典肽']],
  ['高端肽', ['高端肽']],
  ['PQQ饮', ['pqq']],
  ['虾青素Pro', ['虾青素pro', '虾青素por', '虾青素pro版', '虾青pro', '虾青素 pro', '虾青素pro']],
  ['虾青素饮', ['虾青素饮']],
  ['虾青素6粒', ['6粒装虾青素', '6粒虾青素', '虾青素6粒', '6粒装虾青', '6粒虾青']],
  ['虾青素盲盒', ['虾青素盲盒']],
  ['虾青素', ['虾青素', '虾青']],
  ['富铁软糖降糖版6粒', ['6粒装富铁软糖', '6粒富铁软糖', '6粒装软糖', '6粒软糖', '富铁软糖6粒', '6粒富铁']],
  ['富铁软糖降糖版', ['减糖版富铁软糖', '降糖版富铁软糖', '减甜版富铁软糖', '减糖版软糖', '降糖版软糖', '减甜版软糖', '补铁软糖减甜版', '补铁软糖减糖版']],
  ['富铁软糖常糖版', ['常糖版富铁软糖', '常糖版软糖']],
  ['富铁软糖7mg', ['7mg富铁软糖', '富铁软糖7mg', '升级版7mg红枣味富铁软糖', '升级版7mg富铁软糖', '7mg红枣味富铁软糖', '红枣味富铁软糖']],
  ['富铁软糖', ['富铁软糖', '补铁软糖', '软糖']],
  ['白芸豆粉剂', ['白芸豆粉剂', '白芸豆片', '白芸豆粉', '白芸豆']],
  ['抗糖', ['抗糖', '亢糖', '亢唐', '白芸豆抗糖粉剂', '抗糖粉剂']],
  ['白密码', ['白密码', '白番茄饮', '美白饮', '白番茄']],
  ['美白精华', ['美白精华', '淡斑精华', '美白淡斑精华', '精华']],
  ['洁面', ['洁面', '洗面奶', '氨基酸洁面']],
  ['修复贴', ['修复贴', '修复']],
  ['益生菌5支装', ['益生菌5支', '益生菌5只']],
  ['益生菌冻干粉', ['益生菌冻干粉']],
  ['益生菌', ['益生菌']],
  ['元气饮', ['元气饮', '富参饮', '红参饮', '红参气血饮', '富铁饮', '红参铁']],
  ['褪黑素', ['褪黑素']],
  ['女钙', ['液体钙', '女钙', '成人钙', '女性钙']],
  ['儿童钙', ['儿童钙']],
  ['高高钙', ['高高钙']],
  ['多维', ['多维']],
  ['牛初乳', ['牛初乳']],
  ['磷虾油', ['磷虾油']],
  ['水润面膜', ['水润面膜', '面膜']],
  ['手提袋', ['手提袋', '礼品袋']],
  ['书包', ['书包']],
  ['身高贴', ['身高贴']],
  ['梳子', ['梳子', '按摩梳']],
  ['行李箱', ['行李箱']],
  ['小熊盲盒', ['小熊盲盒', '小熊']],
  ['相伴礼盒', ['相伴礼盒']],
];

function findProductInText(text) {
  const textLower = text.toLowerCase();
  for (const [category, keywords] of PRODUCT_KEYWORDS) {
    for (const kw of keywords) {
      const idx = textLower.indexOf(kw.toLowerCase());
      if (idx >= 0) return [category, idx, kw.length];
    }
  }
  if (textLower.includes('便携装') && !textLower.includes('软糖') && !textLower.includes('虾青素') && !textLower.includes('虾青')) {
    const idx = textLower.indexOf('便携装');
    return ['升级肽', idx, 3];
  }
  return [null, -1, 0];
}

function extractSpecFromText(text) {
  let spec = '';
  let m = text.match(/(\d+)\s*(?:ml|毫升)/i);
  if (m) {
    spec = m[1] + 'ml';
    text = text.replace(/\d+\s*(?:ml|毫升)/i, '').trim();
  }
  m = text.match(/(\d+)\s*粒(?:装)?/);
  if (m) {
    spec = m[1] + '粒';
    text = text.replace(/\d+\s*粒(?:装)?/, '').trim();
  }
  if (!spec) {
    m = text.match(/(一|二|两|三|四|五|六|七|八|九|十)\s*粒(?:装)?/);
    if (m) {
      const cn = parseCnNumber(m[1]);
      if (cn) {
        spec = cn + '粒';
        text = text.replace(/(一|二|两|三|四|五|六|七|八|九|十)\s*粒(?:装)?/, '').trim();
      }
    }
  }
  // 1盒3袋升级肽3.0：提取内装袋数作为规格（仅当前面有“盒”时）
  if (!spec) {
    m = text.match(/盒\s*(\d+)\s*袋(?:装)?/);
    if (m) {
      spec = m[1] + '袋';
      text = text.replace(/盒\s*\d+\s*袋(?:装)?/, '盒').trim();
    }
  }
  if (!spec) {
    m = text.match(/(\d+)\s*袋装/);
    if (m) {
      spec = m[1] + '袋';
      text = text.replace(/\d+\s*袋装/, '').trim();
    }
  }
  if (!spec) {
    m = text.match(/(一|二|两|三|四|五|六|七|八|九|十)\s*袋装/);
    if (m) {
      const cn = parseCnNumber(m[1]);
      if (cn) {
        spec = cn + '袋';
        text = text.replace(/(一|二|两|三|四|五|六|七|八|九|十)\s*袋装/, '').trim();
      }
    }
  }
  // 7mg规格（升级版7mg红枣味富铁软糖）
  if (!spec) {
    m = text.match(/(\d+)\s*mg/i);
    if (m) {
      spec = m[1] + 'mg';
      text = text.replace(/(\d+)\s*mg/i, '').trim();
    }
  }
  return [spec, text];
}

function extractQtyAndUnit(text) {
  let m = text.match(/(\d+)\s*(盒|瓶|个|袋|支|只|箱|份)/);
  if (m) {
    const qty = parseInt(m[1], 10);
    const unit = m[2];
    const remaining = text.slice(0, m.index) + text.slice(m.index + m[0].length);
    return [qty, unit, remaining];
  }
  m = text.match(/(一|二|两|三|四|五|六|七|八|九|十)\s*(盒|瓶|个|袋|支|只|箱|份)/);
  if (m) {
    const qty = parseCnNumber(m[1]);
    if (qty) {
      const unit = m[2];
      const remaining = text.slice(0, m.index) + text.slice(m.index + m[0].length);
      return [qty, unit, remaining];
    }
  }
  m = text.match(/[*×x]\s*(\d+)/);
  if (m) {
    const qty = parseInt(m[1], 10);
    const remaining = text.slice(0, m.index) + text.slice(m.index + m[0].length);
    return [qty, '', remaining];
  }
  const iter = text.matchAll(/(?<!\d)(\d{1,2})(?!\d)(?!\.\d)(?!\s*(?:ml|毫升|粒|mg|袋装))/gi);
  for (const mm of iter) {
    const num = parseInt(mm[1], 10);
    if (num > 0 && num < 100) {
      const remaining = text.slice(0, mm.index) + text.slice(mm.index + mm[0].length);
      return [num, '', remaining];
    }
  }
  return [null, '', text];
}

function adjustCategoryWithSpec(category, spec) {
  if (spec === '6粒') {
    if (category === '虾青素') return '虾青素6粒';
    if (category === '富铁软糖') return '富铁软糖降糖版6粒';
  }
  if (spec && spec.endsWith('mg') && category === '富铁软糖') return '富铁软糖7mg';
  return category;
}

function parseSegment(seg) {
  seg = seg.trim();
  if (!seg || seg.length < 1) return null;
  const originalSeg = seg;
  // X换Y / A+B换成C：只保留“换/换成”后面的要求
  const huanIdx = seg.indexOf('换成');
  const huanSimpleIdx = seg.indexOf('换');
  if (huanIdx >= 0) {
    seg = seg.slice(huanIdx + 2);
  } else if (huanSimpleIdx >= 0 && !seg.includes('换货')) {
    seg = seg.slice(huanSimpleIdx + 1);
  }
  seg = seg.trim();
  if (!seg) return null;

  seg = seg.replace(/^(备注[:：]?\s*|私域[-—:]?\s*)/, '');
  seg = seg.replace(/^(发货[：:]?|首发[：:]?|首单[：:]?|本次发出[：:]?|首次发货[：:]?|此次发出[：:]?|全部发[：:]?|发出[：:]?|实发[：:]?|补发[：:]?|发)\s*/, '');
  seg = seg.replace(/^(首发|首单|本次发出|首次发货|此次发出|全部发|发出|实发|补发)\s*/, '');
  seg = seg.replace(/^(加送|赠送|随单发|额外|补发|兑换|多送|加发|送|赠品发|礼包[：:]?)\s*/, '');
  seg = seg.replace(/^扣除积分\s*/, '');
  seg = seg.replace(/^\+\s*/, '');
  seg = seg.replace(/^\d{2,4}[.\-/]\d{1,2}[.\-/]\d{1,2}\s*/, '');
  seg = seg.replace(/^\d{6,8}\s*/, '');
  seg = seg.replace(/^\d+\.\d+号?\s*/, '');
  seg = seg.replace(/【[^】]*】/g, '');
  seg = seg.replace(/【[^】]*】/g, '');
  seg = seg.replace(/^(喝|的|版本)\s*/, '');
  seg = seg.trim();
  if (!seg) return null;

  // 过滤“最新生产日期X月份”“日期”等只剩产品说明的无数量行
  if (/^(最新生产日期|生产日期|发最新|发新鲜|日期)\d*/.test(seg)) return null;

  let [category, prodIdx, prodLen] = findProductInText(seg);
  if (category === null) return null;
  let [spec, segCleaned] = extractSpecFromText(seg);
  [category, prodIdx, prodLen] = findProductInText(segCleaned);
  const [qty, unit, segAfterQty] = extractQtyAndUnit(segCleaned);
  category = adjustCategoryWithSpec(category, spec);
  let finalQty = qty;
  if (finalQty === null) finalQty = 1;

  // 如果数量来自裸数字且 seg 里只有日期说明（如“升级肽3.0最新生产日期4月份的”），忽略
  if (unit === '' && spec === '' && segCleaned.match(/\d+\s*月(?:份)?(?:的)?/) && /(生产日期|最新日期|新鲜日期)/.test(segCleaned)) {
    return null;
  }

  const isExtra = /(加送|赠送|随单发|额外|补发|兑换|送|多送|加发|^\+\d|扣除积分|赠品发|礼包)/.test(originalSeg);

  // product_text 清理掉数量+单位，避免展示“9盒9盒升级肽”
  let displayText = segAfterQty.trim();
  displayText = displayText.replace(/^[装个颗的]+/, '').trim();
  if (!displayText) displayText = segCleaned.trim();

  return {
    product_text: displayText,
    category,
    quantity: finalQty,
    unit,
    spec,
    is_extra: isExtra,
  };
}

function trySplitBySpaces(seg) {
  let temp = seg;
  const positions = [];
  for (const [category, keywords] of PRODUCT_KEYWORDS) {
    const tempLower = temp.toLowerCase();
    for (const kw of keywords) {
      const idx = tempLower.indexOf(kw.toLowerCase());
      if (idx >= 0) {
        positions.push([idx, idx + kw.length]);
        break;
      }
    }
  }
  // 按位置去重：若短关键词完全落在长关键词范围内则忽略
  positions.sort((a, b) => a[0] - b[0] || b[1] - a[1]);
  const deduped = [];
  for (const p of positions) {
    if (!deduped.some(d => p[0] >= d[0] && p[1] <= d[1])) deduped.push(p);
  }
  if (deduped.length <= 1) return [seg];
  const parts = seg.split(/\s{2,}/);
  if (parts.length > 1) return parts;
  const result = [seg];
  const splitPoints = [];
  for (const m of seg.matchAll(/(\d+)\s*(盒|瓶|个|袋|支|只|箱|份)\s+/g)) {
    const after = seg.slice(m.index + m[0].length);
    const [catAfter] = findProductInText(after.slice(0, 20));
    if (!catAfter) continue;
    const before = seg.slice(0, m.index);
    const [catBefore] = findProductInText(before);
    if (!catBefore) continue;
    splitPoints.push(m.index + m[0].length);
  }
  if (splitPoints.length) {
    const out = [];
    let last = 0;
    for (const sp of splitPoints) {
      out.push(seg.slice(last, sp).trim());
      last = sp;
    }
    out.push(seg.slice(last).trim());
    return out.filter(Boolean);
  }
  const splitPoints2 = [];
  for (const m of seg.matchAll(/(\d+)(盒|瓶|个|袋|支|只|箱|份)?/g)) {
    if (m.index === 0) continue;
    const before = seg.slice(0, m.index);
    const [catBefore] = findProductInText(before);
    if (!catBefore) continue;
    const after = seg.slice(m.index + m[0].length);
    if (!after) continue;
    const [catAfter] = findProductInText(after.slice(0, 20));
    if (!catAfter) continue;
    const remaining = seg.slice(m.index + m[0].length, m.index + m[0].length + 2);
    if (remaining && '粒毫'.includes(remaining[0])) continue;
    if (m.index + m[0].length < seg.length && seg[m.index + m[0].length] === '.') continue;
    if (m[2] === '袋' && m.index + m[0].length < seg.length && seg[m.index + m[0].length] === '装') continue;
    splitPoints2.push(m.index);
  }
  if (splitPoints2.length) {
    const out = [];
    let last = 0;
    for (const sp of splitPoints2) {
      out.push(seg.slice(last, sp).trim());
      last = sp;
    }
    out.push(seg.slice(last).trim());
    return out.filter(Boolean);
  }
  return [seg];
}

const LOGISTICS_PATTERNS = [
  '发顺丰|发中通|发圆通|发韵达|发EMS|发邮政|发京东|发德邦|发申通|发极兔|发菜鸟',
  '不发圆通|不发中通|不发韵达|不发顺丰|不发货|不要发',
  '送货上门|送货到家|放门口|放驿站|放快递柜|本人签收',
  '[\u4e00-\u9fa5\w]{1,10}发(?:最新|新鲜)日期',
  '[\u4e00-\u9fa5\w]{1,10}发\d{1,2}月(?:份)?(?:的)?日期',
  '发最新日期|发新鲜日期|发\d{1,2}月(?:份)?(?:的)?日期|发\d{4}年\d{1,2}月',
  '拦截|拒收|退回|退款|退差价|补差价|差价',
  '直播间抽奖半价|直播间免单|抽奖半价|免单|中奖',
  '催发货|催审核|加急|尽快|辛苦|谢谢|拜托',
  '最新日期|新鲜日期',
];

function lineHasProductRequest(line) {
  const s = line.trim();
  if (!s) return false;
  if (/^(追加|补发|加发|加送|赠送|随单发|兑换)/.test(s)) return true;
  if (/\d+\s*(盒|瓶|袋|支|只|粒|个|罐|桶|套|盒装|瓶装|袋装|支装)/.test(s)) return true;
  for (const [category, keywords] of PRODUCT_KEYWORDS) {
    for (const kw of keywords) {
      if (s.includes(kw)) return true;
    }
  }
  return false;
}

function parseRemark(remark) {
  if (!remark || String(remark).trim() === '') return [];
  let text = String(remark).trim();
  for (const pattern of LOGISTICS_PATTERNS) {
    text = text.replace(new RegExp(pattern, 'gi'), ' ');
  }
  const cleanedForCheck = text.replace(/[a-zA-Z0-9_\-\s\|\n\r]/g, '');
  if (!cleanedForCheck || cleanedForCheck.length < 2) return [];

  let lines = text.split('\n');
  const productLines = lines.filter(lineHasProductRequest);
  if (productLines.length === 0) return [];
  text = productLines.join('\n');

  lines = text.split('\n');
  const dedupedLines = [];
  const seenBlocks = new Set();
  for (const line of lines) {
    const ls = line.trim();
    if (ls && !seenBlocks.has(ls)) {
      let isDup = false;
      for (const seen of seenBlocks) {
        if (ls.includes(seen) || seen.includes(ls)) { isDup = true; break; }
      }
      if (!isDup) { seenBlocks.add(ls); dedupedLines.push(line); }
    } else if (!ls) {
      dedupedLines.push(line);
    }
  }
  text = dedupedLines.join('\n');

  const giftPattern = /[（(]([^）)]+)[)）]/g;

  let textNoParens = text.replace(giftPattern, ' ');
  textNoParens = textNoParens.replace(/^\s*\+/, '加送');
  // 处理 X换Y / A+B换成C：每行只保留“换/换成”后面的要求
  const huanLines = [];
  for (const ln of textNoParens.split('\n')) {
    const trimLn = ln.trim();
    const huanIdx = trimLn.indexOf('换成');
    if (huanIdx >= 0) { huanLines.push(trimLn.slice(huanIdx + 2)); continue; }
    const simpleIdx = trimLn.indexOf('换');
    if (simpleIdx >= 0 && !trimLn.includes('换货')) { huanLines.push(trimLn.slice(simpleIdx + 1)); continue; }
    huanLines.push(trimLn);
  }
  textNoParens = huanLines.join('\n');
  let segments = textNoParens.split(/[+＋，,。、\n；;]\s*/);

  const mergedSegments = [];
  let i = 0;
  while (i < segments.length) {
    const segm = segments[i].trim();
    if (!segm) { i++; continue; }
    const [cat] = findProductInText(segm);
    if (!cat && segm.length <= 12 && !segm.includes('月') &&
        /\d+|[一二两三四五六七八九十]+\s*(?:ml|毫升|粒|盒|瓶|袋|支|只|个|箱)?/.test(segm)) {
      if (i + 1 < segments.length) {
        const nextSeg = segments[i + 1].trim();
        const [nextCat, nextIdx, nextLen] = findProductInText(nextSeg);
        if (nextCat) {
          const productKw = nextSeg.slice(nextIdx, nextIdx + nextLen);
          mergedSegments.push(segm + productKw);
          i += 2;
          continue;
        }
      }
    }
    mergedSegments.push(segm);
    i++;
  }
  segments = mergedSegments;

  const items = [];
  for (const segm of segments) {
    const s = segm.trim();
    if (!s) continue;
    if (/^(辛苦|谢谢|发顺丰|顺丰|最新日期|发最新|要求发|送货上门|其余|见小程序|婷婷代拍|永不退款|客户指定|优先安排|发2026|产品比较多|备注一定要|日期|订单编号|成交时间)/.test(s)) continue;
    if (/^(发)?顺丰/.test(s) && s.length <= 10) continue;
    const subSegs = trySplitBySpaces(s);
    for (const ss of subSegs) {
      const t = ss.trim();
      if (!t) continue;
      const item = parseSegment(t);
      if (item) items.push(item);
    }
  }

  // 解析括号内容作为可能的赠品/说明；若是已有货品的规格说明则忽略
  const giftItems = [];
  const normalCategories = new Set(items.map(it => it.category));
  for (const m of text.matchAll(giftPattern)) {
    const content = m[1].trim();
    let subSegments;
    if (content.includes('+') || content.includes('＋')) {
      subSegments = content.split(/[+＋]\s*/);
    } else {
      const merged = content.replace(/，/g, '').replace(/,/g, '');
      subSegments = merged.split(/\s+/);
    }
    for (let segm of subSegments) {
      segm = segm.trim();
      if (!segm || segm.length < 2) continue;
      const item = parseSegment(segm);
      if (item) {
        // 括号内容若是已有货品或其细分规格（同前缀族），忽略不重复计数
        const alreadyHas = [...normalCategories].some(nc => item.category === nc || item.category.startsWith(nc) || nc.startsWith(item.category));
        if (alreadyHas) continue;
        const descOnly = /^(升级版|红枣味|荔枝味|降糖版|减糖版|减甜版|常糖版|入门款|\d+mg|\d+粒|\d+袋|\d+盒|\d+瓶|\d+支)+$/.test(segm.replace(/[（）()]/g, '').trim());
        if (descOnly) continue;
        giftItems.push(item);
      }
    }
  }
  items.push(...giftItems);

  const originalText = String(remark).trim();
  if (/^(赠品发|赠品[：:]?|随单发)/.test(originalText)) {
    for (const item of items) item.is_extra = true;
  }

  return items;
}

// =================== 分类匹配 ===================
function categoriesMatch(remarkCat, actualCat) {
  if (!remarkCat || !actualCat) return false;
  if (remarkCat === actualCat) return true;
  if (remarkCat === '升级肽' && actualCat.startsWith('升级肽')) return true;
  if (remarkCat === '升级肽2.0' && (actualCat === '升级肽2.0' || actualCat === '升级肽2.0枇杷味')) return true;
  if (remarkCat === '经典肽' && actualCat.startsWith('经典肽')) return true;
  if (remarkCat === '虾青素' && actualCat === '虾青素1.0') return true;
  if (remarkCat === '富铁软糖' && ['富铁软糖常糖版', '富铁软糖降糖版', '富铁软糖7mg', '富铁软糖'].includes(actualCat)) return true;
  if (remarkCat === '美白精华' && actualCat.startsWith('美白精华')) return true;
  if (remarkCat === '益生菌' && ['益生菌', '益生菌5支装', '益生菌冻干粉', '益生菌冻干粉3袋'].includes(actualCat)) return true;
  if (remarkCat === '开瓶器' && ['30ml开瓶器', '50ml开瓶器', '大开瓶器', '开瓶器'].includes(actualCat)) return true;
  if (remarkCat === '钙' && (actualCat.startsWith('女钙') || actualCat.startsWith('儿童钙') || actualCat.startsWith('高高钙'))) return true;
  return false;
}

function getMatchGroup(category, packageType) {
  if (category.startsWith('升级肽')) {
    return packageType === '袋装' ? '升级肽袋装' : '升级肽瓶装';
  }
  if (category.startsWith('经典肽')) return '经典肽';
  if (category.startsWith('虾青素') && !category.includes('Pro') && !category.includes('6粒') && !category.includes('饮') && !category.includes('盲盒')) return '虾青素';
  if (category.startsWith('富铁软糖')) return '富铁软糖';
  if (category.startsWith('美白精华')) return '美白精华';
  if (category.startsWith('益生菌')) return '益生菌';
  if (['30ml开瓶器', '50ml开瓶器', '大开瓶器', '开瓶器'].includes(category)) return '开瓶器';
  if (category.startsWith('女钙') || category.startsWith('儿童钙') || category.startsWith('高高钙')) return '钙';
  return category;
}

function getRemarkPackageType(unit, spec, productText) {
  if (unit === '袋') return '袋装';
  if (spec && spec.includes('袋')) return '袋装';
  if (productText && productText.includes('袋装')) return '袋装';
  if (['盒', '瓶', '箱'].includes(unit)) return '瓶装';
  return '瓶装';
}

// =================== 列匹配 (sheet0 → sheet1) ===================
function findColIndex(headers, possibleNames, afterCol = -1, findLast = false) {
  const start = afterCol + 1;
  let result = -1;
  for (let i = start; i < headers.length; i++) {
    const h = headers[i];
    const hStr = h ? String(h).trim() : '';
    if (!hStr) continue;
    for (const name of possibleNames) {
      if (hStr.includes(name)) {
        if (findLast) { result = i; }
        else return i;
      }
    }
  }
  return result;
}

/*
 * sheet0Rows / sheet1Rows: 二维数组，第一行是表头
 * 返回匹配后的行列表（与 Python match_columns_from_workbook 输出结构一致）
 */
function matchColumnsFromWorkbook(sheet0Rows, sheet1Rows) {
  const ws0Headers = (sheet0Rows[0] || []).map(h => h ? String(h).trim() : '');
  const orderIdCol = findColIndex(ws0Headers, ['订单编号', '原始订单编号', '平台订单号', '主订单编号']);
  const webOrderCol = findColIndex(ws0Headers, ['网店订单号', '平台订单号', '线上订单号', '外部订单号']);
  const remarkCol = findColIndex(ws0Headers, ['合并备注', '备注']);

  const ws1Headers = (sheet1Rows[0] || []).map(h => h ? String(h).trim() : '');
  const s1OrderCol = findColIndex(ws1Headers, ['订单编号', '原始订单编号', '平台订单号', '主订单编号']);
  const s1WebOrderCol = findColIndex(ws1Headers, ['网店订单号', '平台订单号', '线上订单号', '外部订单号']);
  const barcodeCol = findColIndex(ws1Headers, ['条码']);
  const codeCol = findColIndex(ws1Headers, ['货品编号']);
  const nameCol = findColIndex(ws1Headers, ['货品名称']);
  const tagCol = findColIndex(ws1Headers, ['货品标记']);
  const specCol = findColIndex(ws1Headers, ['规格']);
  const qtyCol = findColIndex(ws1Headers, ['数量']);
  const priceCol = findColIndex(ws1Headers, ['单价']);
  const discountAmtCol = findColIndex(ws1Headers, ['优惠']);
  const discountRateCol = findColIndex(ws1Headers, ['折扣']);
  const amountCol = findColIndex(ws1Headers, ['金额']);
  const lineRemarkCol = findColIndex(ws1Headers, ['备注'], -1, true);
  const talentIdCol = findColIndex(ws1Headers, ['达人ID']);
  const talentNameCol = findColIndex(ws1Headers, ['达人名称']);
  const unitCol = findColIndex(ws1Headers, ['单位']);
  const giftCol = findColIndex(ws1Headers, ['赠品']);

  const orderMap = {};
  for (let r = 1; r < sheet0Rows.length; r++) {
    const row = sheet0Rows[r];
    let orderId = orderIdCol >= 0 ? toStr(safeGet(row, orderIdCol)).trim() : '';
    const webOrder = webOrderCol >= 0 ? toStr(safeGet(row, webOrderCol)).trim() : '';
    const remark = remarkCol >= 0 ? toStr(safeGet(row, remarkCol)).trim() : '';
    if (orderId && orderId !== 'None' && !(orderId in orderMap)) {
      orderMap[orderId] = [
        webOrder !== 'None' ? webOrder : '',
        remark !== 'None' ? remark : '',
      ];
    }
  }

  const matchedRows = [];
  for (let r = 1; r < sheet1Rows.length; r++) {
    const row = sheet1Rows[r];
    let orderId = s1OrderCol >= 0 ? toStr(safeGet(row, s1OrderCol)).trim() : '';
    if (orderId === 'None') orderId = '';
    let webOrder, mergedRemark;
    if (orderId && orderId in orderMap) {
      [webOrder, mergedRemark] = orderMap[orderId];
    } else {
      webOrder = ''; mergedRemark = '';
    }
    if (!webOrder && s1WebOrderCol >= 0) {
      const s1Web = toStr(safeGet(row, s1WebOrderCol)).trim();
      if (s1Web && s1Web !== 'None') webOrder = s1Web;
    }
    const qtyVal = qtyCol >= 0 ? safeGet(row, qtyCol) : '';
    let quantity = 0;
    try {
      if (qtyVal !== '' && qtyVal !== null && qtyVal !== 'None') quantity = parseInt(parseFloat(qtyVal), 10);
    } catch (e) { quantity = 0; }

    const clean = (val) => {
      const v = val !== null && val !== undefined ? String(val).trim() : '';
      return v === 'None' ? '' : v;
    };

    matchedRows.push({
      order_id: orderId,
      web_order_id: webOrder,
      remark: mergedRemark,
      barcode: barcodeCol >= 0 ? clean(safeGet(row, barcodeCol)) : '',
      product_code: codeCol >= 0 ? clean(safeGet(row, codeCol)) : '',
      product_name: nameCol >= 0 ? clean(safeGet(row, nameCol)) : '',
      product_tag: tagCol >= 0 ? clean(safeGet(row, tagCol)) : '',
      quantity,
      unit_price: priceCol >= 0 ? safeGet(row, priceCol, null) : null,
      amount: amountCol >= 0 ? safeGet(row, amountCol, null) : null,
      spec: specCol >= 0 ? clean(safeGet(row, specCol)) : '',
      talent_id: talentIdCol >= 0 ? clean(safeGet(row, talentIdCol)) : '',
      talent_name: talentNameCol >= 0 ? clean(safeGet(row, talentNameCol)) : '',
      line_remark: lineRemarkCol >= 0 ? clean(safeGet(row, lineRemarkCol)) : '',
      unit: unitCol >= 0 ? clean(safeGet(row, unitCol)) : '',
      discount_amount: discountAmtCol >= 0 ? safeGet(row, discountAmtCol, null) : null,
      discount_rate: discountRateCol >= 0 ? safeGet(row, discountRateCol, null) : null,
      is_gift: giftCol >= 0 ? clean(safeGet(row, giftCol)) : '',
    });
  }
  return matchedRows;
}

// =================== 核对主逻辑 ===================
function verifyOrders(matchedRows, barcodeToInfo, codeToInfo) {
  if (!codeToInfo) codeToInfo = {};
  const orders = {};
  for (const row of matchedRows) {
    const webOrderId = row.web_order_id;
    const orderId = row.order_id || '';
    const groupKey = webOrderId ? webOrderId : orderId;
    if (!groupKey) continue;
    const barcode = row.barcode;
    const productCode = row.product_code || '';
    if (!(groupKey in orders)) {
      orders[groupKey] = { remark: row.remark, items: [], web_order_id: webOrderId, order_id_fallback: !webOrderId };
    }
    let itemInfo = barcodeToInfo[barcode];
    if (!itemInfo && productCode) itemInfo = codeToInfo[productCode] || {};
    if (!itemInfo) itemInfo = {};
    const category = getProductCategory(barcode, barcodeToInfo, productCode, codeToInfo);
    orders[groupKey].items.push({
      order_id: orderId,
      barcode,
      product_code: productCode,
      product_name: row.product_name,
      quantity: row.quantity,
      is_gift: row.is_gift,
      category,
      box_count: itemInfo.box_count !== undefined ? itemInfo.box_count : 1,
      box_count_from_suffix: itemInfo.box_count_from_suffix || false,
      inner_count: itemInfo.inner_count !== undefined ? itemInfo.inner_count : 1,
      package_type: itemInfo.package_type || '瓶装',
      bags_per_pack: itemInfo.bags_per_pack !== undefined ? itemInfo.bags_per_pack : 1,
    });
  }

  const results = [];
  for (const groupKey of Object.keys(orders)) {
    const orderData = orders[groupKey];
    const webOrderId = orderData.web_order_id;
    const remark = orderData.remark;
    const items = orderData.items;
    const remarkItems = parseRemark(remark);

    const actualByGroup = {};
    for (const item of items) {
      const cat = item.category;
      const pkgType = item.package_type;
      const group = getMatchGroup(cat, pkgType);
      if (!actualByGroup[group]) {
        actualByGroup[group] = { box_qty: 0, piece_qty: 0, has_suffix_box_count: false, details: [], is_all_gift: true, cats: [] };
      }
      let pieceQty, detailSuffix;
      if (pkgType === '袋装') {
        pieceQty = item.quantity * item.bags_per_pack;
        detailSuffix = '袋×' + item.bags_per_pack;
      } else {
        pieceQty = item.quantity * item.inner_count;
        detailSuffix = '内装' + item.inner_count;
      }
      actualByGroup[group].box_qty += item.quantity;
      actualByGroup[group].piece_qty += pieceQty;
      if (item.box_count_from_suffix) actualByGroup[group].has_suffix_box_count = true;
      actualByGroup[group].details.push(item.product_name + '*' + item.quantity + '(' + detailSuffix + ')');
      actualByGroup[group].cats.push(cat);
      if (item.is_gift !== '是') actualByGroup[group].is_all_gift = false;
    }

    const remarkByGroup = {};
    for (const ri of remarkItems) {
      const pkgType = getRemarkPackageType(ri.unit, ri.spec, ri.product_text || '');
      const group = getMatchGroup(ri.category, pkgType);
      if (!remarkByGroup[group]) {
        remarkByGroup[group] = { box_qty: 0, piece_qty: 0, items: [], categories: [] };
      }
      let qty = ri.quantity;
      // “1盒3袋”：spec 里的袋数是内装件数，要换算成 piece_qty
      if (ri.spec && ri.spec.includes('袋')) {
        const m = ri.spec.match(/(\d+)袋/);
        if (m) qty = qty * parseInt(m[1], 10);
      }
      if (pkgType === '袋装') {
        remarkByGroup[group].piece_qty += qty;
      } else if (['盒', '箱'].includes(ri.unit)) {
        remarkByGroup[group].box_qty += qty;
      } else if (ri.unit === '' && !ri.spec) {
        // 裸数字（如“2益生菌”）优先按盒数/外包装数核对
        remarkByGroup[group].box_qty += qty;
      } else {
        remarkByGroup[group].piece_qty += qty;
      }
      remarkByGroup[group].items.push(ri);
      if (!remarkByGroup[group].categories.includes(ri.category)) remarkByGroup[group].categories.push(ri.category);
    }

    const verificationDetails = [];
    let isCorrect = true;
    const issues = [];
    const matchedGroups = new Set();

    for (const group of Object.keys(remarkByGroup)) {
      const remarkData = remarkByGroup[group];
      if (group === '未知') {
        for (const ri of remarkData.items) {
          verificationDetails.push({
            remark_item: ri.quantity + ri.unit + ri.product_text,
            actual_matched: '',
            actual_qty: '',
            remark_qty: ri.quantity,
            status: '⚠️备注中有无法识别的货品',
          });
        }
        continue;
      }
      const actualData = actualByGroup[group];
      if (actualData) {
        matchedGroups.add(group);
        const remarkDesc = remarkData.items.map(ri => ri.quantity + ri.unit + ri.product_text).join(' + ');
        const allExtra = remarkData.items.every(ri => ri.is_extra);
        const hasSuffix = actualData.has_suffix_box_count;
        let remarkPieceQty, remarkBoxQty;
        if (hasSuffix) {
          remarkPieceQty = remarkData.piece_qty + remarkData.box_qty;
          remarkBoxQty = 0;
        } else {
          remarkPieceQty = remarkData.piece_qty;
          remarkBoxQty = remarkData.box_qty;
        }
        const boxMatch = actualData.box_qty === remarkBoxQty;
        const pieceMatch = actualData.piece_qty === remarkPieceQty;
        const compareParts = [];
        if (remarkBoxQty > 0) compareParts.push('盒数:实际' + actualData.box_qty + 'vs备注' + remarkBoxQty);
        if (remarkPieceQty > 0) compareParts.push('件数:实际' + actualData.piece_qty + 'vs备注' + remarkPieceQty);
        const compareDesc = compareParts.join(' | ');
        let isMatch;
        if (remarkBoxQty > 0 && remarkPieceQty > 0) isMatch = boxMatch && pieceMatch;
        else if (remarkBoxQty > 0) isMatch = boxMatch;
        else isMatch = pieceMatch;
        if (isMatch) {
          verificationDetails.push({
            remark_item: remarkDesc,
            actual_matched: actualData.details.join('; '),
            actual_qty: '盒' + actualData.box_qty + '/件' + actualData.piece_qty,
            remark_qty: '盒' + remarkBoxQty + '/件' + remarkPieceQty,
            status: '✅正确',
          });
        } else {
          if (allExtra) {
            let extraOk = true;
            if (remarkBoxQty > 0 && actualData.box_qty < remarkBoxQty) extraOk = false;
            if (remarkPieceQty > 0 && actualData.piece_qty < remarkPieceQty) extraOk = false;
            if (extraOk) {
              verificationDetails.push({
                remark_item: remarkDesc,
                actual_matched: actualData.details.join('; '),
                actual_qty: '盒' + actualData.box_qty + '/件' + actualData.piece_qty,
                remark_qty: '盒' + remarkBoxQty + '/件' + remarkPieceQty,
                status: '✅匹配(加送/赠送项, ' + compareDesc + ')',
              });
              continue;
            }
          }
          isCorrect = false;
          issues.push('备注要求' + remarkDesc + '，' + compareDesc);
          verificationDetails.push({
            remark_item: remarkDesc,
            actual_matched: actualData.details.join('; '),
            actual_qty: '盒' + actualData.box_qty + '/件' + actualData.piece_qty,
            remark_qty: '盒' + remarkBoxQty + '/件' + remarkPieceQty,
            status: '❌数量不符(' + compareDesc + ')',
          });
        }
      } else {
        isCorrect = false;
        const remarkDesc = remarkData.items.map(ri => ri.quantity + ri.unit + ri.product_text).join(' + ');
        issues.push('备注要求' + remarkDesc + '但订单中未找到');
        verificationDetails.push({
          remark_item: remarkDesc,
          actual_matched: '未找到',
          actual_qty: 0,
          remark_qty: '盒' + remarkData.box_qty + '/件' + remarkData.piece_qty,
          status: '❌缺失',
        });
      }
    }

    const allRemarkExtra = remarkItems.length > 0 && remarkItems.every(ri => ri.is_extra);
    if (remarkItems.length > 0) {
      for (const group of Object.keys(actualByGroup)) {
        if (!matchedGroups.has(group)) {
          const data = actualByGroup[group];
          if (data.is_all_gift) {
            verificationDetails.push({
              remark_item: '(系统赠品)',
              actual_matched: data.details.join('; '),
              actual_qty: '盒' + data.box_qty + '/件' + data.piece_qty,
              remark_qty: 0,
              status: '🎁系统赠品',
            });
          } else if (allRemarkExtra) {
            verificationDetails.push({
              remark_item: '(备注仅含赠品要求)',
              actual_matched: data.details.join('; '),
              actual_qty: '盒' + data.box_qty + '/件' + data.piece_qty,
              remark_qty: 0,
              status: 'ℹ️正常购买(备注仅含赠品要求)',
            });
          } else {
            issues.push('订单中有' + data.cats.join(',') + '(盒' + data.box_qty + '/件' + data.piece_qty + ')但备注中未提及');
            verificationDetails.push({
              remark_item: '(备注未提及)',
              actual_matched: data.details.join('; '),
              actual_qty: '盒' + data.box_qty + '/件' + data.piece_qty,
              remark_qty: 0,
              status: '⚠️多发货品(备注未提及)',
            });
          }
        }
      }
    }

    if (remarkItems.length === 0 && issues.length === 0) {
      verificationDetails.push({
        remark_item: '(无具体货品要求)',
        actual_matched: '',
        actual_qty: '',
        remark_qty: '',
        status: 'ℹ️备注无具体货品要求',
      });
    }

    const orderIds = Array.from(new Set(items.map(it => it.order_id).filter(Boolean))).sort().join(', ');
    const displayWebOrderId = webOrderId ? webOrderId : '(未匹配到网店订单号)';
    results.push({
      web_order_id: displayWebOrderId,
      order_ids: orderIds,
      remark: remark,
      is_correct: isCorrect,
      issues: issues.length ? issues.join('; ') : '',
      details: verificationDetails,
      items,
    });
  }
  return results;
}

// =================== 导出数据 (供 XLSX 写出) ===================
const SKIPPED_REMARK_ITEMS = ['(系统赠品)', '(备注未提及)', '(无具体货品要求)', '(备注仅含赠品要求)'];

function buildExportData(results) {
  const headers = ['网店订单号', '订单编号', '核对结果', '问题说明', '备注内容', '备注要求货品及数量', '实际发货货品及数量', '明细核对'];
  const rows = [];
  for (const r of results) {
    const remarkItemsStr = r.details
      .filter(d => d.remark_item && !SKIPPED_REMARK_ITEMS.includes(d.remark_item))
      .map(d => d.remark_item).join('\n');
    const actualItemsStr = r.details.filter(d => d.actual_matched).map(d => d.actual_matched).join('\n');
    const detailStr = r.details.map(d => '[' + d.status + '] ' + d.remark_item).join('\n');

    let resultText, resultFill;
    const hasWarnings = !!r.issues && r.is_correct;
    if (r.is_correct && !hasWarnings) { resultText = '✅正确'; resultFill = 'correct'; }
    else if (hasWarnings) { resultText = '⚠️有警告'; resultFill = 'warning'; }
    else { resultText = '❌有误'; resultFill = 'error'; }

    rows.push({
      cells: [r.web_order_id, r.order_ids, resultText, r.issues, r.remark, remarkItemsStr, actualItemsStr, detailStr],
      fill: resultFill,
      issueFill: (r.issues ? (hasWarnings ? 'warning' : 'error') : null),
    });
  }

  const total = results.length;
  const correct = results.filter(r => r.is_correct && !(r.issues && r.is_correct)).length;
  const warnings = results.filter(r => r.is_correct && r.issues).length;
  const errors = results.filter(r => !r.is_correct).length;

  return {
    sheetName: '核对结果',
    headers,
    rows,
    summaryName: '核对汇总',
    summary: [
      ['总订单数', total],
      ['核对正确', correct],
      ['有警告(多发货品)', warnings],
      ['核对有误', errors],
    ],
  };
}

// =================== 模块导出 ===================
const VERIFY_CORE = {
  CN_NUM_MAP, parseCnNumber, safeGet, toStr,
  buildSingleProductInfo, buildBarcodeToInfo, buildCodeToInfo, getProductCategory,
  PRODUCT_KEYWORDS, findProductInText, extractSpecFromText, extractQtyAndUnit, adjustCategoryWithSpec,
  parseSegment, trySplitBySpaces, parseRemark,
  categoriesMatch, getMatchGroup, getRemarkPackageType,
  findColIndex, matchColumnsFromWorkbook,
  verifyOrders, buildExportData,
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = VERIFY_CORE;
}
if (typeof window !== 'undefined') {
  window.VERIFY_CORE = VERIFY_CORE;
}
