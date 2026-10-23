/**
 * BP转型锦囊 — 财务计算库
 * 从 finance-portfolio 已有技能集提取的JS版财务计算逻辑
 * 
 * 来源文件对照:
 *   - formatters.py  → fmtMoney, fmtPct
 *   - tax_calculator.py → calcProfit, calcTax
 *   - financing_optimizer.py → calcROI, calcWACC
 *   - indicator_engine.py → evalThreshold
 *   - config.py → THRESHOLDS, TAX_RATES
 */

// ========== 格式化工具 (from modules/shared/formatters.py) ==========

function fmtMoney(v, unit) {
  var x = Number(v);
  if (!isFinite(x)) return '—';   // 非数值不伪装成 0，避免掩盖数据问题
  function group(n) {
    var s = n.toFixed(2);
    if (/e/i.test(s)) return s;   // ≥1e21 时 toFixed 退化为指数形式，不再插逗号
    return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  // 显式单位优先：调用方传入的是「元」，按指定单位换算后必须保留该单位后缀
  if (unit === '万') return '¥' + group(x / 1e4) + '万';
  if (unit === '亿') return '¥' + group(x / 1e8) + '亿';
  // 未指定单位时保持原有自动换算
  if (Math.abs(x) >= 1e8) return '¥' + group(x / 1e8) + '亿';
  if (Math.abs(x) >= 1e4) return '¥' + group(x / 1e4) + '万';
  return '¥' + x.toFixed(2);
}

function fmtPct(v) {
  return (v * 100).toFixed(1) + '%';
}

// 把外部可控文本（文件名/错误信息/用户输入）转义后再交给 innerHTML，防 DOM 注入
function escHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

// ========== 指标标签与单位（避免按英文键名猜单位） ==========
// 单位口径：% 为百分比；'' 为倍数（无单位）；'次' 为周转次数；'倍' 为保障倍数
const RATIO_META = {
  debtRatio:        { label: '资产负债率',   unit: '%',  dp: 1 },
  currentRatio:     { label: '流动比率',     unit: '',   dp: 2 },
  quickRatio:       { label: '速动比率',     unit: '',   dp: 2 },
  netMargin:        { label: '净利率',       unit: '%',  dp: 1 },
  grossMargin:      { label: '毛利率',       unit: '%',  dp: 1 },
  roe:              { label: 'ROE',          unit: '%',  dp: 1 },
  roa:              { label: 'ROA',          unit: '%',  dp: 1 },
  assetTurnover:    { label: '资产周转率',   unit: '次', dp: 2 },
  interestCoverage: { label: '利息保障倍数', unit: '倍', dp: 2 },
  equityRatio:      { label: '权益比率',     unit: '%',  dp: 1 }
};

function fmtRatio(key, v) {
  var m = RATIO_META[key];
  var n = Number(v);
  if (!isFinite(n)) return (m ? m.label : key) + ': —';
  if (!m) return key + ': ' + n.toFixed(2);
  return m.label + ': ' + n.toFixed(m.dp) + m.unit;
}

// 科目键名 → 中文标签（仅用于展示，不改变数值）
const ACCOUNT_META = {
  revenue: '营业收入', cost: '营业成本', totalAssets: '资产总计', totalLiabilities: '负债合计',
  currentAssets: '流动资产合计', currentLiabilities: '流动负债合计', profit: '净利润', cash: '货币资金',
  adminExpense: '管理费用', financeExpense: '财务费用', equity: '所有者权益合计'
};

function fmtAccount(key, v) {
  var n = Number(v);
  var label = ACCOUNT_META[key] || key;
  if (!isFinite(n)) return label + ': —';
  return label + ': ¥' + (n / 1e4).toFixed(1) + '万';
}

// ========== CSV 编码探测与解码（UTF-8 BOM / UTF-8 / GBK） ==========
// 实测：本构建的 XLSX.read 对 type:'array' + codepage 不生效（GBK 会解成 Latin-1），
// 因此先显式解码成字符串，再以 type:'string' 交给 SheetJS。
function detectCsvEncoding(bytes) {
  if (bytes && bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
    return { encoding: 'utf-8', offset: 3 };
  }
  // Excel「Unicode 文本」导出为 UTF-16 + BOM，必须先于 UTF-8 探测
  if (bytes && bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
    return { encoding: 'utf-16le', offset: 2 };
  }
  if (bytes && bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
    return { encoding: 'utf-16be', offset: 2 };
  }
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    return { encoding: 'utf-8', offset: 0 };
  } catch (e) {
    return { encoding: 'gbk', offset: 0 };
  }
}

function decodeCsvBytes(bytes) {
  var enc = detectCsvEncoding(bytes);
  var view = enc.offset ? bytes.subarray(enc.offset) : bytes;
  try {
    return new TextDecoder(enc.encoding).decode(view);
  } catch (e) {
    // 部分运行环境不支持 utf-16be：交换字节后按 utf-16le 解码
    if (enc.encoding === 'utf-16be') {
      var even = view.length - (view.length % 2);
      var swapped = new Uint8Array(even);
      for (var i = 0; i + 1 < even; i += 2) { swapped[i] = view[i + 1]; swapped[i + 1] = view[i]; }
      try { return new TextDecoder('utf-16le').decode(swapped); } catch (e2) { /* fallthrough */ }
    }
    // 运行环境不支持 gbk 时退回 UTF-8：可能乱码，但不抛错
    return new TextDecoder('utf-8').decode(view);
  }
}

// ========== 盈利计算 (from modules/tax_management/tax_calculator.py) ==========

function calcProfit(params) {
  const rev = params.revenue || 0;
  const vCost = params.variableCost || 0;
  const fCost = params.fixedCost || 0;
  const capCost = params.capitalCost || 0;
  const taxRate = params.taxRate || 25;

  const grossProfit = rev - vCost;
  const operatingProfit = grossProfit - fCost;
  const ebit = operatingProfit - capCost;
  const tax = Math.max(ebit * taxRate / 100, 0);
  const netProfit = ebit - tax;

  return {
    revenue: rev,
    variableCost: vCost,
    grossProfit: grossProfit,
    grossMargin: rev > 0 ? grossProfit / rev : 0,
    operatingProfit: operatingProfit,
    fixedCost: fCost,
    ebit: ebit,
    tax: tax,
    netProfit: netProfit,
    netMargin: rev > 0 ? netProfit / rev : 0,
    effectiveTaxRate: ebit > 0 ? tax / ebit : 0
  };
}

// ========== ROI 投资测算 (from modules/capital_management/financing_optimizer.py) ==========

function calcROI(params) {
  const invest = params.investment || 0;
  const annualReturn = params.annualReturn || 0;
  const years = params.years || 5;
  const discountRate = params.discountRate || 0.08;

  if (invest <= 0) return { roi: 0, payback: NaN, npv: 0, irr: 0, assessment: '— 请输入投资额' };

  // ROI = (总收益 - 投资) / 投资
  const totalReturn = annualReturn * years;
  const roi = (totalReturn - invest) / invest * 100;

  // 回收期 = 投资 / 年收益
  const payback = invest / annualReturn;

  // NPV = Σ(年收益/(1+r)^t) - 投资
  let npv = -invest;
  for (let t = 1; t <= years; t++) {
    npv += annualReturn / Math.pow(1 + discountRate, t);
  }

  return {
    investment: invest,
    annualReturn: annualReturn,
    years: years,
    totalReturn: totalReturn,
    roi: roi,
    payback: payback,
    npv: npv,
    isViable: npv > 0,
    assessment: npv > 0 ? '✅ 值得投资' : (roi > 0 ? '🟡 可回收但增值有限' : '🔴 不建议投资')
  };
}

// ========== WACC 资本成本 (from financing_optimizer.py) ==========

function calcWACC(params) {
  const { equityValue, debtValue, costOfEquity, costOfDebt, taxRate } = params;
  const total = equityValue + debtValue;
  if (total <= 0) return { wacc: 0 };

  const debtAfterTax = costOfDebt * (1 - taxRate / 100);
  const wacc = (equityValue / total) * costOfEquity + (debtValue / total) * debtAfterTax;
  return { wacc, total, equityRatio: equityValue / total, debtRatio: debtValue / total };
}

// ========== 阈值评估 (from modules/shared/thresholds.py) ==========

function evalThreshold(value, ranges, higherIsBetter) {
  for (const [lo, hi, level] of ranges) {
    if (higherIsBetter) {
      if (value >= lo) return level;
    } else {
      if (value <= hi) return level;
    }
  }
  return ranges[ranges.length - 1][2];
}

// ========== 行业基准数据 (from modules/shared/config.py + benchmarks.py) ==========

const BENCHMARKS = {
  '制造业': {
    grossMargin: 25, netMargin: 8, debtRatio: 50,
    currentRatio: 1.5, roe: 10, assetTurnover: 0.7
  },
  '科技业': {
    grossMargin: 40, netMargin: 12, debtRatio: 35,
    currentRatio: 2.0, roe: 12, assetTurnover: 0.55
  },
  '零售业': {
    grossMargin: 22, netMargin: 3, debtRatio: 60,
    currentRatio: 1.2, roe: 8, assetTurnover: 1.5
  },
  '建筑业': {
    grossMargin: 12, netMargin: 3, debtRatio: 70,
    currentRatio: 1.1, roe: 6, assetTurnover: 0.6
  }
};

// ========== AI 响应库 (思维重塑/战略/经营/成长四大模块) ==========

const AI_RESPONSES = {
  '毛利率': '毛利率是BP最核心的指标之一。建议按产品线-客户-区域三层拆解：1)产品线：对比各产品毛利率定位低毛利产品 2)客户：核算大客户vs散客毛利率差异 3)区域：各区域定价和成本差异。影响毛利率的6大因素：原材料价格、人工效率、产品结构、定价策略、规模效应、汇率波动。',
  '净利率': '净利率 = 净利润/营收。与毛利率的差值就是期间费用率。净利率提升路径：1)向上提毛利率(产品升级/定价优化) 2)向下控费用(固定费用降本/变动费用增效) 3)税务筹划(政策红利应享尽享)。行业基准参考：制造业8-12%、科技业12-20%、零售业2-5%。',
  '预算': '全面预算管理四步法：1)编制—增量/零基/滚动三种方法按场景选 2)执行—按月跟踪执行率 3)分析—差异归因(量差/价差/结构差) 4)调整—季度复盘+滚动更新。常见问题：预算松弛(虚报20-30%)、与战略脱节。推荐用AI工具：从历史数据训练+自动识别偏差动因+每月滚动。',
  '成本': '成本管控三维模型：1)变动成本—量价分析法(实际用量vs标准用量、实际价格vs标准价格) 2)固定费用—零基预算法(从零论证每笔支出必要性) 3)隐性成本—流程优化法(审批环节压缩、跨部门协同)。每月做一次成本对标：同比+环比+预算vs实际。',
  'ROI': 'ROI评估不仅要看数字，更要看质量。三步走：1)定量计算：投资额、年收益、年限、折现率→NPV/IRR/回收期 2)定性评估：战略匹配度、技术可行性、团队能力 3)敏感性分析：营收下浮10%/成本上浮10%/折现率变动对NPV的影响。IRR > WACC + 风险溢价才是好项目。',
  '现金流': '现金流是企业的血液。经营现金流分析三步法：1)流量—经营活动净额(正/负？同比趋势？) 2)质量—现金流/净利润比率(>1为佳) 3)预测—13周滚动现金流预测(周为单位，精确到客户回款和供应商付款)。现金流危机预警：现金余额<3个月支出、现金/短债<0.5、经营性现金流连续为负。',
  'BP': '财务BP转型SOP：第一阶段(1-3月)—跟业务开会，了解客户/产品/流程。第二阶段(3-6月)—参与业务决策，提出第一个改善建议。第三阶段(6-12月)—独立项目负责，用数据驱动结果。核心能力：业务理解力、数据分析力、沟通影响力、战略思维力。工具：Excel(基础) → BI(中级) → AI(高级)。',
  '预算预测': '滚动预测是BP的核心武器。与传统预算的区别：预算是一年一次静态的，预测是每月滚动动态的。实施方法：当月结后就更新未来12个月预测。AI在预测中的应用：1)自动识别异常 2)智能归因 3)情景模拟。从"数数字"到"做分析"的转变是关键。',
  '业财融合': '业财融合的四个抓手：1)销售—定价策略参与、客户盈利分析、销售费用ROI 2)采购—供应商评估模型、TCO总拥有成本分析、应付账款管理 3)生产—标准成本制定、差异分析、产能利用率 4)研发—项目投资评估、IP价值管理、研发费用资本化。关键：同一套数据、同一套语言、同一个目标。',
  '财务分析': '财务分析四层穿透：第一层—报表分析(三表勾稽、趋势、结构) 第二层—比率分析(盈利能力/偿债能力/运营效率/成长性) 第三层—拆解分析(杜邦分析、量价差、ABC成本法) 第四层—业务分析(客户盈利、产品盈利、渠道ROI)。最高境界：从第四层反哺第一层。',
  '杜邦分析': '杜邦分析拆解ROE：ROE=净利率×总资产周转率×权益乘数。三个维度：1)净利率(盈利能力) 2)周转率(运营效率) 3)乘数(杠杆水平)。典型的ROE改善路径：先提净利率→再提周转率→最后考虑杠杆。',
  '资产负债率': '资产负债率=总负债/总资产。行业差异大：制造业40-60%、公用事业50-70%、建筑业60-80%、科技业20-40%。判断标准：1)与行业对比(横向) 2)与历史对比(纵向) 3)与利息覆盖倍数结合(动态)。高负债率+低覆盖倍数=高风险信号。',
  '流动比率': '流动比率=流动资产/流动负债。一般行业>1.5为安全，>2为充裕。但不同行业差异大：零售业<1.0也正常(现金周转快)，制造业通常1.5-2.5。结合速动比率一起看更全面。低于1意味着流动资产不足覆盖流动负债，是流动性风险信号。',
  '应收账款': '应收账款管理五步法：1)事前—客户信用评估(新客户/老客户分级授信) 2)事中—账期管理(标准账期+预警机制) 3)事后—催收流程(逾期7天/30天/60天分级别处理) 4)复盘—账龄分析(月度) 5)改善—TOP5逾期客户专项处理。核心指标：应收账款周转天数=应收/日均营收。',
  '存货': '存货周转管理四步法：1)分类—ABC分类(A类严控/B类常规/C类宽松) 2)定标准—安全库存量+补货点 3)监控—周转天数趋势(周度) 4)行动—滞销品处理(打折/退货/报废)。警惕囤货癖——存货占用资金是隐性成本。',
  '投资决策': '投资决策评估矩阵：维度1—财务(NPV>0、IRR>WACC、回收期<行业平均) 维度2—战略(是否核心业务、技术壁垒、竞争地位) 维度3—风险(最坏情景现金流、政策风险、技术替代风险)。评分制：每个维度1-5分，总分>12分的项目优先。',
  '经营杠杆': '经营杠杆(DOL)=边际贡献/EBIT。DOL越高，营收波动对利润影响越大。高DOL行业：航空/钢铁/化工(固定成本占比高)。低DOL行业：咨询/服务(变动成本占比高)。管理启示：高DOL企业应优先追求营收增长而非成本控制(规模效应更显著)。',
  'default': '您的BP问题，我可以从以下维度给您建议：1)财务分析(三表/比率/杜邦) 2)预算预测(编制/滚动/差异) 3)成本管控(量价差/零基/ABC) 4)投资决策(ROI/NPV/IRR) 5)现金流管理(预测/优化/预警) 6)业财融合(业务理解/决策支持)。请描述具体场景。'
};


// ========== CVP 本量利分析 (from modules/budget_management/cost_accounting.py) ==========

function calcCVP(params) {
  const price = params.unitPrice || 0;       // 单价
  const unitVar = params.unitVariable || 0;   // 单位变动成本
  const fixedCost = params.fixedCost || 0;    // 固定成本总额
  const salesVolume = params.salesVolume || 0; // 预计销量
  const targetProfit = params.targetProfit || 0; // 目标利润

  const unitMargin = price - unitVar;          // 单位边际贡献
  const marginRate = price > 0 ? unitMargin / price : 0; // 边际贡献率
  const bepUnits = unitMargin > 0 ? Math.ceil(fixedCost / unitMargin) : 0;  // 保本量
  const bepRevenue = marginRate > 0 ? fixedCost / marginRate : 0;           // 保本额

  // 安全边际
  const actualRevenue = price * salesVolume;
  const marginOfSafety = actualRevenue - bepRevenue;
  const msr = actualRevenue > 0 ? marginOfSafety / actualRevenue * 100 : 0; // 安全边际率

  // 目标利润所需销量
  const targetUnits = unitMargin > 0 ? (fixedCost + targetProfit) / unitMargin : 0;

  // 经营杠杆
  const actualProfit = unitMargin * salesVolume - fixedCost;
  const dol = actualProfit > 0 ? (unitMargin * salesVolume) / actualProfit : 0;

  return {
    unitContributionMargin: unitMargin,
    contributionMarginRate: marginRate,
    bepUnits: bepUnits,
    bepRevenue: bepRevenue,
    safetyMargin: marginOfSafety,
    safetyMarginRate: msr,
    actualProfit: actualProfit,
    targetVolume: targetUnits,
    dol: dol,
    assessment: msr > 30 ? '安全' : (msr > 15 ? '较安全' : (msr > 0 ? '需关注' : '危险'))
  };
}

// ========== 现金流预测 (from modules/capital_management/cash_flow_tracker.py) ==========

function calcCashFlow(params) {
  const { initialBalance, monthlyInflows, monthlyOutflows, minBalance } = params;
  const projection = [];
  let balance = initialBalance;
  let minBal = initialBalance;
  let dangerDays = 0;

  for (let i = 0; i < monthlyInflows.length; i++) {
    balance = balance + (monthlyInflows[i] || 0) - (monthlyOutflows[i] || 0);
    projection.push(Math.round(balance));
    if (balance < minBal) minBal = balance;
    if (balance < (minBalance || 0)) dangerDays++;
  }

  return {
    projection: projection,
    finalBalance: Math.round(balance),
    minBalance: Math.round(minBal),
    dangerMonths: dangerDays,
    avgDailyOutflow: monthlyOutflows.length > 0 ? Math.round(monthlyOutflows.reduce((a,b)=>a+b,0) / monthlyOutflows.length / 30) : 0,
    daysCoverage: initialBalance > 0 && monthlyOutflows.length > 0
      ? (initialBalance / (monthlyOutflows.reduce((a,b)=>a+b,0) / monthlyOutflows.length)).toFixed(1)
      : 0,
    isSafe: dangerDays === 0
  };
}

// ========== 风险评估 (from modules/risk_monitoring/indicator_engine.py + risk_model.py) ==========

function calcRiskScore(indicators) {
  // indicators: { debtRatio, currentRatio, quickRatio, netMargin, roe, revenueGrowth }
  const scoreMap = [];
  if (indicators.debtRatio !== undefined) {
    const v = indicators.debtRatio;
    scoreMap.push({ name: '资产负债率', value: v, score: v < 50 ? 100 : v < 60 ? 80 : v < 70 ? 60 : v < 85 ? 30 : 10, weight: 0.2 });
  }
  if (indicators.currentRatio !== undefined) {
    const v = indicators.currentRatio;
    scoreMap.push({ name: '流动比率', value: v, score: v > 2.0 ? 100 : v > 1.5 ? 80 : v > 1.0 ? 60 : v > 0.8 ? 40 : 20, weight: 0.18 });
  }
  if (indicators.netMargin !== undefined) {
    const v = indicators.netMargin;
    scoreMap.push({ name: '净利率', value: v, score: v > 15 ? 100 : v > 10 ? 80 : v > 5 ? 60 : v > 0 ? 40 : 20, weight: 0.22 });
  }
  if (indicators.roe !== undefined) {
    const v = indicators.roe;
    scoreMap.push({ name: 'ROE', value: v, score: v > 20 ? 100 : v > 15 ? 80 : v > 10 ? 60 : v > 5 ? 40 : 20, weight: 0.2 });
  }
  if (indicators.revenueGrowth !== undefined) {
    const v = indicators.revenueGrowth;
    scoreMap.push({ name: '营收增长', value: v, score: v > 20 ? 100 : v > 10 ? 80 : v > 5 ? 60 : v > 0 ? 40 : 10, weight: 0.2 });
  }

  const totalWeight = scoreMap.reduce((s, i) => s + i.weight, 0);
  const totalScore = scoreMap.reduce((s, i) => s + i.score * i.weight, 0);
  const composite = totalWeight > 0 ? Math.round(totalScore / totalWeight) : 0;

  return {
    items: scoreMap,
    composite: composite,
    level: composite >= 85 ? '🟢 健康' : (composite >= 70 ? '🟡 关注' : (composite >= 50 ? '🟠 预警' : '🔴 危险')),
    riskFactors: scoreMap.filter(i => i.score < 60).map(i => i.name)
  };
}

// ========== 预算预测 (from modules/budget_management/budget_planner.py) ==========

function calcBudgetForecast(history, growthRate) {
  // history: [月度实际数据]
  // growthRate: 预期增长率
  if (!history || history.length < 2) return { error: '至少需要2个月数据' };

  const avg = history.reduce((a,b) => a+b, 0) / history.length;
  // 简单线性趋势
  let slope = 0;
  const n = history.length;
  const xMean = (n - 1) / 2;
  const yMean = avg;
  let num = 0, den = 0;
  for (let i = 0; i < n; i++) {
    num += (i - xMean) * (history[i] - yMean);
    den += (i - xMean) * (i - xMean);
  }
  if (den > 0) slope = num / den;

  // 预测未来3个月
  const forecast = [];
  for (let i = 0; i < 3; i++) {
    const trend = avg + slope * (n + i + 1 - xMean);
    const growth = trend * (1 + growthRate / 100);
    forecast.push(Math.round(growth));
  }

  // 差异分析（最近实际 vs 预测）
  const lastActual = history[history.length - 1];
  const predictedNext = forecast[0];
  const variance = predictedNext - lastActual;
  const varianceRate = lastActual > 0 ? variance / lastActual * 100 : 0;

  return {
    history: history,
    avgMonthly: Math.round(avg),
    trend: Math.round(slope * 100) / 100,
    forecast: forecast,
    variance: Math.round(variance),
    varianceRate: Math.round(varianceRate * 10) / 10,
    direction: slope > 0 ? '📈 上升' : '📉 下降',
    recommendation: varianceRate > 10 ? '预计增长，可适度增加资源投入'
      : varianceRate < -10 ? '预计下降，需控制成本'
      : '趋势平稳，维持当前节奏'
  };
}

// ========== 🏢 企业财务分析 (统一升级版) ==========

function calcFinancialRatios(bs) {
  const r = {};
  r.debtRatio = bs.totalAssets > 0 ? bs.totalLiabilities / bs.totalAssets * 100 : 0;
  r.currentRatio = bs.currentLiabilities > 0 ? bs.currentAssets / bs.currentLiabilities : 0;
  r.quickRatio = bs.currentLiabilities > 0 ? (bs.currentAssets - (bs.inventory || 0)) / bs.currentLiabilities : 0;
  r.netMargin = bs.revenue > 0 ? bs.profit / bs.revenue * 100 : 0;
  r.grossMargin = bs.revenue > 0 ? (bs.revenue - bs.cost) / bs.revenue * 100 : 0;
  r.roe = bs.equity > 0 ? bs.profit / bs.equity * 100 : 0;
  r.roa = bs.totalAssets > 0 ? bs.profit / bs.totalAssets * 100 : 0;
  r.assetTurnover = bs.totalAssets > 0 ? bs.revenue / bs.totalAssets : 0;
  r.interestCoverage = bs.interestExpense > 0 ? (bs.profit + bs.interestExpense) / bs.interestExpense : 0;
  r.equityRatio = bs.totalAssets > 0 ? bs.equity / bs.totalAssets * 100 : 0;
  return r;
}

function calcDuPont(bs) {
  const r = calcFinancialRatios(bs);
  const nm = r.netMargin / 100;
  const lev = bs.equity > 0 ? bs.totalAssets / bs.equity : 0;
  const roe = nm * r.assetTurnover * lev * 100;
  return { roe, nm: r.netMargin, turn: r.assetTurnover, lev, 
    detail: `${nm.toFixed(3)} × ${r.assetTurnover.toFixed(3)} × ${lev.toFixed(2)} = ${roe.toFixed(1)}%` };
}

function calcBenchmark(cv, industry) {
  const b = BENCHMARKS[industry] || BENCHMARKS['制造业'];
  const list = [];
  for (const {key,name,unit} of [{key:'grossMargin',name:'毛利率',unit:'%'},{key:'netMargin',name:'净利率',unit:'%'},
    {key:'debtRatio',name:'资产负债率',unit:'%'},{key:'currentRatio',name:'流动比率',unit:''},
    {key:'roe',name:'ROE',unit:'%'},{key:'assetTurnover',name:'资产周转率',unit:'次'}]) {
    const v = cv[key], bv = b[key];
    if (v===undefined||bv===undefined) continue;
    const hb = !['debtRatio'].includes(key);
    const diff = v - bv;
    const g = hb ? (diff>0?' 优于':diff>-5?' 持平':' 偏低') : (diff>0?' 偏高':diff>-5?' 持平':' 优于');
    list.push({name,company:v,bmk:bv,diff,grade:g,unit});
  }
  return {industry,list,warnCnt:list.filter(x=>x.grade.includes('偏低')||x.grade.includes('偏高')).length};
}

function validateSheets(d) {
  const f = [];
  if (d.pl&&d.detail) { const x = Math.abs((d.pl.revenue||0)-(d.detail.revenue||0)); f.push({id:'CK01',n:'收入核对',s:x<1?'✅':x<100?'🟡':'🔴',diff:x}); }
  if (d.pl&&d.tax) { f.push({id:'CK02',n:'税负率',s:(d.tax.total/(d.pl.revenue||1)*100)<5?'✅':'🟡',v:d.tax.total}); }
  return {checks:f,pass:f.filter(x=>x.s.includes('✅')).length/f.length*100||0};
}

function matchAccounts(sheets) {
  const r = {revenue:0,cost:0,totalAssets:0,totalLiabilities:0,currentAssets:0,currentLiabilities:0,profit:0,cash:0,adminExpense:0,financeExpense:0,equity:0};
  const kw = {revenue:['营业收入','营收'],cost:['营业成本'],totalAssets:['资产总计','总资产'],totalLiabilities:['负债合计','总负债'],
    currentAssets:['流动资产合计'],currentLiabilities:['流动负债合计'],profit:['净利润','利润总额'],cash:['货币资金'],
    adminExpense:['管理费用'],financeExpense:['财务费用'],equity:['所有者权益合计']};
  for (const sh of sheets) {
    for (const row of sh.data) {
      for (let i=0;i<row.length;i++) {
        const c = String(row[i]||'').replace(/\s/g,'');
        // 逐列取第一个可解析数值：0 是合法值，不得用 `||` 落到下一列
        let v = null;
        for (let j = i + 1; j <= i + 2; j++) {
          const n = parseFloat(row[j]);
          if (isFinite(n)) { v = n; break; }
        }
        if (v === null) continue; // 首个数值列胜出；0 是合法值，不落到下一列
        for (const [k,ws] of Object.entries(kw)) { if (ws.some(w=>c.includes(w))) r[k]=Math.max(r[k],v); }
      }
    }
  }
  return r;
}
// ========== 多表 / 多年份解析（2026-10-23 · v2） ==========
// 口径：报表类型 → 年份（标题日期优先）→ 表头列（跳过行次）→ 精确科目。
// 不跨 sheet 取 Math.max；不把行次当金额；结构化表表头不识别时**跳过并告警**，
// 不退回“第一个数字”的简单表解析（宁可不解析，也不静默取错）。

var ACC_SPEC = {
  revenue: ['营业收入', '营业总收入', '营收'],
  cost: ['营业成本', '营业总成本'],
  totalAssets: ['资产总计'],
  totalLiabilities: ['负债合计'],
  currentAssets: ['流动资产合计'],
  currentLiabilities: ['流动负债合计'],
  profitNet: ['净利润', '净利润（净亏损以"-"号填列）', '净利润(净亏损以"-"号填列)'],
  profitGross: ['利润总额', '利润总额（亏损总额以"-"号填列）', '利润总额(亏损总额以"-"号填列)'],
  cash: ['货币资金'],
  adminExpense: ['管理费用'],
  financeExpense: ['财务费用'],
  equity: ['所有者权益合计', '股东权益合计', '所有者权益（或股东权益）合计', '所有者权益(或股东权益)合计']
};

function _normLabel(v) {
  var s = String(v == null ? '' : v).replace(/\u3000/g, ' ');
  s = s.replace(/\s+/g, '');
  s = s.replace(/[“”]/g, '"');
  s = s.replace(/^[一二三四五六七八九十]+、/, '');
  s = s.replace(/^[（(][一二三四五六七八九十]+[）)]、?/, '');
  s = s.replace(/^(减|加|其中)[:：]/, '');
  s = s.replace(/[:：]+$/, '');
  return s;
}

function _parseNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  s = s.replace(/,/g, '').replace(/[¥￥]/g, '').replace(/\s/g, '');
  if (s === '—' || s === '-' || s === '--' || s === '——') return null;
  var neg = /^[（(].*[）)]$/.test(s);
  s = s.replace(/^[（(]/, '').replace(/[）)]$/, '');
  var n = parseFloat(s);
  if (!isFinite(n)) return null;
  return neg ? -n : n;
}

// 逐表识别金额单位。顺序要紧：万元/千元必须排在「元」之前，
// 否则「单位：人民币万元」会被当成元（少乘 10000 且不告警）。
// 覆盖：单位：万元 / 单位　：　万元（全角空格）/ 单位：人民币（万元）/ 单位：人民币／万元 /
//       表名或标题里的（万元）/ 表头「期末余额（万元）」；explicit=false 表示按元计且未标注。
function _detectUnit(rows, sheetName) {
  var t = rows.slice(0, 12).map(function (row) {
    return (row || []).map(function (x) { return String(x == null ? '' : x); }).join(' ');
  }).join(' ') + ' ' + String(sheetName == null ? '' : sheetName);
  // ① 显式「单位：…」：只认**第一个**声明；后续声明与任何括号标注都不得覆盖它
  var ex = t.match(/单位\s*[:：]?\s*(?:人民币)?\s*[（(\\[\/／·、,，\-]*\s*(万元|千元|亿元|百万元|元)/);
  if (ex) {
    var eu = ex[1];
    if (eu === '万元') return { unit: '万元', factor: 10000, explicit: true };
    if (eu === '千元') return { unit: '千元', factor: 1000, explicit: true };
    if (eu === '元') return { unit: '元', factor: 1, explicit: true };
    return { unit: '元', factor: 1, explicit: false, unsupported: eu };
  }
  // ② 没有显式声明时，才看「（万元）」这类括号标注；按出现位置取第一个，多值即标冲突
  var re = /[（(]\s*(?:人民币)?\s*(万元|千元|亿元|百万元|元)\s*[）)]/g;
  var found = [], mm;
  while ((mm = re.exec(t)) !== null) {
    found.push(mm[1]);
    if (found.length >= 20) break;
  }
  if (found.length) {
    var distinct = {};
    for (var fi = 0; fi < found.length; fi++) distinct[found[fi]] = 1;
    var conflict = Object.keys(distinct).length > 1;
    var fu = found[0];
    if (fu === '万元') return { unit: '万元', factor: 10000, explicit: true, conflict: conflict };
    if (fu === '千元') return { unit: '千元', factor: 1000, explicit: true, conflict: conflict };
    if (fu === '元') return { unit: '元', factor: 1, explicit: true, conflict: conflict };
    return { unit: '元', factor: 1, explicit: false, unsupported: fu, conflict: conflict };
  }
  return { unit: '元', factor: 1, explicit: false };
}

// 简单表表头行（单一真源）：_hasSimpleHeader 与 _parseSimpleSheet 共用，避免两处判据漂移。
// 必须同时有「标签列」与「数值列」——只有「明细项目/数量」不算简单表；
// 值列先剥掉尾部单位括号（本期金额（元）→ 本期金额），并接受 本年累计金额 等叫法。
var SIMPLE_VALUE_RE = /^(本期(金额|数|发生额)?|本月金额|金额|余额|期初余额|期初数|期末余额|期末数|本年累计(金额|数|发生额)?)$/;
var SIMPLE_LABEL_RE = /^(科目|项目|明细项目|名称)$/;
function _simpleHeaderIndex(rows) {
  for (var r = 0; r < Math.min(rows.length, 6); r++) {
    var cells = (rows[r] || []).map(_hdrLabel);
    var hasLabel = false, hasValue = false;
    for (var i = 0; i < cells.length; i++) {
      if (SIMPLE_LABEL_RE.test(cells[i])) hasLabel = true;
      if (SIMPLE_VALUE_RE.test(cells[i])) hasValue = true;
    }
    if (hasLabel && hasValue) return r;
  }
  return -1;
}

// 简单表只有自带表头时才允许取数（防止非报表页的孤立数字被当作补缺值）
function _hasSimpleHeader(rows) {
  return _simpleHeaderIndex(rows) >= 0;
}

// other 表被拒时不再静默：表内含**精确**科目（ACC_SPEC 全等）却没识别出简单表表头，
// 由 parseWorkbook 在主循环后补 note（仅当已有结构化报表、该表确实未被采用时）。
function _hasExactAccount(rows) {
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || [];
    for (var c = 0; c < row.length; c++) {
      var label = _normLabel(row[c]);
      if (!label) continue;
      for (var k in ACC_SPEC) if (ACC_SPEC[k].indexOf(label) >= 0) return true;
    }
  }
  return false;
}

// 简单表自带的年份：优先工作表名；标题行要求「期间形」或「短标题 + 表」，
// 且排除 收到/支付/发票/凭证 一类注记，避免「2024年收到发票」把整表归错年。
function _simpleYear(rows, name) {
  var nm = String(name || '').match(/(20\d{2})\s*年/);
  if (nm) return parseInt(nm[1], 10);
  var hIdx = _simpleHeaderIndex(rows);
  var limit = hIdx >= 0 ? hIdx : Math.min(rows.length, 3);
  for (var r = 0; r < limit; r++) {
    var cells = (rows[r] || []).filter(function (x) { return String(x == null ? '' : x).trim() !== ''; });
    if (!cells.length || cells.length > 2) continue;
    var line = cells.join(' ').trim();
    if (/(打印|导出|报送|生成|制表|收到|支付|发票|凭证|发生|调整|备注|说明|附注|对比|目录|索引)/.test(line)) continue;
    var ym = line.match(/^(20\d{2})\s*年/);
    if (!ym) continue;
    var tail = line.slice(ym[0].length).trim();
    if (!tail) return parseInt(ym[1], 10);
    // ① 期间形：度 / 12月 / 1-12月 / 1月-12月 / 12月31日，可带尾部（单位）
    if (/^(?:度|\d{1,2}\s*月?(?:\s*[-~至—–]\s*\d{1,2}\s*月?)?(?:\s*\d{1,2}\s*日)?)\s*(?:[（(][^）)]*[）)])?\s*$/.test(tail)) {
      return parseInt(ym[1], 10);
    }
    // ② 短标题：1-12 个非空格字符（可带尾部单位括号），如「收入表」「收入（万元）」
    if (/^[^\s]{1,12}\s*(?:[（(][^）)]*[）)])?\s*$/.test(tail)) {
      return parseInt(ym[1], 10);
    }
  }
  return null;
}

function _cellRef(r, c) {
  var s = '';
  c = c + 1;
  while (c > 0) {
    var m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s + (r + 1);
}

function _detectYear(text, name, stmt) {
  var nameYear = null;
  var nm = String(name || '').match(/(?:^|[^\d])(20\d{2}|2[0-9])\s*年/);
  if (nm) nameYear = nm[1].length === 2 ? 2000 + parseInt(nm[1], 10) : parseInt(nm[1], 10);
  // 税局/汇总表：sheet 名里的年份更权威（正文可能含打印/报送日期）
  if (stmt === 'tax' && nameYear) return nameYear;
  var lines = String(text || '').split(/\s{2,}/);
  var titleYear = null, bodyYear = null;
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line) continue;
    if (/(打印|导出|报送|生成|制表)/.test(line)) continue;
    var m = line.match(/(20\d{2})\s*年/);
    if (!m) continue;
    if (/12月31日|年度/.test(line) && titleYear === null) titleYear = parseInt(m[1], 10);
    if (bodyYear === null) bodyYear = parseInt(m[1], 10);
  }
  return titleYear || bodyYear || nameYear;
}

function _detectStatement(text, name) {
  var t = String(text || '') + ' ' + String(name || '');
  if (/资产负债表|负债和所有者权益/.test(t)) return 'balance';
  if (/税局(比对|对比)|账面数据|取得的发票/.test(t)) return 'tax';
  if (/利润表|营业收入|营业成本/.test(t)) return 'income';
  if (/管理费用|税金及附加|财务费用|固定资产|工资/.test(t)) return 'detail';
  return 'other';
}

// 表头单元格归一：去掉尾部单位括号（「期末余额（万元）」→「期末余额」），
// 否则整表会因表头「未识别」被跳过（单位已由 _detectUnit 识别并换算）。
function _hdrLabel(v) {
  return _normLabel(v).replace(/[（(](?:人民币)?\s*(?:万元|千元|元)\s*[）)]$/, '');
}

function _headerRow(rows, stmt) {
  function has(cells, arr) { return arr.some(function (x) { return cells.indexOf(x) >= 0; }); }
  for (var r = 0; r < Math.min(rows.length, 15); r++) {
    var cells = (rows[r] || []).map(_hdrLabel);
    if (stmt === 'balance' &&
        has(cells, ['期末余额', '期末数', '年末余额', '期末']) &&
        has(cells, ['年初余额', '年初数', '上年年末余额', '年初'])) return r;
    if (stmt === 'income' &&
        has(cells, ['本年累计金额', '本年累计数', '本年累计', '本期金额', '本期数'])) return r;
  }
  return -1;
}

function _blocks(rows, hIdx) {
  var hdr = rows[hIdx] || [];
  var labelCols = [];
  for (var c = 0; c < hdr.length; c++) {
    var h = _normLabel(hdr[c]);
    if (['资产', '负债和所有者权益', '项目', '科目', '名称'].indexOf(h) >= 0) labelCols.push(c);
  }
  if (!labelCols.length && hIdx > 0) {
    var prev = rows[hIdx - 1] || [];
    for (var pc = 0; pc < prev.length; pc++) {
      var ph = _normLabel(prev[pc]);
      if (['资产', '负债和所有者权益', '项目', '科目', '名称'].indexOf(ph) >= 0) labelCols.push(pc);
    }
  }
  if (!labelCols.length) labelCols.push(0);
  var blocks = [];
  for (var i = 0; i < labelCols.length; i++) {
    var start = labelCols[i];
    var end = (i + 1 < labelCols.length) ? labelCols[i + 1] : Math.max(hdr.length, (rows[hIdx + 1] || []).length);
    var vc = {}, names = {};
    for (var c2 = start + 1; c2 < end; c2++) {
      var h2 = _normLabel(hdr[c2]);
      if (!h2 || /^(行次|行号|序号|差异)$/.test(h2)) continue;
      var sem = null;
      if (h2.indexOf('期末') >= 0 || h2.indexOf('年末') >= 0) sem = 'current';
      else if (h2.indexOf('年初') >= 0 || h2.indexOf('上年年末') >= 0) sem = 'beginning';
      else if (h2.indexOf('本年累计') >= 0 || h2.indexOf('本期') >= 0) sem = 'current';
      else if (h2.indexOf('四季度') >= 0) sem = 'quarter';
      else if (h2.indexOf('上期') >= 0) sem = 'beginning';
      else if (h2.indexOf('金额') >= 0 && vc.current === undefined) sem = 'current';
      if (sem && vc[sem] === undefined) { vc[sem] = c2; names[sem] = h2; }
    }
    if (Object.keys(vc).length) blocks.push({ labelCol: start, valueCols: vc, names: names });
  }
  return blocks;
}

function _parseBlocksInto(rows, hIdx, stmt, year, sheetName, put, factor) {
  var blocks = _blocks(rows, hIdx);
  for (var bi = 0; bi < blocks.length; bi++) {
    var b = blocks[bi];
    for (var r = hIdx + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var label = _normLabel(row[b.labelCol]);
      if (!label) continue;
      var keys = [];
      for (var k in ACC_SPEC) if (ACC_SPEC[k].indexOf(label) >= 0) keys.push(k);
      if (!keys.length) continue;
      for (var sem in b.valueCols) {
        if (sem === 'quarter') continue;
        var v = _parseNum(row[b.valueCols[sem]]);
        if (v === null) continue;
        var y = year, prio = 100;
        if (stmt === 'balance' && sem === 'beginning') { y = (year === null ? null : year - 1); prio = 90; }
        if (stmt === 'income' && sem === 'beginning') { y = (year === null ? null : year - 1); prio = 60; }
        var src = { sheet: sheetName, cell: _cellRef(r, b.valueCols[sem]), column: b.names[sem] || sem };
        // 「年初/上期」列是派生年度：它不代表工作簿里另有一份该年度报表。
        // F10 判断「工作簿是否含多个年度」时必须把这类年份排除，否则单张资产负债表
        // 会凭空造出 2024，使简单表永远无法并入。
        if (sem === 'beginning') src.derived = true;
        for (var ki = 0; ki < keys.length; ki++) put(y, keys[ki], v * factor, src, prio);
      }
    }
  }
}

function _sideKey(label) {
  if (label === '收入' || label === '营业收入' || label === '营业总收入') return 'revenue';
  if (label === '成本' || label === '营业成本' || label === '营业总成本') return 'cost';
  if (label === '利润总额') return 'profitGross';
  if (label === '净利润') return 'profitNet';
  return null;
}

function _sideTable(rows, sheetName, put, factor) {
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || [];
    var yearCols = {};
    for (var c = 0; c < row.length; c++) {
      var m = _normLabel(row[c]).match(/^(20\d{2})年$/);
      if (m) yearCols[parseInt(m[1], 10)] = c;
    }
    var ys = Object.keys(yearCols).map(Number);
    if (ys.length < 2) continue;
    var firstYearCol = Math.min.apply(null, Object.keys(yearCols).map(function (k) { return yearCols[k]; }));
    var labelCol = -1;
    for (var lc = firstYearCol - 1; lc >= 0; lc--) {
      var hit = false;
      for (var rr2 = r + 1; rr2 < Math.min(rows.length, r + 30); rr2++) {
        if (_sideKey(_normLabel((rows[rr2] || [])[lc]))) { hit = true; break; }
      }
      if (hit) { labelCol = lc; break; }
    }
    if (labelCol < 0) continue;
    for (var r2 = r + 1; r2 < rows.length; r2++) {
      var rr = rows[r2] || [];
      var key = _sideKey(_normLabel(rr[labelCol]));
      if (!key) continue;
      for (var yi = 0; yi < ys.length; yi++) {
        var y = ys[yi], c2 = yearCols[y], v = _parseNum(rr[c2]);
        if (v !== null) put(y, key, v * factor, { sheet: sheetName, cell: _cellRef(r2, c2), column: y + '年' }, 80);
      }
    }
    break;
  }
}

function _taxSheet(rows, sheetName, year, put, factor) {
  var labelCol = -1, valueCol = -1, start = -1;
  // 先找“账面数据”表头（真实税局表：C 列）
  for (var r = 0; r < Math.min(rows.length, 12) && valueCol < 0; r++) {
    var row = rows[r] || [];
    for (var c = 0; c < row.length; c++) {
      if (_normLabel(row[c]).indexOf('账面数据') >= 0) { start = r; valueCol = c; break; }
    }
  }
  if (valueCol >= 0) {
    for (var lc = valueCol - 1; lc >= 0; lc--) {
      var lab = _normLabel((rows[start] || [])[lc]);
      // 「项目 | 行次 | 账面数据」布局：行次不是标签列，跳过它继续向左找
      if (lab && !/^(行次|行号|序号)$/.test(lab)) { labelCol = lc; break; }
    }
  }
  // 兜底：找营业收入所在行/列，并在其右侧选一个非“行次”的值列
  if (labelCol < 0) {
    for (var r2 = 0; r2 < Math.min(rows.length, 12); r2++) {
      var rr = rows[r2] || [];
      for (var c2 = 0; c2 < rr.length; c2++) {
        if (_normLabel(rr[c2]) === '营业收入') { labelCol = c2; start = r2; break; }
      }
      if (labelCol >= 0) break;
    }
    if (labelCol >= 0 && valueCol < 0) {
      for (var c3 = labelCol + 1; c3 < Math.min(labelCol + 4, (rows[start] || []).length); c3++) {
        var hdr = _normLabel((rows[0] || [])[c3]) || _normLabel((rows[start - 1] || [])[c3]) || _normLabel((rows[start] || [])[c3]);
        if (/行次|行号|序号/.test(hdr)) continue;
        if (/金额|账面|本年|本期|累计/.test(hdr) || _parseNum((rows[start] || [])[c3]) !== null) { valueCol = c3; break; }
      }
    }
  }
  if (labelCol < 0 || valueCol < 0) return false;
  var produced = false;
  for (var r3 = start; r3 < rows.length; r3++) {
    var rw = rows[r3] || [];
    var label = _normLabel(rw[labelCol]);
    var keys = [];
    for (var k in ACC_SPEC) if (ACC_SPEC[k].indexOf(label) >= 0) keys.push(k);
    if (!keys.length) continue;
    var v = _parseNum(rw[valueCol]);
    if (v === null) continue;
    produced = true;
    for (var ki = 0; ki < keys.length; ki++) {
      put(year, keys[ki], v * factor, { sheet: sheetName, cell: _cellRef(r3, valueCol), column: '账面数据' }, 70);
    }
  }
  return produced;
}

function _parseSimpleSheet(rows, sheetName, put, year, factor) {
  var hIdx = _simpleHeaderIndex(rows), labelCol = 0, currentCol = -1, skipCols = {};
  if (hIdx >= 0) {
    var hdr = (rows[hIdx] || []).map(_hdrLabel);
    for (var lc = 0; lc < hdr.length; lc++) {
      if (SIMPLE_LABEL_RE.test(hdr[lc])) { labelCol = lc; break; }
    }
    for (var c = 0; c < hdr.length; c++) {
      if (/^(行次|行号|序号)$/.test(hdr[c])) skipCols[c] = 1;
      if (SIMPLE_VALUE_RE.test(hdr[c])) currentCol = c;
    }
    if (currentCol < 0) {
      for (var c2 = 0; c2 < hdr.length; c2++) {
        if (c2 === labelCol || skipCols[c2]) continue;
        if (/金额|本期|余额|本年累计/.test(hdr[c2])) { currentCol = c2; break; }
      }
    }
  }
  for (var r2 = (hIdx >= 0 ? hIdx + 1 : 0); r2 < rows.length; r2++) {
    var row = rows[r2] || [];
    var label = _normLabel(row[labelCol]);
    if (!label) continue;
    var keys = [];
    for (var k in ACC_SPEC) if (ACC_SPEC[k].indexOf(label) >= 0) keys.push(k);
    if (!keys.length) continue;
    var v = null, vc = -1;
    if (currentCol >= 0) {
      v = _parseNum(row[currentCol]);
      vc = currentCol;
    } else {
      for (var c3 = labelCol + 1; c3 < Math.min(row.length, 6); c3++) {
        if (skipCols[c3]) continue;
        var n = _parseNum(row[c3]);
        if (n !== null) { v = n; vc = c3; break; }
      }
    }
    if (v === null) continue;
    for (var ki = 0; ki < keys.length; ki++) {
      put(year, keys[ki], v * factor, { sheet: sheetName, cell: _cellRef(r2, vc), column: '本期' }, 100);
    }
  }
}

function parseWorkbook(sheets) {
  var byYear = {}, meta = {}, detected = [], warnings = [], candidates = [];
  var unitEntries = [], simpleCandidates = [];
  var fallbackSkip = {};
  var otherNoHeader = {};   // other 表：有精确科目但表头不符（si → 1）
  var structuredSeen = false;
  function put(year, key, value, src, prio) {
    if (year === null || year === undefined || value === null || !isFinite(value)) return;
    candidates.push({ year: year, key: key, value: value, src: src, prio: prio });
    if (!byYear[year]) { byYear[year] = {}; meta[year] = { _src: {}, _prio: {} }; }
    var p = meta[year]._prio[key];
    if (p !== undefined && prio < p) return;
    byYear[year][key] = value;
    meta[year]._prio[key] = prio;
    meta[year]._src[key] = src;
  }

  for (var si = 0; si < sheets.length; si++) {
    var sheet = sheets[si] || {};
    var rows = sheet.data || [];
    var text = rows.slice(0, 12).map(function (row) {
      return (row || []).map(function (x) { return String(x == null ? '' : x); }).join(' ');
    }).join('  ');
    var stmt = _detectStatement(text, sheet.name);
    var year = _detectYear(text, sheet.name, stmt);
    var unitInfo = _detectUnit(rows, sheet.name);
    var factor = unitInfo.factor;
    if (unitInfo.unsupported) {
      warnings.push(sheet.name + '：标注单位为「' + unitInfo.unsupported +
                    '」，当前仅支持 元/千元/万元，已按元处理（可能相差数量级）；请改为标准单位后重试。');
    }
    if (unitInfo.conflict) {
      warnings.push(sheet.name + '：表内出现多种单位标注，已按首个单位「' + unitInfo.unit +
                    '」换算；请拆分工作表或统一单位后重试。');
    }
    var note = '';
    var structuredParsed = false;
    var simpleEligible = false;
    if (stmt === 'balance' || stmt === 'income') {
      var h = _headerRow(rows, stmt);
      // 有「行次 + 资产总计/负债合计」或「行次 + 本年累计」⇒ 是结构化报表；
      // 只有「科目 | 本期 | 上期」的简单表没有行次，不能因含“营业收入”就当成利润表。
      var looksStructured =
        (stmt === 'balance' && /行次/.test(text) && /资产总计|负债合计/.test(text)) ||
        (stmt === 'income' && /行次/.test(text) && /(本年累计|营业收入|营业成本)/.test(text));
      if (h >= 0 || looksStructured) structuredSeen = true;
      if (year === null) {
        if (looksStructured) {
          warnings.push(sheet.name + '：识别为' + (stmt === 'balance' ? '资产负债表' : '利润表') + '但未识别年份，已跳过。');
          note = '未识别年份，已跳过';
        } else {
          // 有「本期金额」等列但没有行次/年份 ⇒ 更像无年份简单表，按简单表处理（自带表头才取数）
          simpleEligible = _hasSimpleHeader(rows);
        }
      } else if (h < 0) {
        if (looksStructured) {
          warnings.push(sheet.name + '：识别为' + (stmt === 'balance' ? '资产负债表' : '利润表') +
                        '但表头未识别（需要「期末/年初」或「本年累计」列），已跳过（不退回按行取数）。');
          note = '表头未识别，已跳过';
        } else {
          simpleEligible = _hasSimpleHeader(rows);
        }
      } else {
        _parseBlocksInto(rows, h, stmt, year, sheet.name, put, factor);
        if (stmt === 'income') _sideTable(rows, sheet.name, put, factor);
        if (stmt === 'balance') note = '期末→' + year + '；年初→' + (year - 1);
        structuredParsed = true;
      }
    } else if (stmt === 'tax') {
      if (year === null) {
        note = '未识别年份，已跳过';
        // 无年份的税局表也不得被全局简单表兜底捡走（否则「已跳过」与本期有值并存）
        fallbackSkip[sheet.name] = 1;
      } else if (_taxSheet(rows, sheet.name, year, put, factor)) {
        // 解析成功才计入「有结构化数据」——否则简单表会被静默丢弃
        structuredSeen = true;
        structuredParsed = true;
      } else {
        warnings.push(sheet.name + '：税局比对表未识别到可用科目或「账面数据/金额」列，已跳过。');
        note = '表头未识别，已跳过';
        // 解析失败的税局表不得再被简单表兜底捡走：否则页面写「已跳过」而数据却出现在本期
        fallbackSkip[sheet.name] = 1;
      }
    } else if (stmt === 'detail') {
      // 明细表默认不作主取数；只有自带「科目/项目/本期/金额」表头时才当简单表候选
      simpleEligible = _hasSimpleHeader(rows);
      if (!simpleEligible) note = '明细表，未作为主取数';
    } else {
      // 其他表同样要求自带简单表表头：避免「说明」之类页面里的孤立数字被当作补缺值
      simpleEligible = _hasSimpleHeader(rows);
      if (!simpleEligible && _hasExactAccount(rows)) otherNoHeader[si] = 1;
    }
    if (simpleEligible) simpleCandidates.push({ name: sheet.name, rows: rows, unit: unitInfo });
    if (structuredParsed) unitEntries.push({ name: sheet.name, unit: unitInfo.unit, explicit: unitInfo.explicit,
                                             unsupported: unitInfo.unsupported || null });
    detected.push({ name: sheet.name, statement: stmt, year: year, note: note,
                    unit: unitInfo.unit, unitExplicit: unitInfo.explicit,
                    unitUnsupported: unitInfo.unsupported || null,
                    unitConflict: !!unitInfo.conflict });
  }

  // other 表有精确科目却没识别出简单表表头：仅在存在结构化报表（该表已被拒）时提示，
  // 避免「不静默」变成误导（无结构化时该表仍走兜底取数）。
  if (structuredSeen) {
    for (var on in otherNoHeader) {
      if (detected[on]) detected[on].note = '未识别简单表表头（需 科目/项目 + 本期/金额 列）';
    }
  }

  // ── 简单表：不静默丢弃（F10）＋ 逐表单位汇总（单位提示） ──
  var structuredYears = Object.keys(byYear).map(Number)
    .filter(function (y) { return y > 0; })
    .sort(function (a, b) { return a - b; });
  // 「独立年度」= 至少有一条非派生来源（期末/本年/税局/侧表）的年份。
  // 只有「年初/上期」列的年份不单独计入，见 _parseBlocksInto 里 src.derived 的说明。
  var independentYears = structuredYears.filter(function (y) {
    return candidates.some(function (c) { return c.year === y && !c.src.derived; });
  });
  var simpleTables = [];
  function _keysLabel(keys) {
    return keys.map(function (k) { return ACCOUNT_META[k] || k; }).join('、');
  }
  if (!structuredSeen) {
    // 没有任何结构化报表：所有工作表按简单表兜底，落在「本期」(0)；
    // 若表名/标题已标注年份，则按该年取数（F10 在无结构化场景同样生效）
    for (var sj = 0; sj < sheets.length; sj++) {
      var sh = sheets[sj] || {};
      if (fallbackSkip[sh.name]) continue;   // 解析失败的税局表不在此兜底
      var su = _detectUnit(sh.data || [], sh.name);
      var fy = _simpleYear(sh.data || [], sh.name);
      if (fy !== null) {
        var nowY = (typeof Date !== 'undefined') ? new Date().getFullYear() : 2100;
        if (fy < 2000 || fy > nowY + 1) {
          warnings.push('「' + sh.name + '」标注年份 ' + fy + ' 不合理（应在 2000–' + (nowY + 1) +
                        ' 之间），已按本期处理。');
          fy = null;
        }
      }
      var seenKeys = {};
      _parseSimpleSheet(sh.data || [], sh.name, function (y, k, v, src, prio) {
        put(y, k, v, src, prio); seenKeys[k] = 1;
      }, fy === null ? 0 : fy, su.factor);
      if (Object.keys(seenKeys).length) {
        unitEntries.push({ name: sh.name, unit: su.unit, explicit: su.explicit, unsupported: su.unsupported || null });
        if (fy !== null) {
          for (var dn1 = 0; dn1 < detected.length; dn1++) {
            if (detected[dn1].name === sh.name) { detected[dn1].note = '简单表（标注 ' + fy + ' 年）· 按该年取数'; break; }
          }
        }
      }
    }
  } else {
    for (var sc = 0; sc < simpleCandidates.length; sc++) {
      var cand = simpleCandidates[sc];
      var acc = {}, order = [];
      _parseSimpleSheet(cand.rows, cand.name, function (y, k, v, src) {
        if (!acc[k]) { acc[k] = { key: k, value: v, cell: src.cell }; order.push(k); }
      }, 0, cand.unit.factor);
      if (!order.length) continue;

      var explicitYear = _simpleYear(cand.rows, cand.name);
      // 标注年份不得越出现有独立年度太远（防「2030年收入表」凭空造年度并抢占默认视图）
      if (explicitYear !== null && independentYears.length > 0) {
        var minY = independentYears[0], maxY = independentYears[independentYears.length - 1];
        if (explicitYear < minY - 1 || explicitYear > maxY + 1) {
          warnings.push('「' + cand.name + '」标注年份 ' + explicitYear + ' 超出现有年度范围（' +
                        minY + '–' + maxY + '），已按未标年份处理。');
          explicitYear = null;
        }
      }
      var disp, targetYear = null, yearFrom = null;

      if (explicitYear !== null) {
        // 表名/标题已标注年份 ⇒ 直接按该年并入（仍只补缺），不再当成「未标年份」
        targetYear = explicitYear; yearFrom = 'label';
        for (var k1 = 0; k1 < order.length; k1++) {
          var a1 = acc[order[k1]];
          put(targetYear, a1.key, a1.value,
              { sheet: cand.name, cell: a1.cell, column: '本期', simple: true }, 40);
        }
        disp = 'merged';
        warnings.push('「' + cand.name + '」是简单表，已按表内/表名标注并入 ' + targetYear +
                      ' 年（仅补缺，不覆盖结构化报表）：' + _keysLabel(order) + '。');
      } else if (independentYears.length === 0) {
        // 只有解析失败的结构化表：没有可依的年度，按「本期」(0) 取数
        for (var k0 = 0; k0 < order.length; k0++) {
          var a0 = acc[order[k0]];
          put(0, a0.key, a0.value, { sheet: cand.name, cell: a0.cell, column: '本期' }, 100);
        }
        unitEntries.push({ name: cand.name, unit: cand.unit.unit, explicit: cand.unit.explicit,
                           unsupported: cand.unit.unsupported || null });
        for (var dn0 = 0; dn0 < detected.length; dn0++) {
          if (detected[dn0].name === cand.name) { detected[dn0].note = '简单表（未标年份）· 作为「本期」数据'; break; }
        }
        continue;
      } else if (independentYears.length === 1) {
        // 只有一个独立年度 ⇒ 归属唯一，可安全并入；仍以最低优先级“仅补缺”并在页面标注假设
        targetYear = independentYears[0]; yearFrom = 'assumed';
        for (var k2 = 0; k2 < order.length; k2++) {
          var a2 = acc[order[k2]];
          put(targetYear, a2.key, a2.value,
              { sheet: cand.name, cell: a2.cell, column: '本期', simple: true, assumed: true }, 40);
        }
        disp = 'merged';
        warnings.push('「' + cand.name + '」是未标年份的简单表，已按「本期」并入 ' + targetYear +
                      ' 年（仅补缺，不覆盖结构化报表）：' + _keysLabel(order) +
                      '；若该表属于其他期间，请在表内或工作表名标注年份。');
      } else {
        // 多个年度 ⇒ 归属不可判定，不猜：给出可操作警告并列出表名/科目
        disp = 'skipped';
        warnings.push('「' + cand.name + '」是未标年份的简单表：已识别到 ' + _keysLabel(order) +
                      '；工作簿含多个年度（' + independentYears.join('/') +
                      '），无法安全判定归属，未并入年度视图。请在该表内或工作表名标注年份（如「2025年」）后重新拖入。');
      }
      simpleTables.push({
        name: cand.name, unit: cand.unit.unit, disposition: disp, year: targetYear, yearFrom: yearFrom,
        accounts: order.map(function (k) { return acc[k]; })
      });
      // 只有真正参与计算的表才计入单位汇总；被跳过的表仍在「识别到的工作表」里显示单位
      if (disp !== 'skipped') unitEntries.push({ name: cand.name, unit: cand.unit.unit, explicit: cand.unit.explicit,
                                                 unsupported: cand.unit.unsupported || null });
      for (var dn = 0; dn < detected.length; dn++) {
        if (detected[dn].name === cand.name) {
          detected[dn].note = disp === 'skipped'
            ? '简单表（未标年份）· 未并入（工作簿含多个年度）'
            : (yearFrom === 'label'
                ? '简单表（标注 ' + targetYear + ' 年）· 已并入（仅补缺）'
                : '简单表（未标年份）· 已按本期并入 ' + targetYear + ' 年（仅补缺）');
          break;
        }
      }
    }
  }

  var unitSeen = {};
  for (var ue = 0; ue < unitEntries.length; ue++) unitSeen[unitEntries[ue].unit] = 1;
  var unitsUsed = Object.keys(unitSeen);
  var unitSummary = {
    sheets: detected.map(function (x) { return { name: x.name, unit: x.unit, explicit: x.unitExplicit }; }),
    used: unitEntries.map(function (e) {
      return { name: e.name, unit: e.unit, explicit: e.explicit, unsupported: e.unsupported || null };
    }),
    mixed: unitsUsed.length > 1,
    units: unitsUsed,
    target: '元'
  };
  if (unitSummary.mixed) {
    warnings.push('金额单位不一致：' + unitEntries.map(function (e) { return '「' + e.name + '」= ' + e.unit; }).join('，') +
                  '；已分别换算为元后计算（万元×10000、千元×1000）。');
  }

  for (var yk in byYear) {
    var d = byYear[yk];
    if (d.profitNet !== undefined) { d.profit = d.profitNet; d.profitKind = '净利润'; }
    else if (d.profitGross !== undefined) { d.profit = d.profitGross; d.profitKind = '利润总额'; }
  }

  var years = Object.keys(byYear).map(Number).filter(function (y) { return y > 0; }).sort(function (a, b) { return b - a; });
  if (!years.length && byYear[0]) years = [0];
  // 默认年份优先取「独立年度」（有期末/本年/税局等结构化来源），
  // 防止一张标注年份的简单表（如「2024年收入表」）把默认视图从 2025 抢走。
  var indepDesc = independentYears.slice().sort(function (a, b) { return b - a; });
  function _pickYear(list, needAssets) {
    for (var pi = 0; pi < list.length; pi++) {
      var yy = list[pi];
      if (yy <= 0 || !byYear[yy]) continue;
      if (byYear[yy].revenue !== undefined && (!needAssets || byYear[yy].totalAssets !== undefined)) return yy;
    }
    return null;
  }
  var defaultYear = _pickYear(indepDesc, true) || _pickYear(indepDesc, false);
  if (defaultYear === null && indepDesc.length) defaultYear = indepDesc[0];
  if (defaultYear === null) defaultYear = _pickYear(years, true) || _pickYear(years, false);
  if (defaultYear === null) defaultYear = years.length ? years[0] : null;

  function _dedupe(vals) {
    var out = [], seen = {};
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      var k = v.src.sheet + '|' + v.src.cell + '|' + v.value;
      if (!seen[k]) { seen[k] = 1; out.push(v); }
    }
    return out;
  }
  function _check(key, year) {
    var vals = _dedupe(candidates.filter(function (c) { return c.key === key && c.year === year; }));
    if (vals.length < 2) return null;
    var lo = Math.min.apply(null, vals.map(function (v) { return v.value; }));
    var hi = Math.max.apply(null, vals.map(function (v) { return v.value; }));
    // 采用来源（F14）：与「识别科目与来源」同一条判据 —— 直接比 meta 里最终胜出的 src 对象，
    // 不做第二套优先级计算（UB3：同一状态不留两处表述）。
    var adoptedSrc = meta[year] && meta[year]._src[key];
    vals.forEach(function (v) {
      v.adopted = !!(adoptedSrc && adoptedSrc.sheet === v.src.sheet && adoptedSrc.cell === v.src.cell);
      v.assumed = !!v.src.assumed;
    });
    return { year: year, key: key, values: vals, ok: (hi - lo) <= 0.01,
             assumed: vals.some(function (v) { return v.assumed; }) };
  }
  var checks = [];
  var checkKeys = ['revenue', 'cost', 'totalAssets', 'totalLiabilities', 'adminExpense', 'financeExpense', 'equity'];
  years.forEach(function (y) {
    checkKeys.forEach(function (k) {
      var c = _check(k, y);
      if (c) checks.push(c);
    });
  });

  // 动态边界提示：不硬编码 2024/2023
  Object.keys(byYear).map(Number).sort(function (a, b) { return a - b; }).forEach(function (y) {
    if (y <= 0) return;
    var src = meta[y] && meta[y]._src.totalAssets;
    if (src && src.column && src.column.indexOf('年初') >= 0) {
      warnings.push(y + ' 年资产负债表取自「' + src.sheet + '」的「' + src.column + '」（' + y + '-12-31），不是独立的 ' + y + ' 年报表。');
    }
    if (byYear[y].revenue !== undefined && byYear[y].totalAssets === undefined) {
      warnings.push(y + ' 年只有利润表数据（收入/成本/利润），缺少资产负债表，资产负债率/ROE 不可计算。');
    }
  });
  if (defaultYear !== null) {
    var miss = [];
    ['revenue', 'cost', 'totalAssets', 'totalLiabilities'].forEach(function (k) {
      if (byYear[defaultYear][k] === undefined) miss.push(k);
    });
    if (miss.length) warnings.push('默认年份缺少：' + miss.join('、') + '（结果中会显示为空）。');
  }

  var sources = {};
  for (var sy in meta) {
    sources[sy] = {};
    for (var sk in meta[sy]._src) {
      var so = meta[sy]._src[sk];
      // `derived` 只服务 F10 的年度归属判断，不进入对外 sources（保持 schema 稳定）
      if (so && so.derived) {
        var pub = {};
        for (var pf in so) if (pf !== 'derived') pub[pf] = so[pf];
        sources[sy][sk] = pub;
      } else {
        sources[sy][sk] = so;
      }
    }
  }

  return {
    byYear: byYear,
    sources: sources,
    detected: detected,
    warnings: warnings,
    checks: checks,
    years: years,
    defaultYear: defaultYear,
    current: defaultYear !== null ? (byYear[defaultYear] || {}) : {},
    // 未标年份的简单表（F10）：disposition = merged / skipped；skipped 时页面必须展示，不得静默丢弃
    simple: simpleTables,
    // 逐表金额单位（单位提示）：mixed=true 时页面提示「已按 … 换算为元」
    unitInfo: unitSummary
  };
}
