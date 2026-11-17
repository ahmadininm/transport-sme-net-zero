#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Reproduce every analysis, table and figure reported in the manuscript.

Inputs (all in the same folder as this script)
    Analysis.xlsx              worksheet AT1 holds the 44 cleaned survey responses
    qualitative_coding.csv     respondent-level coding register for Q33 to Q37
    scope_classification.csv   respondent-level transport-activity classification

Outputs
    figures/Figure_1.tif, Figure_2.tif, Figure_3.tif, Figure_S1.tif, Figure_S2.tif
    outputs/*.csv                 tables behind the manuscript and supplement
    outputs/validation_log.txt    comparison with the values reported in manuscript

Usage
    python reproduce.py                      recompute, redraw and validate
    python reproduce.py --update-workbook    also refresh the Python-derived sheets
                                                (PY_Bootstrap, PY_Check) in the workbook
    python reproduce.py --no-figures         skip figure generation

Nothing in this script is a stored result. All statistics are computed from AT1 and the two
registers. The "Reported" column of the workbook's Validation sheet is read only to
check the recomputed values against the manuscript.
"""
from __future__ import annotations
import argparse, csv, math, re, sys
from collections import defaultdict
from itertools import combinations
from pathlib import Path

import numpy as np
import openpyxl
from scipy import stats

HERE = Path(__file__).resolve().parent
WORKBOOK = HERE / 'Analysis.xlsx'
CODING_CSV = HERE / 'qualitative_coding.csv'
SCOPE_CSV = HERE / 'scope_classification.csv'

N_EXPECTED = 44
ITEMS = [f'Q{i}' for i in range(6, 33)]
OPEN_QUESTIONS = ['Q33', 'Q34', 'Q35', 'Q36', 'Q37']
LIKERT = {'completely disagree': 1, 'disagree': 2, 'neutral': 3, 'agree': 4, 'completely agree': 5}
NO_RESPONSE_ENTRIES = {'', '-', 'n/a', 'na'}          # blank, hyphen or not-applicable marker
SIZE_LEVELS = ['1-9 employees', '10-49 employees', '50-249 employees']
AGE_LEVELS = ['Less than 1 year', '1-5 years', '6-10 years', 'More than 10 years']
MODES = ['road', 'rail', 'air', 'maritime']
BOOTSTRAP_SEED = 20260824
BOOTSTRAP_RESAMPLES = 5000
ALPHA = 0.05
Z975 = stats.norm.ppf(0.975)

# barrier statements used in the co-occurrence summary (agreement = rating of 4 or 5)
COST_RESOURCE_BARRIERS = ['Q10', 'Q12', 'Q15']                       # technology cost, access to funding, resource constraints
OTHER_BARRIERS = ['Q11', 'Q13', 'Q14', 'Q16', 'Q17']                 # expertise, information, customer demand, tenancy, time
SUPPORT_ITEMS = ['Q28', 'Q29', 'Q30', 'Q31', 'Q32']
# themes plotted in Figure 2 (open-text answers to Q36, Q34 and Q37)
FIGURE2_CODES = ['FINANCE_INVESTMENT', 'GUIDANCE_TRAINING_INFORMATION', 'POLICY_REGULATION_CERTAINTY', 'MEASUREMENT_LCA_STANDARDS',
                 'TECHNOLOGY_FEASIBILITY_TESTING', 'TIME_PEOPLE_CAPACITY', 'SUPPLY_CHAIN_MATERIALS', 'CUSTOMER_MARKET_PROCUREMENT',
                 'COLLABORATION_NETWORKS']
FIGURE2_LABELS = ['Finance, cost or return', 'Guidance, training or advice', 'Policy and regulatory certainty', 'Measurement and carbon accounting',
                  'Technological feasibility and testing', 'Time, staffing or capacity', 'Supply chain and materials', 'Customer demand and procurement',
                  'Collaboration and joint projects']

# closed support item -> open-text code used in the closed/open comparison
CONTRAST = [('Financial incentives', 'Q28', 'FINANCE_INVESTMENT'),
            ('Clear and consistent regulation', 'Q31', 'POLICY_REGULATION_CERTAINTY'),
            ('Collaboration', 'Q30', 'COLLABORATION_NETWORKS'),
            ('Training, expertise and guidance', 'Q29', 'GUIDANCE_TRAINING_INFORMATION')]

# Sensitivity of the collaboration counts to two borderline coding decisions (reported in Supplementary Table S7).
# strict: drops one Q37 answer that refers to expertise provided alongside funding, with no partner named;
# inclusive: adds one Q34 answer that refers to connecting with suppliers.
COLLABORATION_SENSITIVITY = {'strict': {'exclude': [('Q37', 29)], 'include': []},
                             'inclusive': {'exclude': [], 'include': [('Q34', 26)]}}

# codes shown for Q34, Q36 and Q37 in Table 3 (reported even when no response was assigned to them)
TABLE3_CODES = ['FINANCE_INVESTMENT', 'GUIDANCE_TRAINING_INFORMATION', 'POLICY_REGULATION_CERTAINTY', 'MEASUREMENT_LCA_STANDARDS',
                'TECHNOLOGY_FEASIBILITY_TESTING', 'TIME_PEOPLE_CAPACITY', 'SUPPLY_CHAIN_MATERIALS', 'CUSTOMER_MARKET_PROCUREMENT',
                'COLLABORATION_NETWORKS', 'BUSINESS_STAGE', 'INFRASTRUCTURE_ENERGY_PREMISES', 'ACTION_MOTIVATION_SIMPLICITY']

ITEM_LABELS = {
    'Q6': 'Adopted some low-carbon technology', 'Q7': 'Regularly monitors and reports GHG emissions',
    'Q8': 'Formal environmental policy in place', 'Q9': 'Received external support',
    'Q10': 'High technology cost is a significant barrier', 'Q11': 'Lacks technical expertise',
    'Q12': 'Difficulty accessing funding', 'Q13': 'Not enough information on benefits',
    'Q14': 'Customer demand too weak to justify investment', 'Q15': 'Resource constraints hinder investment',
    'Q16': 'Short-term tenancy limits investment', 'Q17': 'Lacks time for decarbonisation',
    'Q18': 'COVID-19 reduced ability to invest', 'Q19': 'Brexit added complexity and uncertainty',
    'Q20': 'Full design authority over products', 'Q21': 'Full control over material selection',
    'Q22': 'Control over process technology selection', 'Q23': 'Full control over fuel and energy supply',
    'Q24': 'Policies and incentives have had a positive influence', 'Q25': 'Customer requirements drive low-carbon practices',
    'Q26': 'Suppliers capable and willing to support', 'Q27': 'Industry associations provide adequate support',
    'Q28': 'Financial incentives would significantly help', 'Q29': 'Technical training and expertise are crucial',
    'Q30': 'Collaboration with other businesses would help', 'Q31': 'Clear, consistent regulatory guidance would help',
    'Q32': 'Case studies from similar businesses would help'}
ITEM_GROUPS = [('Current practices', range(6, 10)), ('Challenges and barriers', range(10, 20)),
               ('Organisational control', range(20, 24)), ('Policy and stakeholder influence', range(24, 28)),
               ('Support and resources', range(28, 33))]


# ----------------------------------------------------------------------------- input
def clean(value) -> str:
    return '' if value is None else re.sub(r'\s+', ' ', str(value)).strip()


def load_at1(path: Path):
    """Read the 44 respondent records from worksheet AT1 (rows 5 to 48, columns ID and Q1 to Q37)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if 'AT1' not in wb.sheetnames:
        raise SystemExit(f'Worksheet AT1 not found in {path.name}')
    rows = [r for r in wb['AT1'].iter_rows(min_row=5, max_col=38, values_only=True) if r[0] not in (None, '')]
    wb.close()
    data = []
    for r in rows:
        rec = {'ID': int(str(r[0]).strip())}
        for q in range(1, 38):
            rec[f'Q{q}'] = clean(r[q])
        data.append(rec)
    ids = [d['ID'] for d in data]
    if len(ids) != N_EXPECTED or len(set(ids)) != N_EXPECTED:
        raise SystemExit(f'AT1 must contain {N_EXPECTED} distinct respondent records; found {len(ids)}')
    for d in data:
        for q in ITEMS:
            key = d[q].lower()
            if key not in LIKERT:
                raise SystemExit(f'Unrecognised response {d[q]!r} for respondent {d["ID"]}, {q}')
            d[q + '_n'] = LIKERT[key]
        d['size'] = d['Q1'].strip(' ;'); d['age'] = d['Q2'].strip(' ;'); d['startup'] = d['Q3'].strip(' ;')
        if d['size'] not in SIZE_LEVELS or d['age'] not in AGE_LEVELS or d['startup'] not in ('Yes', 'No'):
            raise SystemExit(f'Unrecognised descriptor for respondent {d["ID"]}')
        d['modes'] = [m for m in MODES if re.search(rf'\b{m}\b', d['Q4'].lower())]
    return data


def load_scope(path: Path, data):
    scope = {}
    with open(path, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            scope[int(row['respondent_id'])] = row
    if set(scope) != {d['ID'] for d in data}:
        raise SystemExit('scope_classification.csv does not cover exactly the respondents in AT1')
    for d in data:
        s = scope[d['ID']]
        d['scope'] = s['scope_category']
        d['direct'] = d['scope'] == 'Direct transport'
        if d['modes']:                                   # a named mode always means direct transport
            if not d['direct']:
                raise SystemExit(f'Respondent {d["ID"]} names a transport mode but is not classified as direct')
            derived = 'More than one mode' if len(d['modes']) > 1 else d['modes'][0].capitalize()
            if s['mode_category'] != derived:
                raise SystemExit(f'Mode category for respondent {d["ID"]} disagrees with the sector answer')
        d['mode_category'] = s['mode_category'] if d['direct'] else d['scope']
    return scope


def load_coding(path: Path, data):
    """Return {(question, code): set(ids)} and {question: set(non-response ids)} after consistency checks."""
    assign, noresp = defaultdict(set), defaultdict(set)
    seen = set()
    with open(path, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            i, q, c = int(row['respondent_id']), row['question'], row['code']
            if (i, q, c) in seen:
                raise SystemExit(f'Duplicate register row {(i, q, c)}')
            seen.add((i, q, c))
            (noresp[q] if c == 'NO_RESPONSE' else assign[(q, c)]).add(i)
    ids = {d['ID'] for d in data}
    for q in OPEN_QUESTIONS:
        rule = {d['ID'] for d in data if d[q].lower() in NO_RESPONSE_ENTRIES}
        if rule != noresp[q]:
            raise SystemExit(f'{q}: register non-responses {sorted(noresp[q])} differ from the rule applied to AT1 {sorted(rule)}')
        coded = set().union(*[s for (qq, _), s in assign.items() if qq == q]) if any(qq == q for qq, _ in assign) else set()
        if coded != ids - rule:
            raise SystemExit(f'{q}: responses without a code or codes on non-responses: {sorted(coded ^ (ids - rule))}')
    return dict(assign), dict(noresp)


# ----------------------------------------------------------------------------- statistics
def wilson(k: int, n: int):
    p = k / n; z = Z975; den = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / den
    half = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / den
    return centre - half, centre + half


def holm(p):
    p = np.asarray(p, float); m = len(p); order = np.argsort(p, kind='stable')
    adj = np.maximum.accumulate(p[order] * (m - np.arange(m)))
    out = np.empty(m); out[order] = np.clip(adj, 0, 1); return out


def benjamini_hochberg(p):
    p = np.asarray(p, float); m = len(p); order = np.argsort(p, kind='stable')
    adj = np.minimum.accumulate((p[order] * m / (np.arange(m) + 1))[::-1])[::-1]
    out = np.empty(m); out[order] = np.clip(adj, 0, 1); return out


def bootstrap_spearman_ci(x, y, seed_key):
    """Percentile bootstrap interval for Spearman's rho; the seed is specific to the item pair."""
    rng = np.random.default_rng([BOOTSTRAP_SEED, *seed_key])
    n = len(x); idx = rng.integers(0, n, size=(BOOTSTRAP_RESAMPLES, n))
    rx = stats.rankdata(x[idx], axis=1); ry = stats.rankdata(y[idx], axis=1)
    rx -= rx.mean(axis=1, keepdims=True); ry -= ry.mean(axis=1, keepdims=True)
    den = np.sqrt((rx ** 2).sum(axis=1) * (ry ** 2).sum(axis=1))
    ok = den > 0                                         # drop resamples in which an item is constant
    rho = (rx * ry).sum(axis=1)[ok] / den[ok]
    return float(np.percentile(rho, 2.5)), float(np.percentile(rho, 97.5)), int(ok.sum())


def analyse(data, assign, noresp):
    R = {}
    X = np.array([[d[q + '_n'] for q in ITEMS] for d in data], dtype=float)
    direct = np.array([d['direct'] for d in data])
    ids = np.array([d['ID'] for d in data])
    n = len(data); nd = int(direct.sum())
    R['n'], R['n_direct'] = n, nd

    # ---- sample
    count = lambda key, val: sum(1 for d in data if d[key] == val)
    R['sample'] = {'size': {s: count('size', s) for s in SIZE_LEVELS}, 'age': {a: count('age', a) for a in AGE_LEVELS},
                   'startup': {s: count('startup', s) for s in ('Yes', 'No')},
                   'scope': {}, 'mode': {}}
    for d in data:
        R['sample']['scope'][d['scope']] = R['sample']['scope'].get(d['scope'], 0) + 1
        if d['direct']:
            R['sample']['mode'][d['mode_category']] = R['sample']['mode'].get(d['mode_category'], 0) + 1
    R['sample']['five_years_or_less'] = count('age', 'Less than 1 year') + count('age', '1-5 years')
    R['sample']['nonstartup_over_10'] = sum(1 for d in data if d['startup'] == 'No' and d['age'] == 'More than 10 years')

    # ---- item summaries
    items = []
    for j, q in enumerate(ITEMS):
        x = X[:, j]; xd = x[direct]
        counts = [int((x == k).sum()) for k in range(1, 6)]
        agree = counts[3] + counts[4]; lo, hi = wilson(agree, n)
        items.append({'item': q, 'label': ITEM_LABELS[q], 'counts': counts, 'n': n,
                      'median': float(np.median(x)), 'q1': float(np.percentile(x, 25)), 'q3': float(np.percentile(x, 75)),
                      'agree': agree, 'agree_pct': 100 * agree / n, 'ci_lo': 100 * lo, 'ci_hi': 100 * hi,
                      'disagree': counts[0] + counts[1], 'neutral': counts[2],
                      'direct_agree': int((xd >= 4).sum()), 'direct_agree_pct': 100 * float((xd >= 4).mean()),
                      'direct_median': float(np.median(xd))})
    R['items'] = items

    # ---- co-occurrence of barriers and breadth of support endorsement (counts of items agreed per respondent)
    col = lambda q: X[:, ITEMS.index(q)] >= 4
    n_cost = sum(col(q).astype(int) for q in COST_RESOURCE_BARRIERS); n_other = sum(col(q).astype(int) for q in OTHER_BARRIERS)
    n_all = n_cost + n_other; n_sup = sum(col(q).astype(int) for q in SUPPORT_ITEMS)
    R['barriers'] = {'cost_any': int((n_cost >= 1).sum()), 'cost_two_plus': int((n_cost >= 2).sum()), 'cost_all': int((n_cost == 3).sum()),
                     'cost_and_other': int(((n_cost >= 1) & (n_other >= 1)).sum()), 'other_only': int(((n_cost == 0) & (n_other >= 1)).sum()),
                     'none': int((n_all == 0).sum()), 'two_plus': int((n_all >= 2).sum()), 'three_plus': int((n_all >= 3).sum()),
                     'median': float(np.median(n_all)), 'q1': float(np.percentile(n_all, 25)), 'q3': float(np.percentile(n_all, 75)),
                     'cost_any_direct': int(((n_cost >= 1) & direct).sum()), 'none_direct': int(((n_all == 0) & direct).sum()),
                     'support_all_five': int((n_sup == 5).sum()), 'support_four_plus': int((n_sup >= 4).sum()),
                     'distribution': [int((n_all == k).sum()) for k in range(len(COST_RESOURCE_BARRIERS) + len(OTHER_BARRIERS) + 1)]}

    # ---- Spearman correlations, all 351 pairs
    pairs = []
    for a, b in combinations(range(len(ITEMS)), 2):
        res = stats.spearmanr(X[:, a], X[:, b])
        pairs.append({'a': ITEMS[a], 'b': ITEMS[b], 'ia': a, 'ib': b, 'rho': float(res.statistic), 'p': float(res.pvalue)})
    pv = [p['p'] for p in pairs]
    for p, h, b in zip(pairs, holm(pv), benjamini_hochberg(pv)):
        p['holm'], p['bh'] = float(h), float(b)
    for p in pairs:
        if p['bh'] < ALPHA:
            p['ci_lo'], p['ci_hi'], p['boot_n'] = bootstrap_spearman_ci(X[:, p['ia']], X[:, p['ib']], (p['ia'], p['ib']))
    R['pairs'] = pairs
    R['corr_summary'] = {'pairs': len(pairs), 'raw': sum(p['p'] < ALPHA for p in pairs),
                         'holm': sum(p['holm'] < ALPHA for p in pairs), 'bh': sum(p['bh'] < ALPHA for p in pairs)}

    # ---- group comparisons
    size = np.array([SIZE_LEVELS.index(d['size']) for d in data]); startup = np.array([d['startup'] == 'Yes' for d in data])
    k = len(SIZE_LEVELS); kw, mw = [], []
    for j, q in enumerate(ITEMS):
        H, p = stats.kruskal(*[X[size == g, j] for g in range(k)])
        kw.append({'item': q, 'H': float(H), 'eta2': float((H - k + 1) / (n - k)), 'p': float(p)})
        a, b = X[startup, j], X[~startup, j]
        U, p2 = stats.mannwhitneyu(a, b, alternative='two-sided', method='asymptotic')
        mw.append({'item': q, 'U': float(U), 'r': float(2 * U / (len(a) * len(b)) - 1), 'p': float(p2)})
    for rows in (kw, mw):
        pv = [r['p'] for r in rows]
        for r, h, b in zip(rows, holm(pv), benjamini_hochberg(pv)):
            r['holm'], r['bh'] = float(h), float(b)
    R['kw'], R['mw'] = kw, mw
    R['group_n'] = {'size': [int((size == g).sum()) for g in range(k)], 'startup': [int(startup.sum()), int((~startup).sum())]}

    # ---- open-text coding
    D = set(ids[direct].tolist()); allids = set(ids.tolist())
    R['codes'] = [{'question': q, 'code': c, 'n': len(s), 'n_direct': len(s & D), 'ids': sorted(s)}
                  for (q, c), s in sorted(assign.items(), key=lambda kv: (kv[0][0], -len(kv[1]), kv[0][1]))]
    status = {}
    for q in OPEN_QUESTIONS:
        nr = noresp.get(q, set()); resp = allids - nr
        per = {i: sum(i in s for (qq, _), s in assign.items() if qq == q) for i in resp}
        status[q] = {'responses': len(resp), 'no_response': len(nr), 'responses_direct': len(resp & D),
                     'no_response_direct': len(nr & D), 'codes': sum(per.values()), 'multi': sum(v > 1 for v in per.values())}
    R['status'] = status
    words = [len(d[q].split()) for q in OPEN_QUESTIONS for d in data if d['ID'] not in noresp.get(q, set())]
    R['response_words'] = {'n': len(words), 'median': float(np.median(words)), 'q1': float(np.percentile(words, 25)),
                           'q3': float(np.percentile(words, 75))}

    # ---- prompted endorsement versus subsequent open mention
    code_ids = lambda q, c: assign.get((q, c), set())
    contrast = []
    for name, item, code in CONTRAST:
        agree = set(ids[X[:, ITEMS.index(item)] >= 4].tolist())
        s34, s37 = code_ids('Q34', code), code_ids('Q37', code); either = s34 | s37
        contrast.append({'name': name, 'item': item, 'code': code, 'agree': len(agree), 'q34': len(s34), 'q37': len(s37),
                         'either': len(either), 'agree_and_either': len(agree & either), 'either_not_agree': len(either - agree),
                         'agree_d': len(agree & D), 'q34_d': len(s34 & D), 'q37_d': len(s37 & D), 'either_d': len(either & D),
                         'agree_and_either_d': len(agree & either & D)})
    R['contrast'] = contrast

    # ---- collaboration counts under alternative readings of two borderline answers
    agree30 = set(ids[X[:, ITEMS.index('Q30')] >= 4].tolist()); sens = {}
    for name, spec in {'as_coded': {'exclude': [], 'include': []}, **COLLABORATION_SENSITIVITY}.items():
        s34, s37 = set(code_ids('Q34', 'COLLABORATION_NETWORKS')), set(code_ids('Q37', 'COLLABORATION_NETWORKS'))
        for q, i in spec['exclude']: (s34 if q == 'Q34' else s37).discard(i)
        for q, i in spec['include']: (s34 if q == 'Q34' else s37).add(i)
        sens[name] = {'q34': len(s34), 'q37': len(s37), 'either': len(s34 | s37), 'agree_and_either': len((s34 | s37) & agree30)}
    R['collab_sensitivity'] = sens

    # ---- figures quoted in the text
    fin36 = code_ids('Q36', 'FINANCE_INVESTMENT')
    other36 = set().union(*[s for (q, c), s in assign.items() if q == 'Q36' and c != 'FINANCE_INVESTMENT'])
    g34, m34, f34 = code_ids('Q34', 'GUIDANCE_TRAINING_INFORMATION'), code_ids('Q34', 'MEASUREMENT_LCA_STANDARDS'), code_ids('Q34', 'FINANCE_INVESTMENT')
    R['text'] = {'finance_q36_with_other_code': len(fin36 & other36), 'guidance_q34_also_measurement': len(g34 & m34),
                 'finance_and_guidance_q34': len(f34 & g34),
                 'case_study_mentions': sum(1 for q in OPEN_QUESTIONS for d in data if 'case stud' in d[q].lower())}
    R['_X'], R['_direct'], R['_ids'] = X, direct, ids
    return R


# ----------------------------------------------------------------------------- validation keys
def key_values(R):
    """Flat {key: value} dictionary of every value reported in the manuscript and supplement."""
    K = {'sample.n': R['n'], 'sample.direct': R['n_direct']}
    for s, tag in zip(SIZE_LEVELS, ('1-9', '10-49', '50-249')): K[f'sample.size.{tag}'] = R['sample']['size'][s]
    for a, tag in zip(AGE_LEVELS, ('lt1', '1-5', '6-10', 'gt10')): K[f'sample.age.{tag}'] = R['sample']['age'][a]
    K['sample.startup.yes'] = R['sample']['startup']['Yes']; K['sample.startup.no'] = R['sample']['startup']['No']
    for m in ('Road', 'Rail', 'Air', 'Maritime', 'More than one mode'): K[f'sample.mode.{m}'] = R['sample']['mode'].get(m, 0)
    K['sample.adjacent'] = R['sample']['scope'].get('Transport-adjacent supply and technology', 0)
    K['sample.nolink'] = R['sample']['scope'].get('No transport link evident', 0)
    for it in R['items']:
        q = it['item']
        for k_, c in enumerate(it['counts'], 1): K[f'item.{q}.n{k_}'] = c
        for f in ('median', 'q1', 'q3', 'agree', 'agree_pct', 'ci_lo', 'ci_hi', 'direct_agree', 'direct_agree_pct', 'direct_median'):
            K[f'item.{q}.{f}'] = it[f]
    cs = R['corr_summary']
    K.update({'corr.pairs': cs['pairs'], 'corr.raw_lt_05': cs['raw'], 'corr.holm_lt_05': cs['holm'], 'corr.bh_lt_05': cs['bh']})
    for p in R['pairs']:
        if p['bh'] < ALPHA:
            for f in ('rho', 'p', 'holm', 'bh', 'ci_lo', 'ci_hi'): K[f"corr.{p['a']}_{p['b']}.{f}"] = p[f]
    for r in R['kw']:
        for f in ('H', 'eta2', 'p', 'holm', 'bh'): K[f"size.{r['item']}.{f}"] = r[f]
    for r in R['mw']:
        for f in ('U', 'r', 'p', 'holm', 'bh'): K[f"startup.{r['item']}.{f}"] = r[f]
    K['groups.size_raw_lt_05'] = sum(r['p'] < ALPHA for r in R['kw']); K['groups.size_holm_lt_05'] = sum(r['holm'] < ALPHA for r in R['kw'])
    K['groups.startup_raw_lt_05'] = sum(r['p'] < ALPHA for r in R['mw']); K['groups.startup_holm_lt_05'] = sum(r['holm'] < ALPHA for r in R['mw'])
    tot = defaultdict(int)
    for q, st in R['status'].items():
        for f, v in st.items():
            K[f'qual.{q}.{f}'] = v; tot[f] += v
    for f, v in tot.items(): K[f'qual.total.{f}'] = v
    for q in ('Q34', 'Q36', 'Q37'):
        for code in TABLE3_CODES:
            K[f'code.{q}.{code}.n'] = 0; K[f'code.{q}.{code}.n_direct'] = 0
    for c in R['codes']:
        K[f"code.{c['question']}.{c['code']}.n"] = c['n']; K[f"code.{c['question']}.{c['code']}.n_direct"] = c['n_direct']
    for c in R['contrast']:
        for f in ('agree', 'q34', 'q37', 'either', 'agree_and_either', 'agree_d', 'q34_d', 'q37_d', 'either_d', 'agree_and_either_d'):
            K[f"contrast.{c['item']}.{f}"] = c[f]
    for name, d_ in R['collab_sensitivity'].items():
        for f, v in d_.items(): K[f'sens.collab_{name}.{f}'] = v
    for f, v in R['barriers'].items():
        if f == 'distribution':
            for k_, c in enumerate(v): K[f'barrier.dist_{k_}'] = c
        else: K[f'barrier.{f}'] = v
    for f, v in R['text'].items(): K[f'text.{f}'] = v
    K['text.response_words_median'] = R['response_words']['median']
    K['text.response_words_q1'] = R['response_words']['q1']; K['text.response_words_q3'] = R['response_words']['q3']
    return {k: (float(v) if isinstance(v, (np.floating, float)) else int(v)) for k, v in K.items()}


# ----------------------------------------------------------------------------- tables
def write_csv(path: Path, header, rows):
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f); w.writerow(header); w.writerows(rows)


def write_tables(R, out: Path, labels):
    out.mkdir(exist_ok=True)
    s = R['sample']
    rows = [('Employees', k, v) for k, v in s['size'].items()] + [('Years in operation', k, v) for k, v in s['age'].items()] + \
           [('Start-up', k, v) for k, v in s['startup'].items()] + [('Transport activity', k, v) for k, v in s['scope'].items()] + \
           [('Direct transport mode', k, v) for k, v in s['mode'].items()]
    write_csv(out / 'table1_sample.csv', ['characteristic', 'category', 'n', 'percent'], [(a, b, c, round(100 * c / R['n'], 1)) for a, b, c in rows])
    write_csv(out / 'tableS2_items.csv',
              ['item', 'label', 'n1', 'n2', 'n3', 'n4', 'n5', 'median', 'q1', 'q3', 'agree_n', 'agree_pct', 'wilson_lo', 'wilson_hi', 'direct_agree_n', 'direct_agree_pct', 'direct_median'],
              [[i['item'], i['label'], *i['counts'], i['median'], i['q1'], i['q3'], i['agree'], round(i['agree_pct'], 1), round(i['ci_lo'], 1),
                round(i['ci_hi'], 1), i['direct_agree'], round(i['direct_agree_pct'], 1), i['direct_median']] for i in R['items']])
    write_csv(out / 'spearman_351_pairs.csv', ['item_a', 'item_b', 'rho', 'p', 'holm_p', 'bh_p', 'boot_ci_lo', 'boot_ci_hi'],
              [[p['a'], p['b'], round(p['rho'], 6), p['p'], p['holm'], p['bh'], p.get('ci_lo', ''), p.get('ci_hi', '')] for p in R['pairs']])
    write_csv(out / 'table2_correlations_bh.csv', ['item_a', 'item_b', 'rho', 'boot_ci_lo', 'boot_ci_hi', 'p', 'holm_p', 'bh_p'],
              [[p['a'], p['b'], round(p['rho'], 2), round(p['ci_lo'], 2), round(p['ci_hi'], 2), p['p'], p['holm'], p['bh']]
               for p in sorted(R['pairs'], key=lambda p: p['p']) if p['bh'] < ALPHA])
    write_csv(out / 'tableS4_group_comparisons.csv',
              ['item', 'KW_H', 'eta2', 'KW_p', 'KW_holm', 'KW_bh', 'MW_U', 'rank_biserial', 'MW_p', 'MW_holm', 'MW_bh'],
              [[a['item'], round(a['H'], 3), round(a['eta2'], 3), a['p'], a['holm'], a['bh'], b['U'], round(b['r'], 3), b['p'], b['holm'], b['bh']]
               for a, b in zip(R['kw'], R['mw'])])
    write_csv(out / 'tableS5_response_status.csv', ['question', 'responses', 'no_response', 'responses_direct', 'codes_assigned', 'multi_coded'],
              [[q, st['responses'], st['no_response'], st['responses_direct'], st['codes'], st['multi']] for q, st in R['status'].items()])
    write_csv(out / 'tableS7_code_frequencies.csv', ['question', 'code', 'label', 'n', 'percent', 'n_direct', 'percent_direct'],
              [[c['question'], c['code'], labels.get(c['code'], ''), c['n'], round(100 * c['n'] / R['n'], 1), c['n_direct'],
                round(100 * c['n_direct'] / R['n_direct'], 1)] for c in R['codes']])
    write_csv(out / 'tableS8_closed_vs_open.csv',
              ['support', 'closed_item', 'code', 'agreed', 'q34', 'q37', 'q34_or_q37', 'agreed_and_mentioned', 'agreed_direct', 'q34_direct', 'q37_direct',
               'q34_or_q37_direct', 'agreed_and_mentioned_direct'],
              [[c['name'], c['item'], c['code'], c['agree'], c['q34'], c['q37'], c['either'], c['agree_and_either'], c['agree_d'], c['q34_d'], c['q37_d'],
                c['either_d'], c['agree_and_either_d']] for c in R['contrast']])
    write_csv(out / 'figure1_data.csv', ['item', 'label', 'pct_1', 'pct_2', 'pct_3', 'pct_4', 'pct_5', 'agree_n'],
              [[i['item'], i['label'], *[round(100 * c / R['n'], 2) for c in i['counts']], i['agree']] for i in R['items']])
    b = R['barriers']
    write_csv(out / 'tableS3_barrier_cooccurrence.csv', ['measure', 'respondents'],
              [['Agreed with at least one of Q10, Q12, Q15', b['cost_any']], ['Agreed with two or more of Q10, Q12, Q15', b['cost_two_plus']],
               ['Agreed with all of Q10, Q12, Q15', b['cost_all']], ['At least one of Q10, Q12, Q15 and at least one of Q11, Q13, Q14, Q16, Q17', b['cost_and_other']],
               ['At least one of Q11, Q13, Q14, Q16, Q17 but none of Q10, Q12, Q15', b['other_only']], ['None of the eight barrier statements', b['none']],
               ['Two or more of the eight', b['two_plus']], ['Three or more of the eight', b['three_plus']],
               ['Median number of the eight agreed', b['median']], ['First quartile', b['q1']], ['Third quartile', b['q3']],
               ['Agreed with all five support statements', b['support_all_five']]] + [[f'Agreed with exactly {k_} of the eight', c] for k_, c in enumerate(b['distribution'])])
    cn = {(c['question'], c['code']): c['n'] for c in R['codes']}
    write_csv(out / 'figure2_data.csv', ['theme', 'code', 'challenge_q36', 'support_needed_q34', 'suggestion_q37'],
              [[lab, code, cn.get(('Q36', code), 0), cn.get(('Q34', code), 0), cn.get(('Q37', code), 0)] for lab, code in zip(FIGURE2_LABELS, FIGURE2_CODES)])
    write_csv(out / 'figureS2_data.csv', ['support', 'closed_item', 'pct_agreed', 'pct_q34', 'pct_q37'],
              [[c['name'], c['item'], round(100 * c['agree'] / R['n'], 2), round(100 * c['q34'] / R['n'], 2), round(100 * c['q37'] / R['n'], 2)] for c in R['contrast']])
    M = np.corrcoef(stats.rankdata(R['_X'], axis=0), rowvar=False)
    write_csv(out / 'figureS1_matrix.csv', ['item'] + ITEMS, [[q] + [round(float(v), 4) for v in M[j]] for j, q in enumerate(ITEMS)])


# ----------------------------------------------------------------------------- figures
INK, MUTED, GRID = '#1a1a1a', '#5f5f5f', '#d9d9d9'
LIKERT_COLOURS = ['#b5541a', '#eab07c', '#d2d2d2', '#8db6dc', '#1f5f9e']


def make_figures(R, figdir: Path, dpi=1000):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
    from matplotlib.lines import Line2D
    from matplotlib import font_manager
    figdir.mkdir(exist_ok=True)
    have_arial = any(f.name == 'Arial' for f in font_manager.fontManager.ttflist)
    plt.rcParams.update({'font.family': 'Arial' if have_arial else 'DejaVu Sans', 'font.size': 8, 'axes.linewidth': 0.6})
    n = R['n']
    save = lambda fig, name, d=dpi: (fig.savefig(figdir / f'{name}.tif', dpi=d, facecolor='white', pil_kwargs={'compression': 'tiff_lzw'}), plt.close(fig))
    items = {i['item']: i for i in R['items']}

    # ---- Figure 1: all 27 items, bars centred on the neutral category
    fig = plt.figure(figsize=(7.48, 8.3)); ax = fig.add_axes([0.385, 0.078, 0.50, 0.905])
    y = 0; ticks, ticklabels, heads, extent = [], [], [], 0.0
    for gname, rng in ITEM_GROUPS:
        heads.append((y, gname)); y += 1
        for k in rng:
            it = items[f'Q{k}']; pct = np.array(it['counts']) / n * 100
            left = -(pct[0] + pct[1] + pct[2] / 2); extent = max(extent, abs(left), left + pct.sum())
            for j in range(5):
                ax.barh(y, pct[j], left=left, height=0.68, color=LIKERT_COLOURS[j], edgecolor='white', linewidth=0.5); left += pct[j]
            ticks.append(y); ticklabels.append(f"{it['item']}  {it['label']}")
            assert it['agree'] == it['counts'][3] + it['counts'][4]          # label is the agreement count
            ax.text(104, y, f"{it['agree']}/{n}", va='center', ha='left', fontsize=7, color=INK)
            y += 1
        y += 0.35
    assert extent <= 100, 'a bar would extend beyond the axis'                # no bar is clipped
    ax.set_ylim(y - 0.6, -0.4); ax.set_xlim(-100, 100)
    ax.set_yticks(ticks); ax.set_yticklabels(ticklabels, fontsize=7.2, color=INK)
    for yy, g in heads:
        ax.text(-0.765, yy + 0.05, g, transform=ax.get_yaxis_transform(), fontsize=8, fontweight='bold', va='center', ha='left', color=INK)
    ax.axvline(0, color=MUTED, lw=0.6)
    ax.set_xticks([-100, -50, 0, 50, 100]); ax.set_xticklabels(['100', '50', '0', '50', '100'], fontsize=7)
    ax.set_xlabel(f'Percentage of respondents (n = {n}); bars centred on the neutral category', fontsize=7.5, color=INK)
    ax.xaxis.grid(True, color=GRID, lw=0.4); ax.set_axisbelow(True)
    for sp in ('top', 'right', 'left'): ax.spines[sp].set_visible(False)
    ax.tick_params(axis='y', length=0); ax.tick_params(axis='x', length=2, width=0.5)
    ax.text(104, 0.0, 'Agree', fontsize=7, fontweight='bold', ha='left', va='center', color=INK)
    fig.legend([plt.Rectangle((0, 0), 1, 1, color=c) for c in LIKERT_COLOURS],
               ['Completely disagree', 'Disagree', 'Neutral', 'Agree', 'Completely agree'], loc='lower center', ncol=5, frameon=False,
               fontsize=7.2, bbox_to_anchor=(0.5, 0.004), handlelength=1.1, handleheight=0.9, columnspacing=1.4)
    save(fig, 'Figure_1')

    # ---- Figure 2: themes in the open-text answers (challenges, support needed, suggestions)
    cn = {(c['question'], c['code']): c['n'] for c in R['codes']}
    series = [('Q36', 'Challenge faced (Q36)', '#1f5f9e'), ('Q34', 'Additional support needed (Q34)', '#e39a52'), ('Q37', 'Further suggestion (Q37)', '#7a7a7a')]
    fig, ax = plt.subplots(figsize=(7.0, 4.6)); fig.subplots_adjust(left=0.33, right=0.97, top=0.90, bottom=0.11)
    hbar = 0.26
    for i, code in enumerate(FIGURE2_CODES):
        for k_, (q, lab, colr) in enumerate(series):
            v = cn.get((q, code), 0); yb = i + (k_ - 1) * hbar
            ax.barh(yb, v, height=hbar * 0.92, color=colr, edgecolor='white', linewidth=0.4, label=lab if i == 0 else None)
            ax.text(v + 0.35, yb, str(v), va='center', ha='left', fontsize=6.8, color=INK)
    top = max(cn.get((q, c), 0) for q, _, _ in series for c in FIGURE2_CODES)
    ax.set_yticks(range(len(FIGURE2_CODES))); ax.set_yticklabels(FIGURE2_LABELS, fontsize=8, color=INK)
    ax.set_ylim(len(FIGURE2_CODES) - 0.45, -0.55); ax.set_xlim(0, top + 3)
    ax.set_xlabel(f'Number of respondents whose answer was assigned to the code (n = {n})', fontsize=8, color=INK)
    ax.xaxis.grid(True, color=GRID, lw=0.4); ax.set_axisbelow(True)
    for sp in ('top', 'right', 'left'): ax.spines[sp].set_visible(False)
    ax.tick_params(axis='y', length=0); ax.tick_params(axis='x', length=2, width=0.5)
    ax.legend(loc='lower right', frameon=False, fontsize=7.5, handlelength=1.1, handleheight=0.9)
    save(fig, 'Figure_2')

    # ---- Figure S2: closed-item agreement and later open-text mention (descriptive comparison)
    con = R['contrast']; CP, C34, C37 = '#1f5f9e', '#b5541a', '#e39a52'
    fig, ax = plt.subplots(figsize=(7.0, 3.3)); fig.subplots_adjust(left=0.30, right=0.97, top=0.80, bottom=0.17)
    for i, c in enumerate(con):
        p, a, b = 100 * c['agree'] / n, 100 * c['q34'] / n, 100 * c['q37'] / n
        ax.plot([min(a, b), p], [i, i], color='#bdbdbd', lw=1.2, zorder=1, solid_capstyle='round')
        ax.scatter(p, i, s=70, color=CP, zorder=3, edgecolor='white', linewidth=0.8)
        ax.scatter(a, i, s=58, marker='D', color=C34, zorder=3, edgecolor='white', linewidth=0.8)
        ax.scatter(b, i, s=62, marker='s', color=C37, zorder=3, edgecolor='white', linewidth=0.8)
        ax.text(p + 2.2, i, f"{c['agree']}/{n}", va='center', ha='left', fontsize=7.5, color=INK)
        if abs(a - b) < 9:
            ax.text(max(a, b) + 2.2, i - 0.30, f"Q34: {c['q34']}/{n}   Q37: {c['q37']}/{n}", va='center', ha='left', fontsize=7.5, color=INK)
        else:
            ax.text(a, i - 0.30, f"Q34: {c['q34']}/{n}", va='center', ha='center', fontsize=7.5, color=INK)
            ax.text(b, i - 0.30, f"Q37: {c['q37']}/{n}", va='center', ha='center', fontsize=7.5, color=INK)
    ax.set_yticks(range(len(con))); ax.set_yticklabels([f"{c['name']}\n(closed item {c['item']})" for c in con], fontsize=8, color=INK)
    ax.set_ylim(len(con) - 0.45, -0.6); ax.set_xlim(0, 100)
    ax.set_xlabel(f'Percentage of all respondents (n = {n})', fontsize=8, color=INK)
    ax.xaxis.grid(True, color=GRID, lw=0.4); ax.set_axisbelow(True)
    for sp in ('top', 'right', 'left'): ax.spines[sp].set_visible(False)
    ax.tick_params(axis='y', length=0); ax.tick_params(axis='x', length=2, width=0.5)
    handles = [Line2D([], [], marker='o', color='none', markerfacecolor=CP, markeredgecolor='white', markersize=8, label='Agreed with the closed statement'),
               Line2D([], [], marker='D', color='none', markerfacecolor=C34, markeredgecolor='white', markersize=7, label='Mentioned when asked what additional support would help (Q34)'),
               Line2D([], [], marker='s', color='none', markerfacecolor=C37, markeredgecolor='white', markersize=7, label='Mentioned when asked for further suggestions (Q37)')]
    fig.legend(handles=handles, loc='upper left', bbox_to_anchor=(0.02, 1.0), ncol=1, frameon=False, fontsize=7.5, handletextpad=0.3, labelspacing=0.35)
    save(fig, 'Figure_S2', 600)

    # ---- Figure 3: conceptual proposition (no data)
    fig = plt.figure(figsize=(7.2, 4.0)); ax = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, 100); ax.set_ylim(0, 55.5); ax.axis('off')
    def box(x, y0, w, h, title, lines, fc, ec):
        ax.add_patch(FancyBboxPatch((x, y0), w, h, boxstyle='round,pad=0,rounding_size=1.2', fc=fc, ec=ec, lw=1.0))
        ax.text(x + w / 2, y0 + h - 3.0, title, ha='center', va='center', fontsize=8.6, fontweight='bold', color=INK)
        for k_, t in enumerate(lines):
            ax.text(x + w / 2, y0 + h - 7.3 - 3.75 * k_, t, ha='center', va='center', fontsize=7.3, color=INK)
    bw, bh, ytop = 27, 22.5, 31
    box(3, ytop, bw, bh, 'Technological conditions', ['Feasibility and performance in use', 'Testing and validation', 'Operational fit', 'Emissions data and measurement'], '#eaf1f8', '#1f5f9e')
    box(36.5, ytop, bw, bh, 'Organisational conditions', ['Finance: amount, form and timing', 'Staff time and specialist skills', 'Premises and tenure', 'Routines and decision authority'], '#f3f3f3', '#5f5f5f')
    box(70, ytop, bw, bh, 'External conditions', ['Policy clarity and stability', 'Customer demand and procurement', 'Energy and infrastructure access', 'Suppliers and intermediaries'], '#fbefe3', '#b5541a')
    ax.add_patch(FancyBboxPatch((20, 13.2), 60, 9.6, boxstyle='round,pad=0,rounding_size=1.2', fc='white', ec=INK, lw=1.1))
    ax.text(50, 19.9, 'Joint configuration of conditions for a given measure', ha='center', va='center', fontsize=8.6, fontweight='bold', color=INK)
    ax.text(50, 15.9, 'Proposition: a shortfall in one domain can stall implementation when the others are favourable', ha='center', va='center', fontsize=7.1, color=INK)
    for x in (16.5, 50, 83.5):
        ax.add_patch(FancyArrowPatch((x, ytop - 0.4), (50 + (x - 50) * 0.55, 23.1), arrowstyle='-|>', mutation_scale=9, lw=0.9, color=MUTED, shrinkA=0, shrinkB=0))
    ax.add_patch(FancyBboxPatch((27, 1.4), 46, 6.8, boxstyle='round,pad=0,rounding_size=1.2', fc='white', ec=MUTED, lw=1.0, ls=(0, (4, 2.5))))
    ax.text(50, 6.0, 'Verified implementation outcomes', ha='center', va='center', fontsize=8.4, fontweight='bold', color=INK)
    ax.text(50, 3.2, 'Adoption, sustained use, emissions (not measured in this study)', ha='center', va='center', fontsize=7.1, color=MUTED)
    ax.add_patch(FancyArrowPatch((50, 12.9), (50, 8.5), arrowstyle='-|>', mutation_scale=9, lw=0.9, color=MUTED, ls=(0, (4, 2.5)), shrinkA=0, shrinkB=0))
    ax.text(52, 10.7, 'to be tested prospectively', ha='left', va='center', fontsize=7.1, style='italic', color=MUTED)
    for x in (31.5, 65):
        ax.annotate('', xy=(x + 4, ytop + bh / 2), xytext=(x - 0.6, ytop + bh / 2), arrowprops=dict(arrowstyle='<->', lw=0.8, color=MUTED, shrinkA=0, shrinkB=0))
    save(fig, 'Figure_3')

    # ---- Figure S1: Spearman matrix with adjusted significance
    lookup = {(p['a'], p['b']): p for p in R['pairs']}
    M = np.eye(len(ITEMS))
    for (a, b), p in lookup.items():
        M[ITEMS.index(b), ITEMS.index(a)] = p['rho']
    fig, ax = plt.subplots(figsize=(7.3, 6.7)); fig.subplots_adjust(left=0.07, right=0.99, top=0.985, bottom=0.07)
    im = ax.imshow(np.ma.array(M, mask=np.triu(np.ones_like(M, dtype=bool))), cmap='RdBu_r', vmin=-1, vmax=1)
    for i in range(len(ITEMS)):
        for j in range(i):
            p = lookup[(ITEMS[j], ITEMS[i])]; v = p['rho']
            mark = '**' if p['holm'] < ALPHA else '*' if p['bh'] < ALPHA else ''
            txt = '.00' if abs(v) < 0.005 else f'{v:.2f}'.replace('0.', '.').replace('-.', '−.')
            colour = 'white' if abs(v) > 0.6 else INK
            ax.text(j, i - (0.12 if mark else 0), txt, ha='center', va='center', fontsize=4.6, color=colour)
            if mark: ax.text(j, i + 0.27, mark, ha='center', va='center', fontsize=5.5, fontweight='bold', color=colour)
    ax.set_xticks(range(len(ITEMS) - 1)); ax.set_xticklabels(ITEMS[:-1], fontsize=6, rotation=90)
    ax.set_yticks(range(1, len(ITEMS))); ax.set_yticklabels(ITEMS[1:], fontsize=6)
    ax.set_xlim(-0.5, len(ITEMS) - 1.5); ax.set_ylim(len(ITEMS) - 0.5, 0.5)
    for sp in ax.spines.values(): sp.set_visible(False)
    ax.tick_params(length=0)
    cax = fig.add_axes([0.62, 0.80, 0.30, 0.02]); cb = fig.colorbar(im, cax=cax, orientation='horizontal')
    cb.set_label("Spearman's ρ", fontsize=7); cb.ax.tick_params(labelsize=6, length=2); cb.outline.set_linewidth(0.4)
    fig.text(0.62, 0.74, '*  Benjamini–Hochberg adjusted p < .05\n** Holm adjusted p < .05 (351 pairs)', fontsize=6.8, color=INK, va='top')
    save(fig, 'Figure_S1', 600)


# ----------------------------------------------------------------------------- workbook interaction
def read_codebook_labels(path: Path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True); labels = {}
    if 'Codebook' in wb.sheetnames:
        for row in wb['Codebook'].iter_rows(min_row=2, max_col=2, values_only=True):
            if row[0]: labels[str(row[0])] = str(row[1])
    wb.close(); return labels


def read_validation(path: Path):
    """Return rows of the Validation sheet: key, reported value, tolerance and the cached Excel value (None if not yet calculated)."""
    typed = openpyxl.load_workbook(path, read_only=True, data_only=False)
    cached = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if 'Validation' not in typed.sheetnames:
        typed.close(); cached.close(); return []
    rows = []
    for a, b in zip(typed['Validation'].iter_rows(min_row=2, max_col=7, values_only=True),
                    cached['Validation'].iter_rows(min_row=2, max_col=7, values_only=True)):
        if a[0] in (None, ''): continue
        has_formula = isinstance(a[3], str) and a[3].startswith('=')
        excel = b[3] if isinstance(b[3], (int, float)) else None
        rows.append({'key': a[0], 'reported': a[2], 'excel': excel, 'has_formula': has_formula,
                     'tol': a[5] if isinstance(a[5], (int, float)) else 0})
    typed.close(); cached.close(); return rows


def update_workbook(path: Path, R, K):
    """Refresh the two Python-derived sheets. Formula sheets are left untouched and recalculate when the file is opened in Excel."""
    wb = openpyxl.load_workbook(path)
    for name in ('PY_Bootstrap', 'PY_Check'):
        if name in wb.sheetnames:
            ws = wb[name]; ws.delete_rows(1, ws.max_row + 1)
        else:
            ws = wb.create_sheet(name)
    ws = wb['PY_Bootstrap']
    ws.append(['PYTHON-DERIVED. Percentile bootstrap intervals for Spearman correlations significant after Benjamini-Hochberg adjustment. '
               'Not recalculated by Excel: rerun reproduce.py --update-workbook after any change to AT1.'])
    ws.append([f'Resamples: {BOOTSTRAP_RESAMPLES}; base seed: {BOOTSTRAP_SEED}; the generator for each pair is seeded with [base seed, index of item a, index of item b].'])
    ws.append(['item_a', 'item_b', 'rho', 'ci_lo_95', 'ci_hi_95', 'resamples_used'])
    for p in sorted(R['pairs'], key=lambda p: p['p']):
        if p['bh'] < ALPHA:
            ws.append([p['a'], p['b'], p['rho'], p['ci_lo'], p['ci_hi'], p['boot_n']])
    ws = wb['PY_Check']
    ws.append(['key', 'python_value', 'PYTHON-DERIVED. Every value recomputed by reproduce.py from AT1 and the two registers; the Validation sheet looks these up.'])
    for k in sorted(K):
        ws.append([k, K[k]])
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def validate(K, rows, log):
    bad = 0; n_rep = n_xl = 0; missing_excel = python_only = 0
    for r in rows:
        key = r['key']
        if key not in K:
            log.append(f'UNKNOWN KEY in Validation sheet: {key}'); bad += 1; continue
        if isinstance(r['reported'], (int, float)):
            n_rep += 1
            if abs(K[key] - r['reported']) > r['tol'] + 1e-9:
                bad += 1; log.append(f'MISMATCH python vs reported: {key}: python {K[key]:.6g}, reported {r["reported"]}')
        if r['excel'] is None:
            missing_excel += 1 if r['has_formula'] else 0
            python_only += 0 if r['has_formula'] else 1
        else:
            n_xl += 1
            if abs(K[key] - r['excel']) > 1e-6 * max(1, abs(K[key])) + 1e-9:
                bad += 1; log.append(f'MISMATCH python vs Excel formula: {key}: python {K[key]:.10g}, Excel {r["excel"]:.10g}')
    log.append(f'Checked {n_rep} recomputed values against the values reported in the manuscript and {n_xl} against Excel formula results; mismatches: {bad}.')
    log.append(f'{python_only} values are calculated in Python only (bootstrap intervals, response length and the coding sensitivity check).')
    if missing_excel:
        log.append(f'{missing_excel} Excel formulas have no stored result: open and save the workbook in Excel, then run this script again.')
    return bad


def main():
    ap = argparse.ArgumentParser(description='Reproduce the manuscript analyses and figures.')
    ap.add_argument('--update-workbook', action='store_true', help='refresh the Python-derived sheets in Analysis.xlsx')
    ap.add_argument('--no-figures', action='store_true', help='skip figure generation')
    ap.add_argument('--dpi', type=int, default=1000, help='resolution of the TIFF figures (default 1000)')
    args = ap.parse_args()

    data = load_at1(WORKBOOK)
    load_scope(SCOPE_CSV, data)
    assign, noresp = load_coding(CODING_CSV, data)
    R = analyse(data, assign, noresp)
    K = key_values(R)

    out = HERE / 'outputs'
    write_tables(R, out, read_codebook_labels(WORKBOOK))
    if not args.no_figures:
        make_figures(R, HERE / 'figures', args.dpi)
    if args.update_workbook:
        update_workbook(WORKBOOK, R, K)

    cs = R['corr_summary']
    log = [f'Respondents: {R["n"]}; direct transport: {R["n_direct"]}.',
           f'Spearman pairs: {cs["pairs"]}; p < 0.05 unadjusted: {cs["raw"]}; Benjamini-Hochberg: {cs["bh"]}; Holm: {cs["holm"]}.',
           f'Group comparisons with Holm-adjusted p < 0.05: size {K["groups.size_holm_lt_05"]}, start-up {K["groups.startup_holm_lt_05"]}.',
           f'Open-text responses analysed: {K["qual.total.responses"]}; non-response cells: {K["qual.total.no_response"]}; codes assigned: {K["qual.total.codes"]}.']
    for c in R['contrast']:
        log.append(f'{c["name"]}: agreed {c["agree"]}, Q34 {c["q34"]}, Q37 {c["q37"]}, either {c["either"]}, agreed and mentioned {c["agree_and_either"]}.')
    bad = validate(K, read_validation(WORKBOOK), log)
    log.append('VALIDATION PASSED' if bad == 0 else f'VALIDATION FAILED ({bad} mismatches)')
    (out / 'validation_log.txt').write_text('\n'.join(log) + '\n', encoding='utf-8')
    print('\n'.join(log))
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
