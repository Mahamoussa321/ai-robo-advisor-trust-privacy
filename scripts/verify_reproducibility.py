import os, csv, json, hashlib, pathlib, sys
ROOT = pathlib.Path(os.environ.get('ROBO_PROJECT_ROOT', pathlib.Path(__file__).resolve().parents[1]))
expected = [
    'outputs/analysis_summary.json',
    'outputs/study2_construct_reliability.csv',
    'outputs/study2_ols_wu.csv',
    'outputs/study2_high_wu_ml_performance.csv',
    'outputs/study2_actual_use_ml_performance.csv',
    'outputs/study3_nfcs_weighted_indicators.csv',
    'outputs/study3_digital_trading_ml_performance.csv',
    'outputs/figure1_conceptual_model.png',
    'outputs/figure5_study2_ml_roc_high_wu.png',
    'outputs/figure11_ml_performance_heatmap.png',
    'manuscript/robo_advisor_final_submission_ready.docx',
]
missing=[p for p in expected if not (ROOT/p).exists()]
if missing:
    print('Missing expected outputs:')
    print('\n'.join(missing))
    sys.exit(1)
with open(ROOT/'outputs'/'analysis_summary.json', encoding='utf-8') as f:
    s=json.load(f)
checks = {
    'study2_n_complete': s.get('study2_n_complete'),
    'study3_n': s.get('study3_n'),
    'study2_use_rate': round(float(s.get('study2_use_rate')), 4),
}
print('Reproducibility check passed.')
print(json.dumps(checks, indent=2))
