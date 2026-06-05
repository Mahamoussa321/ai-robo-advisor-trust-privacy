"""Run the complete reproducible workflow for the robo-advisor paper.

Usage:
    python scripts/run_all.py

It will regenerate all outputs/ tables + figures, then rebuild manuscript/robo_advisor_final_submission_ready.docx.
"""
import os, sys, subprocess, pathlib, datetime
ROOT = pathlib.Path(__file__).resolve().parents[1]
env = os.environ.copy()
env['ROBO_PROJECT_ROOT'] = str(ROOT)
logdir = ROOT / 'logs'
logdir.mkdir(exist_ok=True)
log = logdir / ('run_all_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.log')
steps = [
    [sys.executable, str(ROOT / 'scripts' / 'final_robo_advisor_analysis.py')],
    [sys.executable, str(ROOT / 'scripts' / 'create_final_submission_docx.py')],
    [sys.executable, str(ROOT / 'scripts' / 'verify_reproducibility.py')],
]
with open(log, 'w', encoding='utf-8') as f:
    for cmd in steps:
        print('RUN:', ' '.join(cmd))
        f.write('RUN: ' + ' '.join(cmd) + '\n')
        proc = subprocess.run(cmd, cwd=str(ROOT), env=env, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        f.write(proc.stdout + '\n')
        print(proc.stdout)
        if proc.returncode != 0:
            raise SystemExit(f'Step failed with code {proc.returncode}: {cmd}. See {log}')
print(f'Complete. Log saved to {log}')
