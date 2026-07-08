from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from parser import parse_pdf, rows_to_csv

SAMPLE = Path('/Users/alfred/.hermes/cache/documents/doc_d90c4ea283ce_TO_ON_26_126_Job_Descriptions_WeekDAY_July_Control_Period_eff_July.pdf')


def main():
    rows = parse_pdf(str(SAMPLE))
    assert rows, 'expected rows to be extracted'
    pd_rows = [r for r in rows if r.record_type == 'PD time']
    stby_rows = [r for r in rows if r.record_type == 'STBY']
    split_rows = [r for r in rows if r.record_type == 'Split']

    assert len(pd_rows) >= 1000, f'expected many PD time rows, got {len(pd_rows)}'
    assert len(stby_rows) >= 1, 'expected at least one STBY row'
    assert len(split_rows) >= 1, 'expected at least one Split row'

    first = rows[0]
    assert first.job_number == '10001', f'unexpected first job {first.job_number}'
    assert pd_rows[0].start and pd_rows[0].end and pd_rows[0].duration
    assert rows_to_csv(rows).startswith('job_number,valid_days,page')

    print('OK')
    print(f'total records: {len(rows)}')
    print(f'PD time: {len(pd_rows)}')
    print(f'STBY: {len(stby_rows)}')
    print(f'Split: {len(split_rows)}')
    print(f'jobs with extracted records: {len(set(r.job_number for r in rows))}')


if __name__ == '__main__':
    main()
