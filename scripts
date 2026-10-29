#!/usr/bin/env python3
"""DZIF Excel -> validated JSON. Never silently treat unstructured ranges as tables."""
from __future__ import annotations
import datetime as dt
import hashlib
import json
import os
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

SOURCE = Path(os.environ.get('DZIF_MASTER_XLSX', 'DZIF_Intelligence_Master.xlsx'))
TARGET = Path('DZIF_Intelligence_Master.json')
REPORT = Path('DZIF_Master_Validation_Report.json')
STATUS = Path('DZIF_Master_Status.json')
GENERATOR = 'dzif-excel-export-1.0'


def atom(x):
    if x is None or isinstance(x, (str, int, bool)):
        return x
    if isinstance(x, float):
        if x != x or x in (float('inf'), float('-inf')):
            return None
        return x
    if isinstance(x, (dt.datetime, dt.date, dt.time)):
        return x.isoformat()
    if isinstance(x, dt.timedelta):
        return str(x)
    return str(x)


def save_json(p, v):
    p.write_text(json.dumps(v, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')


def run():
    if not SOURCE.is_file():
        raise FileNotFoundError(f'Excel-Datei nicht gefunden: {SOURCE}. Bitte exakt so benennen.')
    file_sha = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
    formulas_wb = load_workbook(SOURCE, data_only=False, read_only=False, keep_links=True)
    values_wb = load_workbook(SOURCE, data_only=True, read_only=False, keep_links=True)
    result = {
        'schema_version': '1.0',
        'generator_version': GENERATOR,
        'source': {'file': SOURCE.name, 'sha256': file_sha},
        'sheet_order': list(formulas_wb.sheetnames),
        'sheets': {},
    }
    report = {'source': SOURCE.name, 'sha256': file_sha, 'sheet_count': len(formulas_wb.sheetnames),
              'total_tables': 0, 'total_records': 0, 'total_cells': 0,
              'warnings': [], 'errors': [], 'sheets': [], 'ready_for_monitoring': False}

    for ws in formulas_wb.worksheets:
        val_ws = values_wb[ws.title]
        all_cells = []
        formulas = []
        links = []
        cells_count = 0
        nonempty_rows = set()
        for row in ws.iter_rows():
            for cell in row:
                cached = val_ws[cell.coordinate].value
                is_formula = cell.data_type == 'f'
                value = atom(cached if is_formula else cell.value)
                if cell.value is not None:
                    cells_count += 1
                    nonempty_rows.add(cell.row)
                    all_cells.append([cell.coordinate, value])
                if is_formula:
                    formulas.append({'cell': cell.coordinate, 'formula': str(cell.value), 'cached_value': value})
                    if cached is None:
                        report['warnings'].append(f'{ws.title}!{cell.coordinate}: Formel ohne gespeicherten Ergebniswert')
                if cell.hyperlink:
                    target = cell.hyperlink.target or cell.hyperlink.location
                    if target:
                        links.append({'cell': cell.coordinate, 'url': str(target)})

        sections = []
        for table_name in ws.tables:
            excel_table = ws.tables[table_name]
            min_col, min_row, max_col, max_row = range_boundaries(excel_table.ref)
            columns = [atom(val_ws.cell(min_row, col).value) for col in range(min_col, max_col + 1)]
            invalid = any(not isinstance(c, str) or not c.strip() for c in columns)
            duplicates = len([str(c).strip().lower() for c in columns]) != len(set(str(c).strip().lower() for c in columns))
            if invalid or duplicates:
                report['errors'].append(f'{ws.title}/{table_name}: leere oder doppelte Spaltenüberschrift')
                continue
            records = []
            for row_num in range(min_row + 1, max_row + 1):
                values = [atom(val_ws.cell(row_num, col).value) for col in range(min_col, max_col + 1)]
                if any(val is not None for val in values):
                    records.append({'excel_row': row_num, 'values': dict(zip(columns, values))})
            sections.append({'table_name': table_name, 'excel_range': excel_table.ref,
                             'columns': columns, 'record_count': len(records), 'records': records,
                             'extraction': 'excel_table'})
            report['total_tables'] += 1
            report['total_records'] += len(records)

        # No speculative header guessing: preserve cells, mark for human review.
        status = 'structured' if sections else ('raw_only' if len(nonempty_rows) >= 3 else 'notes_or_empty')
        if status == 'raw_only':
            report['warnings'].append(f'{ws.title}: befuelltes Blatt ohne Excel-Tabelle; Inhalte gespeichert, aber keine sicheren Datensaetze ableitbar')
        if len(sections) > 0:
            # Informational cells outside formal tables are still retained above.
            table_ranges = [range_boundaries(ws.tables[t].ref) for t in ws.tables]
            outside = 0
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    if not any(a <= cell.column <= c and b <= cell.row <= d for a, b, c, d in table_ranges):
                        outside += 1
            if outside > 25:
                report['warnings'].append(f'{ws.title}: {outside} weitere befuellte Zellen ausserhalb benannter Excel-Tabellen; in cells vorhanden, nicht als Records klassifiziert')

        result['sheets'][ws.title] = {
            'status': status, 'cell_count': cells_count,
            'tables': sections, 'cells': all_cells,
            'formulas': formulas, 'hyperlinks': links
        }
        report['total_cells'] += cells_count
        report['sheets'].append({'name': ws.title, 'status': status,
                                 'tables': len(sections), 'records': sum(t['record_count'] for t in sections),
                                 'cells': cells_count})

    if report['total_tables'] == 0:
        report['errors'].append('Keine definierten Excel-Tabellen gefunden. Bitte zuerst die Datenbereiche in Excel als Tabelle formatieren (Strg+T) oder einen auf die Arbeitsmappe angepassten Konverter verwenden.')
    # A sheet with a multirow data area but no Excel Table cannot be safely treated as structured.
    unstructured = [s['name'] for s in report['sheets'] if s['status'] == 'raw_only']
    if unstructured:
        report['errors'].append('Nicht strukturierte Datenbereiche: ' + ', '.join(unstructured))
    missing_caches = [x for x in report['warnings'] if 'ohne gespeicherten Ergebniswert' in x]
    if missing_caches:
        report['errors'].append(f'{len(missing_caches)} Formeln ohne gespeicherten Wert; Excel in Desktop-App neu berechnen und speichern.')
    report['ready_for_monitoring'] = len(report['errors']) == 0
    result['validation'] = {
        'ready_for_monitoring': report['ready_for_monitoring'],
        'total_sheets': report['sheet_count'], 'total_tables': report['total_tables'],
        'total_records': report['total_records'], 'total_cells': report['total_cells'],
        'warnings': report['warnings'][:50],
    }
    save_json(REPORT, report)
    save_json(STATUS, {
        'ready_for_monitoring': report['ready_for_monitoring'],
        'source_file': SOURCE.name,
        'source_sha256': file_sha,
        'checked_at_utc': dt.datetime.now(dt.timezone.utc).isoformat(),
        'total_sheets': report['sheet_count'],
        'total_tables': report['total_tables'],
        'total_records': report['total_records'],
        'errors': report['errors'][:20],
        'warnings_count': len(report['warnings'])
    })
    if report['ready_for_monitoring']:
        # Only publish a validated candidate. Never overwrite the last valid JSON on failure.
        if TARGET.exists():
            try:
                previous = json.loads(TARGET.read_text(encoding='utf-8'))
                if previous.get('source', {}).get('sha256') == file_sha and previous.get('generator_version') == GENERATOR:
                    print('Keine Aenderung an Excel und Konverter: validierte JSON bleibt unveraendert.')
                    return 0
            except (OSError, ValueError):
                pass
        result['source']['exported_at_utc'] = dt.datetime.now(dt.timezone.utc).isoformat()
        save_json(TARGET, result)
        print(f'OK: {len(formulas_wb.sheetnames)} Blaetter, {report["total_tables"]} Tabellen, {report["total_records"]} Datensaetze, {report["total_cells"]} Zellen.')
        return 0
    print('STOPP: Excel hat Bereiche, die ohne Strukturpruefung nicht sicher in Records umgewandelt werden koennen.')
    print('Letzte gueltige JSON wurde NICHT ueberschrieben. Bericht: ' + str(REPORT))
    for e in report['errors'][:30]:
        print('FEHLER:', e)
    return 2


if __name__ == '__main__':
    try:
        sys.exit(run())
    except Exception as exc:
        save_json(REPORT, {'ready_for_monitoring': False, 'errors': [f'{type(exc).__name__}: {exc}']})
        save_json(STATUS, {'ready_for_monitoring': False, 'source_file': SOURCE.name, 'source_sha256': None, 'errors': [f'{type(exc).__name__}: {exc}']})
        print(f'STOPP: {type(exc).__name__}: {exc}', file=sys.stderr)
        sys.exit(1)
