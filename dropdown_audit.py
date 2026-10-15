#!/usr/bin/env python3
"""List the drop-down lists in an .xlsx file that do not actually reject other values.

A list rule only refuses an entry when two attributes agree: showErrorMessage is on,
and the error style is Stop (Stop is also what Excel assumes when no style is written).
With showErrorMessage off, or with the style set to Warning or Information, anything
typed into the cell is accepted.

Usage:  python3 dropdown_audit.py workbook.xlsx [more.xlsx ...]
Exit status is 1 if any list rule does not enforce, so it can run in CI.
"""
import sys
import openpyxl


def check(path):
    problems = 0
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            if dv.type != 'list':
                continue
            style = dv.errorStyle or 'stop'
            reasons = []
            if not dv.showErrorMessage:
                reasons.append('showErrorMessage is off: no alert, any value is kept')
            if style != 'stop':
                reasons.append('error style is %s: the user can click past the alert' % style)
            if dv.showDropDown:
                # The attribute name reads backwards: True hides the in-cell arrow.
                reasons.append('showDropDown is True, which hides the arrow')
            where = '%s!%s' % (ws.title, dv.sqref)
            if reasons:
                problems += 1
                print('NOT ENFORCED  %s  (%s)' % (where, dv.formula1))
                for r in reasons:
                    print('    - ' + r)
            else:
                print('ok            %s  (%s)' % (where, dv.formula1))
    return problems


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    total = sum(check(p) for p in sys.argv[1:])
    sys.exit(1 if total else 0)
