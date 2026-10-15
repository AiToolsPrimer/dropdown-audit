# dropdown-audit

Find the drop-down lists in an `.xlsx` file that look like rules but accept anything typed into them.

A drop-down in Excel is a Data Validation rule of type `list`. Whether it refuses a value that is not on the list depends on two attributes of `<dataValidation>` that nothing in the menu prompts you to check:

- **`showErrorMessage`.** With it off, no alert appears and whatever was typed stays in the cell. Microsoft's VBA reference describes the same switch, `Validation.ShowError`, as "True if the data validation error message will be displayed whenever the user enters invalid data".
- **`errorStyle`**: `stop`, `warning` or `information`. Only Stop refuses the entry. Microsoft: "To stop people from entering data that isn't in the drop-down list, select Stop." When the attribute is missing, the schema default is `stop`.

## The openpyxl default

In openpyxl 3.1.5 (checked on 15 October 2026) a new `DataValidation` starts with `showErrorMessage=False`. The list example in the openpyxl documentation sets an error message but never switches the alert on:

```python
dv = DataValidation(type="list", formula1='"Dog,Cat,Bat"', allow_blank=True)
dv.error = 'Your entry is not in the list'
```

So the message is never shown, and the list enforces nothing. This is what gets written:

```xml
<dataValidation sqref="A2:A100" showDropDown="0" showInputMessage="0" showErrorMessage="0" allowBlank="1" type="list">
```

The fix is one argument, plus the style written out so nobody has to remember the default:

```python
dv = DataValidation(type="list", formula1='"Open,Closed,Waiting"', allow_blank=True,
                    showErrorMessage=True, errorStyle="stop",
                    errorTitle="Not on the list", error="Pick a value from the drop-down.")
```

One more attribute whose name reads backwards. In the openpyxl documentation's words: "Excel and LibreOffice interpret the parameter showDropDown=True as the dropdown arrow should be hidden."

## Usage

```bash
pip install openpyxl
python3 dropdown_audit.py workbook.xlsx
```

`make_enforced.py` writes a test file with one loose and one enforced list. The audit on that file:

```
NOT ENFORCED  Tickets!A2:A100  ("Open,Closed,Waiting")
    - showErrorMessage is off: no alert, any value is kept
ok            Tickets!B2:B100  ("Open,Closed,Waiting")
```

The exit code is 1 if any list does not enforce and 0 otherwise, so it can run as a check before a workbook is sent out.

## What it does not catch

- **Pasting.** Microsoft's "More on data validation": "When data is copied or filled, the messages do not appear." LibreOffice's help says the same of Calc: "if you insert a value in the cell either with drag-and-drop or by copying and pasting, the validity rule will not take effect."
- **Other rule types.** Whole number, decimal, date, length and custom rules are skipped; only lists are checked.

## Sources

- Microsoft Support: Create a drop-down list; More on data validation
- Microsoft Learn: Validation.ShowError property (Excel VBA reference)
- openpyxl documentation: Validating cells
- LibreOffice Help: Validity of cell contents (Calc)
- ECMA-376, SpreadsheetML: `ST_DataValidationErrorStyle`
- Explained in plain English: [How to make a drop down list in Excel, and make it hold](https://aitoolsprimer.com/docs-drop-down-list-excel.html) (AI Tools Primer)

## License

MIT
