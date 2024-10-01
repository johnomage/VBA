# VBA

This repository contains resources and tools for working with Excel VBA.

## Packages

- xlwings
- watchgod
- XVBA VSCode extension (optional)

## Usage

### Editing VBA with xlwings

To edit VBA code using xlwings, use the following command in the terminal:

```
xlwings vba edit --file "C:\path\to\your\Home Sales Data.xlsm"
```

**Note:** This command only works on Windows.

### Troubleshooting

If you encounter the following error:

```
pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Programmatic access to Visual Basic Project is not trusted\n', 'xlmain11.chm', 0, -2146827284), None)
```

This error typically occurs when trying to access the VBA project in an Excel workbook using Python and the win32com library. The issue is related to Excel's security settings. To resolve it:

1. Open Excel and go to File > Options > Trust Center.
2. Click on "Trust Center Settings".
3. In the Trust Center window, select "Macro Settings" on the left.
4. Check the box next to "Trust access to the VBA project object model".
5. Click OK to save the changes and close Excel.

After making these changes, restart your Python script and try accessing the Excel workbook again. This should resolve the "Programmatic access to Visual Basic Project is not trusted" error.

## Contributing

Feel free to contribute to this repository by submitting pull requests or opening issues for any bugs or feature requests.

## License
Apache 2.0



