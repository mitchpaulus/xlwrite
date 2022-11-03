# `xlwrite`

`xlwrite` is a command line utility to take data from text files and put them into Excel files.

## Examples

By default, `xlwrite` expects tab-separated data.
You can then write the block of data to an excel file like:

```console
xlwrite block A1 datafile.tsv spreadsheet.xlsx
```

If you need to write to a particular sheet, you can specify that using the normal Excel range syntax.
Notice that you will normally need to get the single quotes past your shell.
Typically surrounding with double quotes will suffice.

```console
xlwrite block "'My Sheet'!A1" datafile.tsv spreadsheet.xlsx
```
