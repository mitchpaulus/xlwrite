# `xlwrite`

`xlwrite` is a command line utility to take data from text files and put them into Excel files.

## Examples

By default, `xlwrite` expects tab-separated data.
You can then write the block of data to an excel file like:

```sh
xlwrite block A1 datafile.tsv spreadsheet.xlsx
```

If you need to write to a particular sheet, you can specify that using the normal Excel range syntax.
Notice that you will normally need to get the single quotes past your shell.
Typically surrounding with double quotes will suffice.

```sh
xlwrite block "'My Sheet'!A1" datafile.tsv spreadsheet.xlsx
```

## Installation

`xlwrite` is currently compiled for x64 machines on Linux, Windows, and MacOS.
In the [releases](https://github.com/mitchpaulus/xlwrite/releases) you will find compiled single file binaries.
These come in two flavors: *framework-dependent* and *self-contained*.
The *framework-dependent* versions require a .NET runtime to be available on the machine.
The *self-contained* versions should run with no additional dependencies.
Because the *framework-dependent* version doesn't need all the runtimes bundled, it is significantly smaller.

But once you have the executable, put it in your `PATH` environment variable and you should be on your way!
