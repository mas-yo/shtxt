

# shtxt

Shtxt is flexible, extensible, configurable spread-sheet-to-text converter.

```
Usage:
  shtxt [options] [<input-files>...]

Arguments:
  <input-files>

Options:
  -p, --input-pattern <input-pattern>          Input file name pattern
  -l, --version-list <version-list>            Version list file
  -r, --current-version <current-version>      Current version
  -o, --output-dir <output-dir>                Output directory
  -n, --newline <newline>                      Newline code(cr,lf,crlf)
  -f, --text-format <text-format>              Output format(csv,tsv)
  --comment-starts-with <comment-starts-with>  String that indicates comment
  --table-name-tag <table-name-tag>            Tag string for table name
  --column-name-tag <column-name-tag>          Tag string for column name
  --column-control-tag <column-control-tag>    Tag string for column control
  -c, --config-file <config-file>              Config file
  --version                                    Show version information
  -?, -h, --help                               Show help and usage information
```



## Features

- Load spread sheet files, detect control commands, interpret and process, then convert to text file
- Can switch line/column by versioning system
- Most of parameters are configurable by command line arguments and config file
- Converts multi files on multi threads
- Achieves software flexibility with data flow design

