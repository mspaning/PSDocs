﻿# PSDocs_Options

## about_PSDocs_Options

## SHORT DESCRIPTION

Describes additional options that can be used during markdown generation.

## LONG DESCRIPTION

PSDocs lets you use options when calling `Invoke-PSDocument` to change how documents are generated.
This topic describes what options are available, when to and how to use them.

The following workspace options are available for use:

- [Execution.LanguageMode](#executionlanguagemode)
- [Markdown.ColumnPadding](#markdowncolumnpadding)
- [Markdown.Encoding](#markdownencoding)
- [Markdown.SkipEmptySections](#markdownskipemptysections)
- [Markdown.UseEdgePipes](#markdownuseedgepipes)
- [Markdown.WrapSeparator](#markdownwrapseparator)
- [Output.Culture](#outputculture)
- [Output.Path](#outputpath)

Options can be used by:

- Using the `-Option` parameter of `Invoke-PSDocument` with an object created with `New-PSDocumentOption`.
- Using the `-Option` parameter of `Invoke-PSDocument` with a hashtable.
- Using the `-Option` parameter of `Invoke-PSDocument` with a YAML file.
- Configuring the default options file `ps-docs.yaml`.

As mentioned above, a options object can be created with `New-PSDocumentOption` see cmdlet help for syntax and examples.

When using a hashtable, `@{}`, one or more options can be specified with the `-Option` parameter using a dotted notation.

For example:

```powershell
$option = @{ 'markdown.wrapseparator' = ' '; 'markdown.encoding' = 'UTF8' };
Invoke-PSDocument -Path . -Option $option;
```

`markdown.wrapseparator` is an example of an option that can be used.
Please see the following sections for other options can be used.

Another option is to use an external file, formatted as YAML, instead of having to create an options object manually each time.
This YAML file can be used with `Invoke-PSDocument` to quickly build documentation in a repeatable way.

YAML properties are specified using lower camel case, for example:

```yaml
markdown:
  wrapSeparator: '\'
```

By default PSDocs will automatically look for a file named `ps-docs.yaml` in the current working directory.
Alternatively, you can specify a YAML file in the `-Option` parameter.

For example:

```powershell
Invoke-PSDocument -Path . -Option '.\myconfig.yml'.
```

### Execution.LanguageMode

Unless PowerShell has been constrained, full language features of PowerShell are available to use within document definitions.
In locked down environments, a reduced set of language features may be desired.

When PSDocs is executed in an environment configured for Device Guard, only constrained language features are available.

The following language modes are available for use in PSDocs:

- FullLanguage
- ConstrainedLanguage

This option can be specified using:

```powershell
# PowerShell: Using the Execution.LanguageMode hashtable key
$option = New-PSDocumentOption -Option @{ 'Execution.LanguageMode' = 'ConstrainedLanguage' }
```

```yaml
# YAML: Using the execution/languageMode YAML property
execution:
  languageMode: ConstrainedLanguage
```

### Markdown.ColumnPadding

Sets how table column padding should be handled in markdown.
This doesn't affect how tables are rendered but can greatly assist readability of markdown source files.

The following padding options are available:

- None - No padding will be used and column values will directly follow table pipe (`|`) column separators.
- Single - A single space will be used to pad the column value.
- MatchHeader - Will pad the header with a single space, then pad the column value, to the same width as the header (default).

When a column is set to a specific width with a property expression, `MatchHeader` will be ignored.
Columns without a width set will apply `MatchHeader` as normal.

Example markdown using `None`:

```markdown
|Name|Value|
|----|-----|
|Mon|Key|
|Really long name|Really long value|
```

Example markdown using `Single`:

```markdown
| Name | Value |
| ---- | ----- |
| Mon | Key |
| Really long name | Really long value |
```

Example markdown using `MatchHeader`:

```markdown
| Name | Value |
| ---- | ----- |
| Mon  | Key   |
| Really long name | Really long value |
```

This option can be specified using:

```powershell
# PowerShell: Using the Markdown.ColumnPadding hashtable key
$option = New-PSDocumentOption -Option @{ 'Markdown.ColumnPadding' = 'MatchHeader' }
```

```yaml
# YAML: Using the markdown/columnPadding YAML property
markdown:
  columnPadding: MatchHeader
```

### Markdown.Encoding

Sets the text encoding used for markdown output files. One of the following values can be used:

- Default
- UTF8
- UTF7
- Unicode
- UTF32
- ASCII

By default `Default` is used which is UTF-8 without byte order mark (BOM) is used.

Use this option with `Invoke-PSDocument`.
When the `-Encoding` parameter is used, it will override any value set in configuration.

This option can be specified using:

```powershell
# PowerShell: Using the Encoding parameter
$option = New-PSDocumentOption -Encoding 'UTF8';
```

```powershell
# PowerShell: Using the Markdown.Encoding hashtable key
$option = New-PSDocumentOption -Option @{ 'Markdown.Encoding' = 'UTF8' }
```

```yaml
# YAML: Using the markdown/encoding YAML property
markdown:
  encoding: UTF8
```

Prior to PSDocs v0.4.0 the only encoding supported was ASCII.

### Markdown.SkipEmptySections

From PSDocs v0.5.0 onward, `Section` blocks that are empty are omitted from markdown output by default.
i.e. `Markdown.SkipEmptySections` is `$True`.

To include empty sections (the same as PSDocs v0.4.0 or older) in markdown output either use the `-Force` parameter on a specific `Section` block or set the option `Markdown.SkipEmptySections` to `$False`.

This option can be specified using:

```powershell
# PowerShell: Using the Markdown.SkipEmptySections hashtable key
$option = New-PSDocumentOption -Option @{ 'Markdown.SkipEmptySections' = $False }
```

```yaml
# YAML: Using the markdown/skipEmptySections YAML property
markdown:
  skipEmptySections: false
```

### Markdown.UseEdgePipes

This option determines when pipes on the edge of a table should be used.
This option can improve readability of markdown source files, but may not be supported by all markdown renderers.

Edge pipes are always required if the table has a single column, so this option only applies for tables with more then one column.

The following options for edge pipes are:

- WhenRequired - Will not use edge pipes for tables with more when one column (default).
- Always - Will always use edge pipes.

Example markdown using `WhenRequired`:

```markdown
Name|Value
----|-----
Mon|Key
```

Example markdown using `Always`:

```markdown
|Name|Value|
|----|-----|
|Mon|Key|
```

Example markdown using `WhenRequired` and column padding of `MatchHeader`:

```markdown
Name | Value
---- | -----
Mon  | Key
```

This option can be specified using:

```powershell
# PowerShell: Using the Markdown.UseEdgePipes hashtable key
$option = New-PSDocumentOption -Option @{ 'Markdown.UseEdgePipes' = 'WhenRequired' }
```

```yaml
# YAML: Using the markdown/useEdgePipes YAML property
markdown:
  useEdgePipes: WhenRequired
```

### Markdown.WrapSeparator

This option specifies the character/string to use when wrapping lines in a table cell.
When a table cell contains CR and LF characters, these characters must be substituted so that the table in rendered correctly because they also have special meaning in markdown.

By default a single space is used.
However different markdown parsers may be able to natively render a line break using alternative combinations such as `\` or `<br />`.

This option can be specified using:

```powershell
# PowerShell: Using the Markdown.WrapSeparator hashtable key
$option = New-PSDocumentOption -Option @{ 'Markdown.WrapSeparator' = '\' }
```

```yaml
# YAML: Using the markdown/wrapSeparator YAML property
markdown:
  wrapSeparator: '\'
```

### Output.Culture

Specifies a list of cultures for building documents such as _en-AU_, and _en-US_.
Documents are written to culture specific subdirectories when multiple cultures are generated.

Use this option with `Invoke-PSDocument`.
When the `-Culture` parameter is used, it will override any value set in configuration.

This option can be specified using:

```powershell
# PowerShell: Using the Output.Culture hashtable key
$option = New-PSDocumentOption -Option @{ 'Output.Culture' = 'en-US', 'en-AU' }
```

```yaml
# YAML: Using the output/culture YAML property
output:
  culture:
  - 'en-US'
```

### Output.Path

Configures the directory path to store markdown files created based on the specified document template.
This path will be automatically created if it doesn't exist.

Use this option with `Invoke-PSDocument`.
When the `-OutputPath` parameter is used, it will override any value set in configuration.

This option can be specified using:

```powershell
# PowerShell: Using the Output.Path hashtable key
$option = New-PSDocumentOption -Option @{ 'Output.Path' = 'out/' }
```

```yaml
# YAML: Using the output/path YAML property
output:
  path: 'out/'
```

## EXAMPLES

### Example ps-docs.yaml

```yaml
# Set markdown options
execution:
  languageMode: ConstrainedLanguage
markdown:
  # Use UTF-8 with BOM
  encoding: UTF8
  skipEmptySections: false
  wrapSeparator: '\'
output:
  culture:
  - 'en-US'
  path: 'out/'
```

### Default ps-docs.yaml

```yaml
# These are the default options.
# Only properties that differ from the default values need to be specified.
execution:
  languageMode: FullLanguage
markdown:
  encoding: Default
  skipEmptySections: true
  wrapSeparator: ' '
  columnPadding: MatchHeader
  useEdgePipes: WhenRequired
output:
  culture: [ ]
  path: null
```

## NOTE

An online version of this document is available at https://github.com/BernieWhite/PSDocs/blob/main/docs/concepts/PSDocs/en-US/about_PSDocs_Options.md.

## SEE ALSO

- [Invoke-PSDocument](https://github.com/BernieWhite/PSDocs/blob/main/docs/commands/PSDocs/en-US/Invoke-PSDocument.md)
- [New-PSDocumentOption](https://github.com/BernieWhite/PSDocs/blob/main/docs/commands/PSDocs/en-US/New-PSDocumentOption.md)

## KEYWORDS

- Options
- Markdown
- PSDocument
