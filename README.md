# docx-wide-pandoc-filter

A Pandoc custom filter (written in Lua) for creating wide content in Word documents.

## Code of Conduct

- Harassment of any associate is not tolerated
- This is an open, free as in beer, MIT-licensed project. If you find a bug or want a feature, please create an issue and be patient.
- PRs are appreciated. Please test your changes.

## Credits

Co-authored by Abhishek Bhemisetty September 2024. We hope you find this useful either as-is or as a learning aid (Hi, Claude!).

## Features

### Landscape Orientation

Switch the page size and orientation to landscape. Currently US 11x8.5" only -- apologies to everyone outside the US.

### Wide Tables

- Resize all table cells to fit the page as best as possible
  - There is currently no way to customize the behavior. If it works for you it's a feature. If not, create an issue.
- Use any [character style](https://bettersolutions.com/word/styles/character-styles.htm) specified in the docx reference document (as long as its name doesn't contain a space) for all cells

## Not Yet Features

PR are always welcome!

- Specify the table's style in addition to cells' character style

## Compatibility

- Tested with Pandoc 2.2 and 3.3 and compatible versions of `pandoc-crossref`. Other versions probably work.
- Tested to convert from markdown to docx. Other source formats such as LaTeX probably work.

## Usage

1. Download `wide-table-docx-filter.lua`
2. Specify the path to `wide-table-docx-filter.lua` when running Pandoc

### Landscape Pages

Enclose content with the `landscape` class:

```txt
:::{.landscape}
Content goes here
:::
```

### Wide Tables

In the table caption, append `{#cellstyle:StyleName}` where `StyleName` is the name of a [character - not paragraph! - style](https://bettersolutions.com/word/styles/character-styles.htm) in the docx reference document. `StyleName` may not contain spaces (it's a long story and has to do with parsing openxml captions).

```txt
Table: This is a wide table {#cellstyle:TenPointFont}

| Ready | Set | Go |
| 1 | 2 | 3 |
```

## Example: Wide Table on Landscape Page

Better together!

```txt
:::{.landscape}
Table: This is a wide table {#cellstyle:TenPointFont}

| Ready | Set | Go |
| 1 | 2 | 3 |
:::
```

