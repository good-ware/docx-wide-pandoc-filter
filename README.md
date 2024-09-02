# docx-wide-table-pandoc-filter

A Pandoc custom filter for creating wide tables in Word documents.

## Code of Conduct

- Harassment of any of our users or associates is not tolerated
- This is a open source free as in beer project. If you find a bug or want a feature, create an issue and be patient.
- PRs are always welcomed. Please test your changes.
- Be nice

## Features

- Switch the page size and orientation to landscape (currently US 11x8.5" only -- sorry to the rest of the world)
- Resize cells to fit the page as best as possible
- Use any [character style](https://bettersolutions.com/word/styles/character-styles.htm) defined in the reference document in cells as long as the name doesn't contain a space

## Compatibility

- Tested with Pandoc 2.2 and 3.3 and compatible versions of `pandoc-crossref`
- Probably works with additional versions but you're on your own. Please let us know via issues and we will update the README.
- Tested for converting from markdown to docx. Other 'from' formats such as LaTeX are likely to work.

## Usage

1. Download `wide-table-docx-filter.lua`
2. Specify the path to `wide-table-docx-filter.lua` when running Pandoc

### Switching to Landscape

Wide tables probably require landscape pages. Unfortunately, only landscape is provided. Other sizes such as A4 could be added. Please submit an issue if desired.

Enclose content with the `landscape` class:

```txt
:::{.landscape}
Text goes here
:::
```

### Wide Tables

In the table caption, append `{#style:StyleName}` where `StyleName` is the name of a [character - not paragraph! - style](https://bettersolutions.com/word/styles/character-styles.htm) in the reference document. `StyleName` may not contain spaces (it's a long story and has to do with parsing openxml captions).

```txt
Table: This is a wide table {#style:TenPointFont}
```
txt
