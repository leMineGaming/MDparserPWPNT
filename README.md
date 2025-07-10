# PWPNT MDparser
![test](https://badgen.net/badge/status/stable/green?icon=github)
![test2](https://badgen.net/badge/latest/v1.1/blue?icon=version)

This PowerPoint add-in allows you to quickly convert Markdown-formatted text into a series of styled slides. It supports headings, bullet points, ordered lists, and inline formatting like **bold**, *italic*, _underline_, ==highlight==, and ~~strikethrough~~.

## Features

- **Insert Markdown as Slides:** Paste your Markdown and instantly generate slides.
- **Headings to Slides:**
  - `# Heading` creates a new main slide (title or content, context-aware).
  - `## Subheading` creates a subtitle and splits content into separate slides.
  - `### Sub-subheading` creates a styled sub-heading within slides without splitting content.
- **Lists:** Unordered (`- item`) and ordered (`1. item`) lists are supported, including indentation for sub-lists.
- **Inline Formatting:**
  - `**bold**`
  - `*italic*`
  - `_underline_` or `__underline__`
  - `==highlight==`
  - `~~strikethrough~~`
 
## Slide Creation Rules

- If a `#` heading is followed by another `#` heading, or by a `##` heading with no content before the next `#`, a title slide is created.
- If a `#` heading is followed by `##` subtitles with content, each `##` section creates its own content slide (with the main `#` as the title and the `##` as the subtitle).
- Subtitles on content slides are styled (bold, larger font, no bullet).
- Only list items beginning with `- ` or `1. ` receive bullets/numbers.

## Custom Formatting

- `==highlight==` → Highlights text in red.
- `~~strikethrough~~` → Applies strikethrough.

Icons in the app are made by Icons8
icons8.com
