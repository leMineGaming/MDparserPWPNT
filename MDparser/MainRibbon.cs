using Microsoft.Office.Tools.Ribbon;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MDparser
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e) { }

        private void insertMarkdown_Click(object sender, RibbonControlEventArgs e)
        {
            using (var form = new MarkdownInputForm())
            {
                if (form.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

                string markdown = form.MarkdownTextInput;
                var app = Globals.ThisAddIn.Application;
                var presentation = app.ActivePresentation;

                // Get the index of the currently selected slide (1-based)
                int selectedIndex = 1;
                var sel = app.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides && sel.SlideRange.Count > 0)
                    selectedIndex = sel.SlideRange[1].SlideIndex;
                else if (presentation.Slides.Count > 0)
                    selectedIndex = presentation.Slides.Count; // fallback to last slide

                int insertIndex = selectedIndex + 1; // Insert after selected slide

                PowerPoint.Slide currentSlide = null;
                PowerPoint.Shape bodyShape = null;
                string currentTitle = null;

                var lines = markdown.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                int i = 0;
                while (i < lines.Length)
                {
                    string line = lines[i].Trim();

                    // Detect top-level title
                    if (line.StartsWith("# "))
                    {
                        string nextTitle = line.Substring(2).Trim();

                        // Look ahead to check for a ## subtitle
                        int lookahead = i + 1;
                        while (lookahead < lines.Length && string.IsNullOrWhiteSpace(lines[lookahead]))
                            lookahead++;

                        bool nextIsHash = lookahead < lines.Length && lines[lookahead].Trim().StartsWith("# ");
                        bool nextIsDoubleHash = lookahead < lines.Length && lines[lookahead].Trim().StartsWith("## ");

                        if (nextIsHash)
                        {
                            // # followed by # : pure title slide
                            currentSlide = presentation.Slides.Add(insertIndex, PowerPoint.PpSlideLayout.ppLayoutTitle);
                            insertIndex++;
                            currentSlide.Shapes[1].TextFrame.TextRange.Text = nextTitle;
                            currentSlide.Shapes[2].TextFrame.TextRange.Text = "";
                            currentTitle = nextTitle;
                            bodyShape = null;
                            i = lookahead;
                            continue;
                        }
                        else if (nextIsDoubleHash)
                        {
                            // See if the ## is immediately followed by another # (with only whitespace in between)
                            string subtitle = lines[lookahead].Trim().Substring(3).Trim();
                            int postSubtitle = lookahead + 1;

                            while (postSubtitle < lines.Length && string.IsNullOrWhiteSpace(lines[postSubtitle]))
                                postSubtitle++;

                            bool afterSubtitleIsHash = postSubtitle < lines.Length && lines[postSubtitle].Trim().StartsWith("# ");
                            bool afterSubtitleIsDoubleHash = postSubtitle < lines.Length && lines[postSubtitle].Trim().StartsWith("## ");

                            // If ## is followed by # (and no content in between), treat as title slide
                            if (afterSubtitleIsHash)
                            {
                                currentSlide = presentation.Slides.Add(insertIndex, PowerPoint.PpSlideLayout.ppLayoutTitle);
                                insertIndex++;
                                currentSlide.Shapes[1].TextFrame.TextRange.Text = nextTitle;
                                currentSlide.Shapes[2].TextFrame.TextRange.Text = subtitle;
                                currentTitle = nextTitle;
                                bodyShape = null;
                                i = postSubtitle;
                                continue;
                            }
                            // If ## is followed by ## with no content, treat as title slide (rare)
                            else if (afterSubtitleIsDoubleHash)
                            {
                                currentSlide = presentation.Slides.Add(insertIndex, PowerPoint.PpSlideLayout.ppLayoutTitle);
                                insertIndex++;
                                currentSlide.Shapes[1].TextFrame.TextRange.Text = nextTitle;
                                currentSlide.Shapes[2].TextFrame.TextRange.Text = subtitle;
                                currentTitle = nextTitle;
                                bodyShape = null;
                                i = postSubtitle;
                                continue;
                            }
                            // If there is content after ## (before next ## or #), treat as normal section slide(s)
                            else
                            {
                                currentTitle = nextTitle;
                                i = lookahead; // will process ## in next iteration
                                continue;
                            }
                        }
                        else
                        {
                            // # at end or with content, treat as normal text slide
                            currentSlide = presentation.Slides.Add(insertIndex, PowerPoint.PpSlideLayout.ppLayoutText);
                            insertIndex++;
                            currentSlide.Shapes[1].TextFrame.TextRange.Text = nextTitle;
                            currentSlide.Shapes[2].TextFrame.TextRange.Text = "";
                            currentTitle = nextTitle;
                            bodyShape = currentSlide.Shapes[2];
                            i++;
                            // Process body content
                            while (i < lines.Length && !lines[i].Trim().StartsWith("# ") && !lines[i].Trim().StartsWith("## "))
                            {
                                string bodyLine = lines[i];
                                if (IsUnorderedList(bodyLine, out string ulText, out int ulIndent))
                                {
                                    AddParagraphWithFormatting(bodyShape, ulText, false, isOrdered: false, orderNum: 0, indent: ulIndent, isUnordered: true);
                                }
                                else if (IsOrderedList(bodyLine, out int orderNum, out string olText, out int olIndent))
                                {
                                    AddParagraphWithFormatting(bodyShape, olText, false, isOrdered: true, orderNum: orderNum, indent: olIndent, isUnordered: false);
                                }
                                else if (bodyLine.Trim().StartsWith("### "))
                                {
                                    // Handle ### headings as styled sub-headings
                                    string subHeading = bodyLine.Trim().Substring(4).Trim();
                                    AddParagraphWithFormatting(bodyShape, subHeading, false, isOrdered: false, orderNum: 0, indent: 0, isUnordered: false, isSubHeading: true);
                                }
                                else if (!string.IsNullOrWhiteSpace(bodyLine))
                                {
                                    AddParagraphWithFormatting(bodyShape, bodyLine.Trim(), false, isOrdered: false, orderNum: 0, indent: 0, isUnordered: false);
                                }
                                i++;
                            }
                            continue;
                        }
                    }

                    // Detect section subtitle (## ...)
                    if (line.StartsWith("## "))
                    {
                        string currentSubtitle = line.Substring(3).Trim();

                        // Gather all lines until next ## or #, or end
                        List<string> bodyLines = new List<string>();
                        int j = i + 1;
                        while (j < lines.Length && !lines[j].Trim().StartsWith("## ") && !lines[j].Trim().StartsWith("# "))
                        {
                            bodyLines.Add(lines[j]);
                            j++;
                        }

                        currentSlide = presentation.Slides.Add(insertIndex, PowerPoint.PpSlideLayout.ppLayoutText);
                        insertIndex++;
                        currentSlide.Shapes[1].TextFrame.TextRange.Text = currentTitle ?? "";

                        var bodyTextRange = currentSlide.Shapes[2].TextFrame.TextRange;
                        bodyTextRange.Text = ""; // Clear any default

                        // Insert subtitle as the first paragraph, styled, no bullet
                        bodyTextRange.InsertAfter(currentSubtitle + "\r");
                        var subtitleRange = bodyTextRange.Paragraphs(1);
                        subtitleRange.Font.Bold = Office.MsoTriState.msoTrue;
                        subtitleRange.Font.Size = 26; // slightly larger
                        subtitleRange.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoFalse;

                        // Insert a blank paragraph after the subtitle for the body to start clean
                        bodyTextRange.InsertAfter("\r");

                        bodyShape = currentSlide.Shapes[2];

                        foreach (var bodyLine in bodyLines)
                        {
                            if (IsUnorderedList(bodyLine, out string ulText, out int ulIndent))
                            {
                                AddParagraphWithFormatting(bodyShape, ulText, false, isOrdered: false, orderNum: 0, indent: ulIndent, isUnordered: true);
                            }
                            else if (IsOrderedList(bodyLine, out int orderNum, out string olText, out int olIndent))
                            {
                                AddParagraphWithFormatting(bodyShape, olText, false, isOrdered: true, orderNum: orderNum, indent: olIndent, isUnordered: false);
                            }
                            else if (bodyLine.Trim().StartsWith("### "))
                            {
                                // Handle ### headings as styled sub-headings
                                string subHeading = bodyLine.Trim().Substring(4).Trim();
                                AddParagraphWithFormatting(bodyShape, subHeading, false, isOrdered: false, orderNum: 0, indent: 0, isUnordered: false, isSubHeading: true);
                            }
                            else if (!string.IsNullOrWhiteSpace(bodyLine))
                            {
                                AddParagraphWithFormatting(bodyShape, bodyLine.Trim(), false, isOrdered: false, orderNum: 0, indent: 0, isUnordered: false);
                            }
                        }

                        i = j;
                        continue;
                    }

                    // If we get here, just skip
                    i++;
                }
            }
        }

        private void AddParagraphWithFormatting(PowerPoint.Shape shape, string markdownText, bool isTitle, bool isOrdered, int orderNum, int indent, bool isUnordered = false, bool isSubHeading = false)
        {
            var tr2 = shape.TextFrame2.TextRange;
            string prefix = "";
            if (isOrdered)
            {
                prefix = $"{orderNum}. ";
            }
            else if (isUnordered)
            {
                prefix = "• ";
            }

            string plainText = StripMarkdownFormatting(markdownText);

            int insertStart = tr2.Length + 1; // PowerPoint TextRange2 is 1-based!
            tr2.InsertAfter(prefix + plainText + "\r");
            int idx = tr2.Paragraphs.Count;
            if (idx == 0) idx = 1;
            var para2 = tr2.Paragraphs[idx];
            var pf2 = para2.ParagraphFormat;

            pf2.LeftIndent = 20f * indent;
            pf2.FirstLineIndent = (isOrdered || isUnordered) ? -10f : 0f;
            pf2.IndentLevel = indent + 1;

            if (isTitle)
            {
                para2.Font.Size = 32;
                para2.Font.Bold = Office.MsoTriState.msoTrue;
                para2.Font.UnderlineStyle = Office.MsoTextUnderlineType.msoNoUnderline;
                pf2.Bullet.Visible = Office.MsoTriState.msoFalse;
            }
            else if (isSubHeading)
            {
                para2.Font.Size = 22;
                para2.Font.Bold = Office.MsoTriState.msoTrue;
                para2.Font.UnderlineStyle = Office.MsoTextUnderlineType.msoNoUnderline;
                pf2.Bullet.Visible = Office.MsoTriState.msoFalse;
            }
            else
            {
                para2.Font.Size = 18;
                para2.Font.Bold = Office.MsoTriState.msoFalse;
                para2.Font.UnderlineStyle = Office.MsoTextUnderlineType.msoNoUnderline;
                pf2.Bullet.Visible = Office.MsoTriState.msoFalse;
            }

            ApplyInlineFormattingToPowerPoint(para2, prefix, markdownText, plainText);
        }

        private string StripMarkdownFormatting(string text)
        {
            // Remove ==highlight== and ~~strikethrough~~ as well as bold/italic/underline
            return Regex.Replace(text,
                @"(\=\=([^\=]+)\=\=)|(\~\~([^\~]+)\~\~)|(\*\*([^\*]+)\*\*)|(\*([^\*]+)\*)|(__([^_]+)__)|(_([^_]+)_)",
                match =>
                {
                    if (match.Groups[2].Success) return match.Groups[2].Value;      // ==highlight==
                    if (match.Groups[4].Success) return match.Groups[4].Value;      // ~~strikethrough~~
                    if (match.Groups[6].Success) return match.Groups[6].Value;      // bold **text**
                    if (match.Groups[8].Success) return match.Groups[8].Value;      // italic *text*
                    if (match.Groups[10].Success) return match.Groups[10].Value;    // underline __text__
                    if (match.Groups[12].Success) return match.Groups[12].Value;    // underline _text_
                    return match.Value;
                });
        }

        private void ApplyInlineFormattingToPowerPoint(Office.TextRange2 para, string prefix, string markdownText, string plainText)
        {
            int prefixLen = prefix.Length;
            int plainIdx = 0, mdIdx = 0;
            var formatting = new List<(int start, int len, string type)>();

            while (mdIdx < markdownText.Length)
            {
                // ==highlight==
                if (markdownText[mdIdx] == '=' && mdIdx + 1 < markdownText.Length && markdownText[mdIdx + 1] == '=')
                {
                    int end = markdownText.IndexOf("==", mdIdx + 2);
                    if (end > mdIdx + 2)
                    {
                        formatting.Add((plainIdx + prefixLen, end - (mdIdx + 2), "highlight"));
                        mdIdx += 2;
                        for (int k = mdIdx; k < end; ++k, ++plainIdx) ;
                        mdIdx = end + 2;
                        continue;
                    }
                }
                // ~~strikethrough~~
                if (markdownText[mdIdx] == '~' && mdIdx + 1 < markdownText.Length && markdownText[mdIdx + 1] == '~')
                {
                    int end = markdownText.IndexOf("~~", mdIdx + 2);
                    if (end > mdIdx + 2)
                    {
                        formatting.Add((plainIdx + prefixLen, end - (mdIdx + 2), "strikethrough"));
                        mdIdx += 2;
                        for (int k = mdIdx; k < end; ++k, ++plainIdx) ;
                        mdIdx = end + 2;
                        continue;
                    }
                }
                // **bold**
                if (markdownText[mdIdx] == '*' && mdIdx + 1 < markdownText.Length && markdownText[mdIdx + 1] == '*')
                {
                    int end = markdownText.IndexOf("**", mdIdx + 2);
                    if (end > mdIdx + 2)
                    {
                        formatting.Add((plainIdx + prefixLen, end - (mdIdx + 2), "bold"));
                        mdIdx += 2;
                        for (int k = mdIdx; k < end; ++k, ++plainIdx) ;
                        mdIdx = end + 2;
                        continue;
                    }
                }
                // __underline__
                if (markdownText[mdIdx] == '_' && mdIdx + 1 < markdownText.Length && markdownText[mdIdx + 1] == '_')
                {
                    int end = markdownText.IndexOf("__", mdIdx + 2);
                    if (end > mdIdx + 2)
                    {
                        formatting.Add((plainIdx + prefixLen, end - (mdIdx + 2), "underline"));
                        mdIdx += 2;
                        for (int k = mdIdx; k < end; ++k, ++plainIdx) ;
                        mdIdx = end + 2;
                        continue;
                    }
                }
                // *italic*
                if (markdownText[mdIdx] == '*' && (mdIdx == 0 || markdownText[mdIdx - 1] != '*'))
                {
                    int end = markdownText.IndexOf('*', mdIdx + 1);
                    if (end > mdIdx + 1 && (end + 1 >= markdownText.Length || markdownText[end + 1] != '*'))
                    {
                        formatting.Add((plainIdx + prefixLen, end - (mdIdx + 1), "italic"));
                        mdIdx += 1;
                        for (int k = mdIdx; k < end; ++k, ++plainIdx) ;
                        mdIdx = end + 1;
                        continue;
                    }
                }
                // _underline_
                if (markdownText[mdIdx] == '_' && (mdIdx == 0 || markdownText[mdIdx - 1] != '_'))
                {
                    int end = markdownText.IndexOf('_', mdIdx + 1);
                    if (end > mdIdx + 1 && (end + 1 >= markdownText.Length || markdownText[end + 1] != '_'))
                    {
                        formatting.Add((plainIdx + prefixLen, end - (mdIdx + 1), "underline"));
                        mdIdx += 1;
                        for (int k = mdIdx; k < end; ++k, ++plainIdx) ;
                        mdIdx = end + 1;
                        continue;
                    }
                }
                mdIdx++;
                plainIdx++;
            }

            foreach (var fmt in formatting)
            {
                try
                {
                    var range = para.Characters[fmt.start + 1, fmt.len];
                    if (fmt.type == "bold") range.Font.Bold = Office.MsoTriState.msoTrue;
                    if (fmt.type == "italic") range.Font.Italic = Office.MsoTriState.msoTrue;
                    if (fmt.type == "underline") range.Font.UnderlineStyle = Office.MsoTextUnderlineType.msoUnderlineSingleLine;
                    if (fmt.type == "highlight")
                    {
                        // Yellow highlight
                        range.Font.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    if (fmt.type == "strikethrough")
                    {
                        range.Font.Strike = Office.MsoTextStrike.msoSingleStrike;
                    }
                }
                catch { }
            }
        }

        private bool IsUnorderedList(string line, out string text, out int indent)
        {
            indent = 0;
            text = null;
            int i = 0;
            int spaces = 0, tabs = 0;
            while (i < line.Length && (line[i] == '\t' || line[i] == ' '))
            {
                if (line[i] == '\t') tabs++;
                if (line[i] == ' ') spaces++;
                i++;
            }
            indent = tabs + spaces / 4; // 4 spaces = 1 indent
            if (line.Substring(i).StartsWith("- "))
            {
                text = line.Substring(i + 2).Trim();
                return true;
            }
            return false;
        }

        private bool IsOrderedList(string line, out int number, out string text, out int indent)
        {
            indent = 0;
            number = 0;
            text = null;
            int i = 0;
            int spaces = 0, tabs = 0;
            while (i < line.Length && (line[i] == '\t' || line[i] == ' '))
            {
                if (line[i] == '\t') tabs++;
                if (line[i] == ' ') spaces++;
                i++;
            }
            indent = tabs + spaces / 4;

            // e.g. 1. text
            int dot = line.IndexOf('.', i);
            if (dot > i && int.TryParse(line.Substring(i, dot - i), out number)
                && line.Length > dot + 1 && line[dot + 1] == ' ')
            {
                text = line.Substring(dot + 2).Trim();
                return true;
            }
            return false;
        }
    }
}