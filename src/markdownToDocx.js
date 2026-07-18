import { marked } from "marked";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ExternalHyperlink,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
} from "docx";

const HEADING_LEVELS = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
};

const CODE_FONT = "Consolas";
const CODE_SHADING = { fill: "F2F2F2" };
const QUOTE_INDENT = { left: 480 };
const QUOTE_BORDER = {
  left: { style: BorderStyle.SINGLE, size: 12, color: "CCCCCC", space: 8 },
};

function alignFor(align) {
  if (align === "center") return AlignmentType.CENTER;
  if (align === "right") return AlignmentType.RIGHT;
  return AlignmentType.LEFT;
}

function makeRun(text, fmt) {
  return new TextRun({
    text,
    bold: fmt.bold,
    italics: fmt.italics,
    strike: fmt.strike,
    font: fmt.code ? CODE_FONT : undefined,
    shading: fmt.code ? CODE_SHADING : undefined,
    style: fmt.hyperlink ? "Hyperlink" : undefined,
  });
}

function inlineToRuns(tokens, fmt = {}) {
  const runs = [];
  for (const token of tokens || []) {
    switch (token.type) {
      case "text":
      case "escape":
        runs.push(makeRun(token.text, fmt));
        break;
      case "strong":
        runs.push(...inlineToRuns(token.tokens, { ...fmt, bold: true }));
        break;
      case "em":
        runs.push(...inlineToRuns(token.tokens, { ...fmt, italics: true }));
        break;
      case "del":
        runs.push(...inlineToRuns(token.tokens, { ...fmt, strike: true }));
        break;
      case "codespan":
        runs.push(makeRun(token.text, { ...fmt, code: true }));
        break;
      case "link":
        runs.push(
          new ExternalHyperlink({
            link: token.href,
            children: inlineToRuns(
              token.tokens || [{ type: "text", text: token.text }],
              { ...fmt, hyperlink: true }
            ),
          })
        );
        break;
      case "image":
        runs.push(makeRun(`[Image: ${token.text || token.href}]`, { ...fmt, italics: true }));
        break;
      case "br":
        runs.push(new TextRun({ text: "", break: 1 }));
        break;
      default:
        if (token.tokens) runs.push(...inlineToRuns(token.tokens, fmt));
        else if (token.raw) runs.push(makeRun(token.raw, fmt));
    }
  }
  return runs;
}

function quoteWrap(paragraphOptions, quoted) {
  if (!quoted) return paragraphOptions;
  return {
    ...paragraphOptions,
    indent: QUOTE_INDENT,
    border: QUOTE_BORDER,
  };
}

function listPrefixRun(ordered, index) {
  const text = ordered ? `${index + 1}. ` : "• ";
  return new TextRun({ text });
}

function blockToDocx(token, ctx = { listLevel: 0, quoted: false }) {
  switch (token.type) {
    case "heading":
      return [
        new Paragraph(
          quoteWrap(
            {
              heading: HEADING_LEVELS[token.depth] || HeadingLevel.HEADING_6,
              children: inlineToRuns(token.tokens),
              spacing: { before: 240, after: 120 },
            },
            ctx.quoted
          )
        ),
      ];

    case "paragraph":
      return [
        new Paragraph(
          quoteWrap(
            {
              children: inlineToRuns(token.tokens),
              spacing: { after: 160 },
            },
            ctx.quoted
          )
        ),
      ];

    case "code": {
      const lines = token.text.split("\n");
      return lines.map(
        (line) =>
          new Paragraph({
            children: [makeRun(line || " ", { code: true })],
            shading: CODE_SHADING,
            spacing: { after: 0 },
          })
      );
    }

    case "blockquote": {
      const nested = (token.tokens || []).flatMap((t) =>
        blockToDocx(t, { ...ctx, quoted: true })
      );
      return nested;
    }

    case "list": {
      const paragraphs = [];
      token.items.forEach((item, index) => {
        const level = ctx.listLevel;
        const itemBlocks = item.tokens || [];
        let firstTextHandled = false;

        for (const block of itemBlocks) {
          if (block.type === "text" || block.type === "paragraph") {
            const runs = inlineToRuns(block.tokens || [{ type: "text", text: block.text }]);
            const paragraphOptions = {
              children: runs,
              spacing: { after: 80 },
              indent: { left: 360 + level * 360 },
            };
            if (!token.ordered) paragraphOptions.bullet = { level };
            else paragraphOptions.children = [listPrefixRun(true, index), ...runs];
            paragraphs.push(new Paragraph(quoteWrap(paragraphOptions, ctx.quoted)));
            firstTextHandled = true;
          } else if (block.type === "list") {
            paragraphs.push(
              ...blockToDocx(block, { ...ctx, listLevel: level + 1 })
            );
          } else {
            paragraphs.push(...blockToDocx(block, ctx));
          }
        }

        if (!firstTextHandled) {
          const paragraphOptions = {
            children: [],
            spacing: { after: 80 },
            indent: { left: 360 + level * 360 },
          };
          if (!token.ordered) paragraphOptions.bullet = { level };
          else paragraphOptions.children = [listPrefixRun(true, index)];
          paragraphs.push(new Paragraph(quoteWrap(paragraphOptions, ctx.quoted)));
        }
      });
      return paragraphs;
    }

    case "table": {
      const headerRow = new TableRow({
        tableHeader: true,
        children: token.header.map(
          (cell) =>
            new TableCell({
              shading: { fill: "E8E8E8" },
              children: [
                new Paragraph({
                  children: inlineToRuns(cell.tokens),
                  alignment: alignFor(cell.align),
                }),
              ],
            })
        ),
      });

      const dataRows = token.rows.map(
        (row) =>
          new TableRow({
            children: row.map(
              (cell) =>
                new TableCell({
                  children: [
                    new Paragraph({
                      children: inlineToRuns(cell.tokens),
                      alignment: alignFor(cell.align),
                    }),
                  ],
                })
            ),
          })
      );

      return [
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [headerRow, ...dataRows],
        }),
        new Paragraph({ text: "", spacing: { after: 160 } }),
      ];
    }

    case "hr":
      return [
        new Paragraph({
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: "999999", space: 1 },
          },
          spacing: { after: 200 },
        }),
      ];

    case "space":
      return [];

    default:
      if (token.tokens) return token.tokens.flatMap((t) => blockToDocx(t, ctx));
      if (token.raw && token.raw.trim()) {
        return [new Paragraph({ children: [makeRun(token.raw, {})] })];
      }
      return [];
  }
}

export async function markdownToDocx(markdownText) {
  const tokens = marked.lexer(markdownText);
  const children = tokens.flatMap((token) => blockToDocx(token));

  const doc = new Document({
    sections: [{ children }],
  });

  return Packer.toBlob(doc);
}
