import { Extension, Mark } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import TextAlign from '@tiptap/extension-text-align'
import TextStyle from '@tiptap/extension-text-style'
import Color from '@tiptap/extension-color'
import Highlight from '@tiptap/extension-highlight'
import FontFamily from '@tiptap/extension-font-family'
import Subscript from '@tiptap/extension-subscript'
import Superscript from '@tiptap/extension-superscript'
import Table from '@tiptap/extension-table'
import TableRow from '@tiptap/extension-table-row'
import TableHeader from '@tiptap/extension-table-header'
import TableCell from '@tiptap/extension-table-cell'
import Image from '@tiptap/extension-image'
import Link from '@tiptap/extension-link'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    fontSize: {
      setFontSize: (size: string) => ReturnType
      unsetFontSize: () => ReturnType
    }
  }
}

/** Font size stored on the textStyle mark, e.g. "12pt". */
export const FontSize = Extension.create({
  name: 'fontSize',
  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          fontSize: {
            default: null,
            parseHTML: (el) => el.style.fontSize || null,
            // --dw-fs lets reading mode scale authored sizes with the reader's text-size setting
            renderHTML: (attrs) => (attrs.fontSize ? { style: `font-size: ${attrs.fontSize}; --dw-fs: ${attrs.fontSize}` } : {}),
          },
        },
      },
    ]
  },
  addCommands() {
    return {
      setFontSize:
        (size) =>
        ({ chain }) =>
          chain().setMark('textStyle', { fontSize: size }).run(),
      unsetFontSize:
        () =>
        ({ chain }) =>
          chain().setMark('textStyle', { fontSize: null }).removeEmptyTextStyle().run(),
    }
  },
})

/**
 * Document colours are kept exactly as authored in Word. For display they are routed through
 * a CSS variable so the theme can nudge them into a readable lightness range (see style.css).
 */
const ThemedColor = Color.extend({
  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          color: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-color') || el.style.color || null,
            renderHTML: (attrs) =>
              attrs.color ? { 'data-color': attrs.color, style: `--dw-c: ${attrs.color}` } : {},
          },
        },
      },
    ]
  },
})

const CellWithBackground = TableCell.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      backgroundColor: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-bg') || el.style.backgroundColor || null,
        renderHTML: (attrs) => (attrs.backgroundColor ? { 'data-bg': attrs.backgroundColor, style: `--dw-bg: ${attrs.backgroundColor}` } : {}),
      },
    }
  },
})

const SizedImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      width: {
        default: null,
        parseHTML: (el) => Number(el.getAttribute('width')) || null,
        renderHTML: (attrs) => (attrs.width ? { width: attrs.width } : {}),
      },
      height: {
        default: null,
        parseHTML: (el) => Number(el.getAttribute('height')) || null,
        renderHTML: (attrs) => (attrs.height ? { height: attrs.height } : {}),
      },
    }
  },
})

/** Anchors a comment (stored separately, keyed by commentId) to a range of text. Ranges may overlap. */
export const CommentMark = Mark.create({
  name: 'comment',
  inclusive: false,
  excludes: '',
  addAttributes() {
    return { commentId: { default: null, parseHTML: (el) => el.getAttribute('data-comment-id') } }
  },
  parseHTML() {
    return [{ tag: 'span[data-comment-id]' }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'dw-comment', 'data-comment-id': HTMLAttributes.commentId }, 0]
  },
})

/** Remembers which source .docx paragraph a node came from (used to place annotations in the original file). */
const SourceParagraph = Extension.create({
  name: 'sourceParagraph',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: { srcPara: { default: null, rendered: false, keepOnSplit: false } },
      },
    ]
  },
})

/** Highlights made by the reader (as opposed to ones authored in the Word file) are annotations. */
const AnnotatedHighlight = Highlight.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      user: {
        default: null,
        parseHTML: (el) => (el.getAttribute('data-user') ? true : null),
        renderHTML: (attrs) => (attrs.user ? { 'data-user': '1' } : {}),
      },
    }
  },
})

export const extensions = [
  SourceParagraph,
  CommentMark,
  StarterKit,
  Underline,
  TextStyle,
  ThemedColor,
  FontFamily,
  FontSize,
  AnnotatedHighlight.configure({ multicolor: true }),
  Subscript,
  Superscript,
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  Link.configure({ openOnClick: false, autolink: true }),
  SizedImage.configure({ inline: true, allowBase64: true }),
  Table.configure({ resizable: false }),
  TableRow,
  TableHeader,
  CellWithBackground,
]
