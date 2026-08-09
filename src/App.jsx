import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ShadingType } from 'docx'
import { useState, useRef, useEffect, useCallback } from 'react'
import './App.css'

const STORAGE_KEY = 'writing-space:files'
const WIDTH_KEY = 'writing-space:sidebar-width'
const DARK_KEY = 'writing-space:dark'

function makeFile(name = 'Untitled') {
  return { id: crypto.randomUUID(), name, content: '', updatedAt: Date.now() }
}

function wordCount(html) {
  const text = html.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').trim()
  return text ? text.split(/\s+/).length : 0
}

function charCount(html) {
  return html.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').length
}

function snippetOf(html) {
  const text = html.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim()
  return text.slice(0, 90)
}

function formatDate(ts) {
  if (!ts) return ''
  return new Date(ts).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
}

const NOTE_COLORS = ['#646cff', '#41d1ff', '#a78bfa', '#34d399', '#fb923c']
function colorForId(id) {
  let sum = 0
  for (let i = 0; i < id.length; i++) sum += id.charCodeAt(i)
  return NOTE_COLORS[sum % NOTE_COLORS.length]
}

const FONTS = ['Calibri, sans-serif', 'Georgia, serif', 'Arial, sans-serif', '"Times New Roman", serif', '"Courier New", monospace']
const SIZES = [
  { label: '10', value: '2' },
  { label: '12', value: '3' },
  { label: '14', value: '4' },
  { label: '18', value: '5' },
  { label: '24', value: '6' },
  { label: '32', value: '7' },
]

const FONT_SIZE_MAP = { '1': 16, '2': 20, '3': 24, '4': 28, '5': 36, '6': 48, '7': 64 } // legacy size -> half-points

function rgbToHex(rgb) {
  const m = rgb.match(/\d+/g)
  if (!m) return undefined
  return m.slice(0, 3).map(x => Number(x).toString(16).padStart(2, '0')).join('')
}

function buildTextRun(text, fmt) {
  return new TextRun({
    text,
    bold: fmt.bold || undefined,
    italics: fmt.italics || undefined,
    underline: fmt.underline ? {} : undefined,
    strike: fmt.strike || undefined,
    color: fmt.color,
    font: fmt.font,
    size: fmt.size,
    shading: fmt.highlightColor
      ? { fill: fmt.highlightColor, type: ShadingType.CLEAR, color: 'auto' }
      : undefined,
  })
}

function getRuns(node, fmt) {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent ? [buildTextRun(node.textContent, fmt)] : []
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return []

  const tag = node.tagName
  const next = { ...fmt }
  if (tag === 'B' || tag === 'STRONG') next.bold = true
  if (tag === 'I' || tag === 'EM') next.italics = true
  if (tag === 'U') next.underline = true
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') next.strike = true
  if (tag === 'FONT') {
    if (node.getAttribute('face')) next.font = node.getAttribute('face').split(',')[0].replace(/"/g, '').trim()
    if (node.getAttribute('size')) next.size = FONT_SIZE_MAP[node.getAttribute('size')] || next.size
    if (node.getAttribute('color')) next.color = node.getAttribute('color').replace('#', '')
  }
  if (node.style?.color) next.color = rgbToHex(node.style.color)
  if (node.style?.backgroundColor) next.highlightColor = rgbToHex(node.style.backgroundColor)
  if (node.style?.fontFamily) next.font = node.style.fontFamily.split(',')[0].replace(/"/g, '').trim()

  let runs = []
  node.childNodes.forEach(child => { runs = runs.concat(getRuns(child, next)) })
  return runs
}

function alignmentOf(el) {
  const align = el.style?.textAlign
  if (align === 'center') return AlignmentType.CENTER
  if (align === 'right') return AlignmentType.RIGHT
  if (align === 'justify') return AlignmentType.JUSTIFIED
  return AlignmentType.LEFT
}

function blockToParagraph(el) {
  const runs = getRuns(el, {})
  const opts = { children: runs.length ? runs : [new TextRun('')], alignment: alignmentOf(el) }
  if (el.tagName === 'H1') opts.heading = HeadingLevel.HEADING_1
  if (el.tagName === 'H2') opts.heading = HeadingLevel.HEADING_2
  if (el.tagName === 'H3') opts.heading = HeadingLevel.HEADING_3
  return new Paragraph(opts)
}

// Recursively walks a <ul>/<ol>, tracking nesting depth so sub-lists inside
// <li> elements export at the correct indent level instead of flattening.
function flattenList(listEl, level, usesNumberingRef) {
  const paragraphs = []
  const isOrdered = listEl.tagName === 'OL'
  if (isOrdered) usesNumberingRef.value = true

  Array.from(listEl.children).forEach(li => {
    if (li.tagName !== 'LI') return

    const ownContent = li.cloneNode(true)
    ownContent.querySelectorAll('ul, ol').forEach(nested => nested.remove())
    const runs = getRuns(ownContent, {})

    const opts = { children: runs.length ? runs : [new TextRun('')] }
    if (isOrdered) opts.numbering = { reference: 'numbered-list', level: Math.min(level, 8) }
    else opts.bullet = { level: Math.min(level, 8) }
    paragraphs.push(new Paragraph(opts))

    const nestedLists = li.querySelectorAll(':scope > ul, :scope > ol')
    nestedLists.forEach(nested => {
      const child = flattenList(nested, level + 1, usesNumberingRef)
      paragraphs.push(...child.paragraphs)
    })
  })

  return { paragraphs, usesNumbering: usesNumberingRef.value }
}

function buildParagraphs(root) {
  const paragraphs = []
  const usesNumberingRef = { value: false }

  root.childNodes.forEach(node => {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent.trim()) paragraphs.push(new Paragraph({ children: [new TextRun(node.textContent)] }))
      return
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return

    const tag = node.tagName
    if (tag === 'UL' || tag === 'OL') {
      const flushed = flattenList(node, 0, usesNumberingRef)
      paragraphs.push(...flushed.paragraphs)
      return
    }
    if (['P', 'DIV', 'H1', 'H2', 'H3'].includes(tag)) {
      paragraphs.push(blockToParagraph(node))
      return
    }
    const runs = getRuns(node, {})
    if (runs.length) paragraphs.push(new Paragraph({ children: runs }))
  })

  if (!paragraphs.length) paragraphs.push(new Paragraph({ children: [new TextRun('')] }))
  return { paragraphs, usesNumbering: usesNumberingRef.value }
}

function App() {
  const [files, setFiles] = useState(() => {
    const saved = localStorage.getItem(STORAGE_KEY)
    if (saved) {
      const parsed = JSON.parse(saved)
      if (parsed.length) return parsed
    }
    return [makeFile('Untitled')]
  })

  const [darkMode, setDarkMode] = useState(() => {
    const saved = localStorage.getItem(DARK_KEY)
    return saved ? saved === 'true' : false
  })

  const [activeId, setActiveId] = useState(files[0].id)
  const [sidebarWidth, setSidebarWidth] = useState(() => {
    const saved = localStorage.getItem(WIDTH_KEY)
    return saved ? Number(saved) : 240
  })
  const [font, setFont] = useState(FONTS[0])
  const [mobileView, setMobileView] = useState('list') // 'list' | 'editor' — only matters below 720px
  const [moreMenuOpen, setMoreMenuOpen] = useState(false)
  const dragging = useRef(false)
  const editorRef = useRef(null)

  const active = files.find(f => f.id === activeId) ?? files[0]

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(files))
  }, [files])

  useEffect(() => {
    localStorage.setItem(WIDTH_KEY, String(sidebarWidth))
  }, [sidebarWidth])

  useEffect(() => {
    localStorage.setItem(DARK_KEY, String(darkMode))
  }, [darkMode])

  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== active.content) {
      editorRef.current.innerHTML = active.content
    }
  }, [activeId])

  const persistContent = () => {
    if (!editorRef.current) return
    const html = editorRef.current.innerHTML
    setFiles(prev =>
      prev.map(f => (f.id === activeId ? { ...f, content: html, updatedAt: Date.now() } : f))
    )
  }

  const exec = (command, value = null) => {
    editorRef.current?.focus()
    document.execCommand(command, false, value)
    persistContent()
  }

  const renameActive = (name) => {
    setFiles(prev => prev.map(f => (f.id === activeId ? { ...f, name } : f)))
  }

  const openFile = (id) => {
    setActiveId(id)
    setMobileView('editor')
  }

  const addFile = () => {
    const file = makeFile('Untitled')
    setFiles(prev => [file, ...prev])
    setActiveId(file.id)
    setMobileView('editor')
  }

  const deleteFile = (id) => {
    setFiles(prev => {
      const next = prev.filter(f => f.id !== id)
      const remaining = next.length ? next : [makeFile('Untitled')]
      if (id === activeId) setActiveId(remaining[0].id)
      return remaining
    })
    if (id === activeId) setMobileView('list')
  }

  const goBack = () => setMobileView('list')

  const downloadTxt = () => {
    const text = editorRef.current?.innerText ?? ''
    const blob = new Blob([text], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${active.name || 'Untitled'}.txt`
    a.click()
    URL.revokeObjectURL(url)
  }

  const exportDocx = async () => {
    if (!editorRef.current) return
    const { paragraphs, usesNumbering } = buildParagraphs(editorRef.current)

    const doc = new Document({
      numbering: usesNumbering
        ? {
            config: [{
              reference: 'numbered-list',
              levels: [
                { level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                { level: 1, format: 'lowerLetter', text: '%2.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
                { level: 2, format: 'lowerRoman', text: '%3.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: 2160, hanging: 360 } } } },
              ],
            }],
          }
        : undefined,
      sections: [{
        properties: {},
        children: [
          new Paragraph({ text: active.name || 'Untitled', heading: HeadingLevel.TITLE }),
          ...paragraphs,
        ],
      }],
    })

    const blob = await Packer.toBlob(doc)
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${active.name || 'Untitled'}.docx`
    a.click()
    URL.revokeObjectURL(url)
  }

  const onDragStart = () => {
    dragging.current = true
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
  }

  const onDrag = useCallback((e) => {
    if (!dragging.current) return
    setSidebarWidth(Math.min(400, Math.max(180, e.clientX)))
  }, [])

  const onDragEnd = useCallback(() => {
    dragging.current = false
    document.body.style.cursor = ''
    document.body.style.userSelect = ''
  }, [])

  useEffect(() => {
    window.addEventListener('mousemove', onDrag)
    window.addEventListener('mouseup', onDragEnd)
    return () => {
      window.removeEventListener('mousemove', onDrag)
      window.removeEventListener('mouseup', onDragEnd)
    }
  }, [onDrag, onDragEnd])

  return (
    <div className={darkMode ? 'app dark' : 'app'} data-mobile-view={mobileView}>
      {/* Desktop ribbon (hidden on phones) */}
      <header className="ribbon ribbon-desktop">
        <div className="ribbon-group">
          <button className="ribbon-btn" onClick={() => exec('undo')} title="Undo">↺</button>
          <button className="ribbon-btn" onClick={() => exec('redo')} title="Redo">↻</button>
        </div>

        <div className="ribbon-group">
          <select
            className="ribbon-select"
            value={font}
            onChange={(e) => { setFont(e.target.value); exec('fontName', e.target.value) }}
          >
            {FONTS.map(f => <option key={f} value={f}>{f.split(',')[0].replace(/"/g, '')}</option>)}
          </select>
          <select className="ribbon-select ribbon-select-narrow" onChange={(e) => exec('fontSize', e.target.value)} defaultValue="3">
            {SIZES.map(s => <option key={s.value} value={s.value}>{s.label}</option>)}
          </select>
        </div>

        <div className="ribbon-group">
          <button className="ribbon-btn" onClick={() => exec('bold')} title="Bold"><b>B</b></button>
          <button className="ribbon-btn" onClick={() => exec('italic')} title="Italic"><i>I</i></button>
          <button className="ribbon-btn" onClick={() => exec('underline')} title="Underline"><u>U</u></button>
          <button className="ribbon-btn" onClick={() => exec('strikeThrough')} title="Strikethrough"><s>S</s></button>
        </div>

        <div className="ribbon-group">
          <select className="ribbon-select" onChange={(e) => exec('formatBlock', e.target.value)} defaultValue="p">
            <option value="p">Paragraph</option>
            <option value="h1">Heading 1</option>
            <option value="h2">Heading 2</option>
            <option value="h3">Heading 3</option>
          </select>
        </div>

        <div className="ribbon-group">
          <button className="ribbon-btn" onClick={() => exec('justifyLeft')} title="Align left">⯇</button>
          <button className="ribbon-btn" onClick={() => exec('justifyCenter')} title="Align center">≡</button>
          <button className="ribbon-btn" onClick={() => exec('justifyRight')} title="Align right">⯈</button>
          <button className="ribbon-btn" onClick={() => exec('justifyFull')} title="Justify">☰</button>
        </div>

        <div className="ribbon-group">
          <button className="ribbon-btn" onClick={() => exec('insertUnorderedList')} title="Bullet list">•≡</button>
          <button className="ribbon-btn" onClick={() => exec('insertOrderedList')} title="Numbered list">1≡</button>
        </div>

        <div className="ribbon-group">
          <label className="ribbon-color" title="Text color">
            A
            <input type="color" defaultValue="#e8e8ea" onChange={(e) => exec('foreColor', e.target.value)} />
          </label>
          <label className="ribbon-color" title="Highlight">
            ▧
            <input type="color" defaultValue="#fff176" onChange={(e) => exec('hiliteColor', e.target.value)} />
          </label>
          <button className="ribbon-btn" onClick={() => exec('removeFormat')} title="Clear formatting">Tx</button>
        </div>

        <div className="ribbon-group ribbon-group-end">
          <button className="ribbon-btn" onClick={() => window.print()} title="Print">🖶</button>
          <button className="ribbon-btn" onClick={downloadTxt} title="Download as .txt">⇩</button>
          <button className="ribbon-btn" onClick={exportDocx} title="Export as .docx">.docx</button>
          <button
            className="ribbon-btn"
            onClick={() => setDarkMode(d => !d)}
            title={darkMode ? 'Switch to light mode' : 'Switch to dark mode'}
          >
            {darkMode ? '☀' : '☾'}
          </button>
        </div>
      </header>

      {/* Mobile top app bar (Samsung Notes style) */}
      <header className="mobile-topbar">
        {mobileView === 'editor' ? (
          <>
            <button className="icon-btn" onClick={goBack} aria-label="Back to notes">←</button>
            <input
              className="mobile-topbar-title"
              value={active.name}
              onChange={(e) => renameActive(e.target.value)}
              placeholder="Untitled"
              aria-label="Document title"
            />
            <button className="icon-btn" onClick={() => deleteFile(active.id)} aria-label="Delete note">🗑</button>
            <div className="mobile-more-wrap">
              <button className="icon-btn" onClick={() => setMoreMenuOpen(o => !o)} aria-label="More options">⋮</button>
              {moreMenuOpen && (
                <div className="mobile-more-menu">
                  <button onClick={() => { exportDocx(); setMoreMenuOpen(false) }}>Export .docx</button>
                  <button onClick={() => { downloadTxt(); setMoreMenuOpen(false) }}>Export .txt</button>
                  <button onClick={() => { window.print(); setMoreMenuOpen(false) }}>Print</button>
                  <button onClick={() => { setDarkMode(d => !d); setMoreMenuOpen(false) }}>
                    {darkMode ? 'Light mode' : 'Dark mode'}
                  </button>
                </div>
              )}
            </div>
          </>
        ) : (
          <>
            <h1 className="mobile-topbar-heading">Notes</h1>
            <button
              className="icon-btn"
              onClick={() => setDarkMode(d => !d)}
              aria-label={darkMode ? 'Switch to light mode' : 'Switch to dark mode'}
            >
              {darkMode ? '☀' : '☾'}
            </button>
          </>
        )}
      </header>

      <main className="page-split">
        <aside className="panel panel-left" style={{ width: sidebarWidth }}>
          <div className="panel-head">
            <h2>Library</h2>
            <button className="add-btn" onClick={addFile} aria-label="New note" title="New note">+</button>
          </div>
          <p className="panel-sub">{files.length} note{files.length === 1 ? '' : 's'}</p>
          <ul className="file-list">
            {files.map(f => (
              <li
                key={f.id}
                className={f.id === activeId ? 'file-item active' : 'file-item'}
                style={{ '--note-accent': colorForId(f.id) }}
              >
                <button className="file-item-btn" onClick={() => openFile(f.id)}>
                  <span className="file-item-name">{f.name || 'Untitled'}</span>
                  <span className="file-item-snippet">{snippetOf(f.content) || 'No additional text'}</span>
                  <span className="file-item-meta">{formatDate(f.updatedAt)} · {wordCount(f.content)} words</span>
                </button>
                <button
                  className="file-item-delete"
                  onClick={() => deleteFile(f.id)}
                  aria-label={`Delete ${f.name}`}
                  title="Delete"
                >
                  ×
                </button>
              </li>
            ))}
          </ul>
        </aside>

        <button className="fab" onClick={addFile} aria-label="New note">+</button>

        <div className="drag-handle" onMouseDown={onDragStart} role="separator" aria-orientation="vertical" />

        <section className="canvas">
          <div className="canvas-page">
            <input
              className="canvas-title"
              value={active.name}
              onChange={(e) => renameActive(e.target.value)}
              placeholder="Untitled document"
              aria-label="Document title"
            />
            <div
              ref={editorRef}
              className="canvas-input"
              contentEditable
              suppressContentEditableWarning
              onInput={persistContent}
              style={{ fontFamily: font }}
              aria-label="Document body"
            />
          </div>
          <div className="canvas-status">
            <span>{wordCount(active.content)} words</span>
            <span>{charCount(active.content)} characters</span>
            <span>Saved</span>
          </div>
        </section>
      </main>

      {/* Mobile bottom formatting toolbar (Samsung Notes style) */}
      <div className="mobile-bottombar">
        <button className="ribbon-btn" onClick={() => exec('undo')} title="Undo">↺</button>
        <button className="ribbon-btn" onClick={() => exec('redo')} title="Redo">↻</button>
        <button className="ribbon-btn" onClick={() => exec('bold')} title="Bold"><b>B</b></button>
        <button className="ribbon-btn" onClick={() => exec('italic')} title="Italic"><i>I</i></button>
        <button className="ribbon-btn" onClick={() => exec('underline')} title="Underline"><u>U</u></button>
        <button className="ribbon-btn" onClick={() => exec('strikeThrough')} title="Strikethrough"><s>S</s></button>
        <button className="ribbon-btn" onClick={() => exec('insertUnorderedList')} title="Bullet list">•≡</button>
        <button className="ribbon-btn" onClick={() => exec('insertOrderedList')} title="Numbered list">1≡</button>
        <label className="ribbon-color" title="Text color">
          A
          <input type="color" defaultValue="#e8e8ea" onChange={(e) => exec('foreColor', e.target.value)} />
        </label>
        <label className="ribbon-color" title="Highlight">
          ▧
          <input type="color" defaultValue="#fff176" onChange={(e) => exec('hiliteColor', e.target.value)} />
        </label>
      </div>
    </div>
  )
}

export default App