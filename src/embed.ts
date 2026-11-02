import { createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import './core/default-worker'
import { DocViewer, type DocViewerProps } from './react/DocViewer'

export type MountOptions = Omit<DocViewerProps, 'styleTarget'>

export interface MountedViewer {
  update(next: Partial<MountOptions>): void
  destroy(): void
}

/** Framework-free embedding: renders the viewer into `el`. */
export function mount(el: HTMLElement, options: MountOptions): MountedViewer {
  let root: Root | null = createRoot(el)
  let props = options
  const shadow = el.getRootNode() instanceof ShadowRoot ? (el.getRootNode() as ShadowRoot) : undefined
  const render = () => root?.render(createElement(DocViewer, { ...props, styleTarget: shadow }))
  render()
  return {
    update(next) {
      props = { ...props, ...next }
      render()
    },
    destroy() {
      root?.unmount()
      root = null
    },
  }
}

/**
 * `<office-viewer src="/file.xlsx" style="height:600px"></office-viewer>`
 * Rendered inside a shadow root so host page CSS cannot leak in.
 * Attributes: src, name, editable, toolbar ("false" to hide), thumbnails ("false" to hide);
 * real-time collaboration: collab (hub WebSocket URL), room, user (display name).
 * Property: `file` (File/Blob/ArrayBuffer) for local data.
 */
export class OfficeViewerElement extends HTMLElement {
  static observedAttributes = ['src', 'name', 'toolbar', 'thumbnails', 'editable', 'collab', 'room', 'user']
  private viewer?: MountedViewer
  private _file?: Blob | ArrayBuffer

  get file() {
    return this._file
  }
  set file(v: Blob | ArrayBuffer | undefined) {
    this._file = v
    this.sync()
  }

  connectedCallback() {
    if (this.viewer) return
    const shadow = this.shadowRoot ?? this.attachShadow({ mode: 'open' })
    const host = document.createElement('div')
    host.style.cssText = 'width:100%;height:100%'
    shadow.innerHTML = '<style>:host{display:block;height:480px;position:relative}</style>'
    shadow.appendChild(host)
    this.viewer = mount(host, this.props())
  }

  disconnectedCallback() {
    this.viewer?.destroy()
    this.viewer = undefined
    if (this.shadowRoot) this.shadowRoot.innerHTML = ''
  }

  attributeChangedCallback() {
    this.sync()
  }

  private props(): MountOptions {
    return {
      src: this._file ?? this.getAttribute('src'),
      name: this.getAttribute('name') ?? undefined,
      toolbar: this.getAttribute('toolbar') !== 'false',
      thumbnails: this.getAttribute('thumbnails') !== 'false',
      editable: this.hasAttribute('editable') && this.getAttribute('editable') !== 'false',
      collab:
        this.getAttribute('collab') && this.getAttribute('room')
          ? { transport: this.getAttribute('collab')!, room: this.getAttribute('room')!, user: { name: this.getAttribute('user') || 'Guest' } }
          : undefined,
      onLoad: (client) => this.dispatchEvent(new CustomEvent('load', { detail: { kind: client.kind, name: client.name } })),
      onError: (error) => this.dispatchEvent(new CustomEvent('error', { detail: error })),
    }
  }

  private sync() {
    this.viewer?.update(this.props())
  }
}

export function defineElement(tag = 'office-viewer') {
  if (typeof customElements !== 'undefined' && !customElements.get(tag)) customElements.define(tag, OfficeViewerElement)
}

defineElement()

export { DocViewer }
