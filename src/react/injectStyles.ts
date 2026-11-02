import css from './styles.css?inline'

const ID = 'docs-engine-styles'

/** Injects the viewer stylesheet once into a document or shadow root. */
export function injectStyles(target: Document | ShadowRoot | undefined = typeof document !== 'undefined' ? document : undefined) {
  if (!target) return
  const host = target instanceof Document ? target.head : target
  if (host.querySelector?.(`style#${ID}`)) return
  const el = (target instanceof Document ? target : target.ownerDocument).createElement('style')
  el.id = ID
  el.textContent = css
  host.appendChild(el)
}

export { css as stylesText }
