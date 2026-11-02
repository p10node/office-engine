export type Zoom = number | 'fit'

const STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3]

export function ZoomControl({ zoom, effective, onChange }: { zoom: Zoom; effective: number; onChange: (z: Zoom) => void }) {
  const step = (dir: 1 | -1) => {
    const cur = effective
    const next = dir > 0 ? STEPS.find((s) => s > cur + 0.001) : [...STEPS].reverse().find((s) => s < cur - 0.001)
    if (next) onChange(next)
  }
  return (
    <div className="de-zoom" role="group" aria-label="Zoom">
      <button type="button" aria-label="Zoom out" onClick={() => step(-1)}>
        −
      </button>
      <span title={zoom === 'fit' ? 'Fit width (click for 100%)' : 'Click to fit width'} onClick={() => onChange(zoom === 'fit' ? 1 : 'fit')}>
        {zoom === 'fit' ? 'Fit' : `${Math.round(effective * 100)}%`}
      </span>
      <button type="button" aria-label="Zoom in" onClick={() => step(1)}>
        +
      </button>
    </div>
  )
}
