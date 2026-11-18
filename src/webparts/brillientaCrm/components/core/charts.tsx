import * as React from 'react';

export interface IDatum {
  label: string;
  value: number;
  hint?: string;
  /** Short axis label (full label is shown in the tooltip). */
  short?: string;
}

/** Categorical order (validated palette) – assigned in fixed order, never cycled. */
const SERIES = ['var(--series-1)', 'var(--series-2)', 'var(--series-3)', 'var(--series-4)', 'var(--series-5)', 'var(--series-6)'];

interface ITip { x: number; y: number; text: string }

function useTip(): [ITip | undefined, (e: React.MouseEvent, text: string) => void, () => void, React.RefObject<HTMLDivElement>] {
  const [tip, setTip] = React.useState<ITip>();
  const ref = React.useRef<HTMLDivElement>(null);
  const show = (e: React.MouseEvent, text: string): void => {
    const box = ref.current?.getBoundingClientRect();
    if (box) { setTip({ x: e.clientX - box.left + 12, y: e.clientY - box.top - 8, text }); }
  };
  return [tip, show, () => setTip(undefined), ref];
}

const Tip: React.FC<{ tip?: ITip }> = ({ tip }) => (tip ? <div className="viz-tip" style={{ left: tip.x, top: tip.y }}>{tip.text}</div> : null);

/** Vertical columns, e.g. revenue by month. Single series → single hue. */
export const ColumnChart: React.FC<{ data: IDatum[]; format?: (n: number) => string; height?: number; ariaLabel: string }> =
  ({ data, format = n => n.toLocaleString(), height = 180, ariaLabel }) => {
    const [tip, show, hide, ref] = useTip();
    const max = Math.max(1, ...data.map(d => d.value));
    const w = 100 / Math.max(1, data.length);
    const peak = data.reduce((best, d, i) => (d.value > data[best].value ? i : best), 0);
    return (
      <div className="viz" ref={ref} style={{ position: 'relative' }}>
        <svg width="100%" height={height + 22} role="img" aria-label={ariaLabel} preserveAspectRatio="none" viewBox={`0 0 100 ${height + 22}`}>
          {[0.25, 0.5, 0.75, 1].map(g => (
            <line key={g} x1={0} x2={100} y1={height - g * height + 0.5} y2={height - g * height + 0.5} className="viz-grid" vectorEffect="non-scaling-stroke" />
          ))}
          <line x1={0} x2={100} y1={height} y2={height} className="viz-axis" vectorEffect="non-scaling-stroke" />
          {data.map((d, i) => {
            const h = (d.value / max) * (height - 16);
            const bw = Math.max(0.5, w * 0.62);
            return (
              <g key={d.label} onMouseMove={e => show(e, `${d.label}: ${format(d.value)}`)} onMouseLeave={hide}>
                <rect x={i * w} y={0} width={w} height={height} fill="transparent" />
                {d.value > 0 && <rect x={i * w + (w - bw) / 2} y={height - h} width={bw} height={h} rx={0.6} className="viz-bar" />}
              </g>
            );
          })}
        </svg>
        <div className="viz-xlabels">
          {data.map((d, i) => <span key={d.label} style={{ width: `${w}%` }}>{data.length > 8 && i % 2 ? '' : (d.short || d.label)}</span>)}
        </div>
        {data[peak] && data[peak].value > 0 && (
          <div className="viz-caption">Peak: {data[peak].label} · {format(data[peak].value)}</div>
        )}
        <Tip tip={tip} />
      </div>
    );
  };

/** Horizontal bars with direct labels (ranked categories). */
export const BarList: React.FC<{ data: IDatum[]; format?: (n: number) => string; max?: number; onClick?: (d: IDatum) => void }> =
  ({ data, format = n => n.toLocaleString(), max, onClick }) => {
    const m = max ?? Math.max(1, ...data.map(d => d.value));
    if (!data.length) { return <div className="viz-empty">No data</div>; }
    return (
      <div className="viz viz-barlist">
        {data.map(d => (
          <div key={d.label} className="viz-barrow" title={`${d.label}: ${format(d.value)}${d.hint ? ` · ${d.hint}` : ''}`}
            onClick={onClick ? () => onClick(d) : undefined} style={onClick ? { cursor: 'pointer' } : undefined}>
            <span className="viz-barlabel">{d.label}</span>
            <span className="viz-track"><span className="viz-fill" style={{ width: `${Math.max(d.value > 0 ? 1.5 : 0, (d.value / m) * 100)}%` }} /></span>
            <span className="viz-barvalue">{format(d.value)}</span>
          </div>
        ))}
      </div>
    );
  };

/** Part-to-whole for ≤ 6 categories (rest folded into "Other"), with legend + table-like values. */
export const Donut: React.FC<{ data: IDatum[]; format?: (n: number) => string; ariaLabel: string; centerLabel?: string }> =
  ({ data, format = n => n.toLocaleString(), ariaLabel, centerLabel }) => {
    const [tip, show, hide, ref] = useTip();
    const sorted = [...data].filter(d => d.value > 0).sort((a, b) => b.value - a.value);
    const top = sorted.slice(0, 5);
    const rest = sorted.slice(5).reduce((s, d) => s + d.value, 0);
    const slices = rest > 0 ? [...top, { label: 'Other', value: rest }] : top;
    const total = slices.reduce((s, d) => s + d.value, 0);
    if (!total) { return <div className="viz-empty">No data</div>; }
    const R = 42;
    const C = 2 * Math.PI * R;
    let acc = 0;
    return (
      <div className="viz viz-donut" ref={ref} style={{ position: 'relative' }}>
        <svg width={128} height={128} viewBox="0 0 100 100" role="img" aria-label={ariaLabel}>
          <g transform="rotate(-90 50 50)">
            {slices.map((d, i) => {
              const len = (d.value / total) * C;
              const gap = slices.length > 1 ? Math.min(1.2, len / 2) : 0;
              const el = (
                <circle key={d.label} cx={50} cy={50} r={R} fill="none" strokeWidth={12}
                  stroke={d.label === 'Other' ? 'var(--viz-other)' : SERIES[i]}
                  strokeDasharray={`${Math.max(0, len - gap)} ${C}`} strokeDashoffset={-acc}
                  onMouseMove={e => show(e, `${d.label}: ${format(d.value)} (${Math.round((d.value / total) * 100)}%)`)} onMouseLeave={hide} />
              );
              acc += len;
              return el;
            })}
          </g>
          <text x={50} y={48} textAnchor="middle" className="viz-center">{format(total)}</text>
          <text x={50} y={60} textAnchor="middle" className="viz-center-sub">{centerLabel || 'total'}</text>
        </svg>
        <ul className="viz-legend">
          {slices.map((d, i) => (
            <li key={d.label}>
              <i style={{ background: d.label === 'Other' ? 'var(--viz-other)' : SERIES[i] }} />
              <span className="viz-legend-label">{d.label}</span>
              <span className="viz-legend-value">{format(d.value)}</span>
            </li>
          ))}
        </ul>
        <Tip tip={tip} />
      </div>
    );
  };

/** Ordered stages with conversion rates (ordinal single hue, light → dark). */
export const Funnel: React.FC<{ data: IDatum[] }> = ({ data }) => {
  const max = Math.max(1, ...data.map(d => d.value));
  const steps = ['var(--seq-300)', 'var(--seq-400)', 'var(--seq-500)', 'var(--seq-600)', 'var(--seq-700)'];
  return (
    <div className="viz viz-funnel">
      {data.map((d, i) => {
        const top = data[0] ? data[0].value : 0;
        const rate = i > 0 && top ? Math.round((d.value / top) * 100) : undefined;
        return (
          <div key={d.label} className="viz-funnel-row" title={`${d.label}: ${d.value.toLocaleString()}${rate !== undefined ? ` (${rate}% of ${data[0].label.toLowerCase()})` : ''}`}>
            <span className="viz-barlabel">{d.label}</span>
            <span className="viz-funnel-track">
              <span className="viz-funnel-fill" style={{ width: `${Math.max(d.value ? 2 : 0, (d.value / max) * 100)}%`, background: steps[Math.min(i, steps.length - 1)] }} />
            </span>
            <span className="viz-barvalue">{d.value.toLocaleString()}{rate !== undefined && <small> · {rate}%</small>}</span>
          </div>
        );
      })}
    </div>
  );
};

/** Progress against a target (quota attainment). */
export const Meter: React.FC<{ value: number; target: number; label?: string }> = ({ value, target, label }) => {
  const pct = target > 0 ? Math.round((value / target) * 100) : 0;
  return (
    <div className="viz-meter" title={label}>
      <span className="viz-meter-track"><span className="viz-meter-fill" style={{ width: `${Math.min(100, pct)}%` }} /></span>
      <span className="viz-barvalue">{target > 0 ? `${pct}%` : '—'}</span>
    </div>
  );
};
