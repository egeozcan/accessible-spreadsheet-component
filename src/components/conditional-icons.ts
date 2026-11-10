import { css, html, svg, type TemplateResult } from 'lit';
import type { ConditionalIconSetName } from '../types.js';

/**
 * Icons of the conditional-format icon sets. Each icon is drawn in one color
 * through `currentColor`, so forced-colors mode recolors it, and icons of a
 * set differ in shape, so they stay distinct without color.
 */

const GREEN = '#2e7d32';
const LIGHT_GREEN = '#558b2f';
const AMBER = '#a36a00';
const ORANGE = '#c75300';
const RED = '#c62828';
const BLUE = '#1565c0';

const ARROW = 'M8 1.5 13.5 7.5H10V14.5H6V7.5H2.5Z';

function arrow(degrees: number) {
  return svg`<path d="${ARROW}" fill="currentColor" transform="rotate(${degrees} 8 8)"/>`;
}

function bars(filled: number) {
  const heights = [4, 7, 10, 13];
  return heights.map(
    (h, i) => svg`<rect
      x="${1.5 + i * 3.5}"
      y="${14.5 - h}"
      width="2.5"
      height="${h}"
      fill="${i < filled ? 'currentColor' : 'none'}"
      stroke="currentColor"
      stroke-width="1"
    />`
  );
}

type Icon = [color: string, shape: unknown];

const ICONS: Record<ConditionalIconSetName, Icon[]> = {
  arrows3: [
    [RED, arrow(180)],
    [AMBER, arrow(90)],
    [GREEN, arrow(0)],
  ],
  arrows5: [
    [RED, arrow(180)],
    [ORANGE, arrow(135)],
    [AMBER, arrow(90)],
    [LIGHT_GREEN, arrow(45)],
    [GREEN, arrow(0)],
  ],
  shapes3: [
    [RED, svg`<path d="M8 1.5 14.5 8 8 14.5 1.5 8Z" fill="currentColor"/>`],
    [AMBER, svg`<path d="M8 2 14.5 13.5H1.5Z" fill="currentColor"/>`],
    [GREEN, svg`<circle cx="8" cy="8" r="6" fill="currentColor"/>`],
  ],
  symbols3: [
    [RED, svg`<path d="M4 4 12 12M12 4 4 12" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>`],
    [
      AMBER,
      svg`<path d="M8 2.8V9.2" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"/><circle cx="8" cy="12.8" r="1.5" fill="currentColor"/>`,
    ],
    [
      GREEN,
      svg`<path d="M3 8.5 6.5 12 13 4.5" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>`,
    ],
  ],
  bars5: [0, 1, 2, 3, 4].map((n): Icon => [BLUE, bars(n)]),
};

/**
 * Styles for the icons, to include in the `static styles` of any component
 * that renders them. The UA style sheet exempts `<svg>` from forced colors,
 * so the icons are recolored here to stay visible on any system background.
 */
export const conditionalIconStyles = css`
  .cf-icon-svg {
    flex: 0 0 auto;
    color: var(--cf-icon-color);
  }

  @media (forced-colors: active) {
    .cf-icon-svg {
      color: CanvasText;
    }
  }
`;

/** The icon at `index` (0 = the icon for the lowest numbers) of an icon set, as a decorative SVG. */
export function conditionalIcon(set: ConditionalIconSetName, index: number): TemplateResult {
  const icons = ICONS[set];
  const [color, shape] = icons[Math.max(0, Math.min(index, icons.length - 1))];
  return html`<svg
    class="cf-icon-svg"
    viewBox="0 0 16 16"
    width="14"
    height="14"
    aria-hidden="true"
    focusable="false"
    style="--cf-icon-color: ${color}"
  >${shape}</svg>`;
}
