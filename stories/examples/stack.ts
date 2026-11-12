/**
 * Puts elements under each other in a box that fills its parent. The grid
 * gets the height the others leave.
 */
export function stack(...elements: HTMLElement[]): HTMLElement {
  const box = document.createElement('div');
  box.style.cssText = 'display: flex; flex-direction: column; height: 100%;';
  for (const element of elements) {
    if (element.localName === 'y11n-spreadsheet') {
      element.style.flex = '1';
      element.style.minHeight = '0';
    }
  }
  box.append(...elements);
  return box;
}
