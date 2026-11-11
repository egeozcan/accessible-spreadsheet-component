import readme from '../../README.md?raw';

/**
 * A section of the project README, as Markdown: the text under the heading
 * `title`, up to the next heading of the same or a higher level. The
 * reference pages are made of these, so the README stays the one place where
 * the API is described.
 *
 * Links to other README sections become plain text (the sections are separate
 * pages here), and `<details>` blocks are opened up.
 *
 * @param title - Heading text, without its `#` marks
 * @param heading - A heading to put above the section, e.g. `'## Filters'`
 */
export function readmeSection(title: string, heading?: string): string {
  const lines = readme.split('\n');
  const body: string[] = [];
  let level = 0;
  let inFence = false;

  for (const line of lines) {
    if (line.startsWith('```')) inFence = !inFence;
    const match = inFence ? null : /^(#{1,6}) (.+)$/.exec(line);
    if (level === 0) {
      if (match && match[2].trim() === title) level = match[1].length;
      continue;
    }
    if (match && match[1].length <= level) break;
    body.push(line);
  }
  if (level === 0) throw new Error(`The README has no section "${title}"`);

  const text = body
    .join('\n')
    .replace(/\[([^\]]+)\]\(#[^)]*\)/g, '$1')
    .replace(/<\/?details>\n?/g, '')
    .replace(/<summary>(.*?)<\/summary>/g, '**$1**')
    .replace(/\n---\s*$/, '')
    .trim();
  return heading ? `${heading}\n\n${text}` : text;
}
