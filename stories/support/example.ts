import type { StoryObj } from '@storybook/web-components-vite';

/** The text of every file in `stories/examples`, by path. */
const files = import.meta.glob<string>('../examples/*.ts', { query: '?raw', import: 'default', eager: true });

/** Files that examples import and that are not examples themselves */
const HELPERS = ['grid-from-rows', 'stack'];

function read(name: string): string {
  const text = files[`../examples/${name}.ts`];
  if (text === undefined) throw new Error(`There is no example file stories/examples/${name}.ts`);
  return text.trim();
}

/**
 * Story parameters for the example in `stories/examples/<name>.ts`.
 *
 * The Code panel and the "Show code" blocks get the file as it is, so what a
 * reader sees is what runs. The comment at the top of the file is the story's
 * description. A file that imports a helper (`gridFromRows`, `stack`) gets the
 * helper appended, which keeps every listing complete enough to copy.
 *
 * @param name - File name without its extension
 * @param options - `height`: the height of the example on docs pages, in px
 */
export function example(name: string, options: { height?: number } = {}): NonNullable<StoryObj['parameters']> {
  const source = read(name);
  const comment = /^\/\*\*\n([\s\S]*?)\n \*\//.exec(source);
  const description = comment ? comment[1].replace(/^ \* ?/gm, '') : '';

  let code = source;
  for (const helper of HELPERS) {
    if (source.includes(`./${helper}.js`)) {
      code += `\n\n\n// ─── ${helper}.ts ${'─'.repeat(40)}\n\n${read(helper)}`;
    }
  }

  return {
    docs: {
      description: { story: description },
      source: { code, language: 'ts' },
    },
    ...(options.height ? { frameHeight: options.height } : {}),
  };
}
