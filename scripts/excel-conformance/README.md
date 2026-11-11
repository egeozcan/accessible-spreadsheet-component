Recomputes the conformance fixtures in Microsoft Excel for Mac (JXA).
Run from a scratch directory: write `cases.json` (`{file,i,f,grid,rows,cols}` per fixture), run
`osascript -l JavaScript run.js` (needs Excel with a blank workbook open), then `node compare.js`.
Paths in compare.js are absolute to the author's checkout; adjust them. Functions this Excel build
doesn't know come back as #NAME? and are left as unverified (`src: "expected"`).
