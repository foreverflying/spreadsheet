## This project comes from an interview assessment requirement

To realize a sheet page, with the size of 100 * 100, can support basic formula.

# Features

1. Click the grid, to input some values.
2. If the input value is numeric, the grid has this value, or else its value is 0.
3. When losing focus or moving to other grids, the editing of the grid takes effect.
4. Use Shift + Arrow Keys to move to adjacent grids.
5. Enter key can help to move to the lower grid too.
6. Can scroll horizontally or vertically.
7. Can input formula by starting with "=", e.g. "= A1 + A2 - 10".
8. "+ - * / %" are supported in the formula.
9. Supported functions: MIN(...) MAX(...) SUM(...), e.g. "= SUM(100, 200, A1, B1) + MIN(A2, B2)"
10. ":" can be used to represent all the grids in a rectangle, e.g. "= SUM(A1:B2)"
11. The formula grid's text is shown in blue.
12. You'll get a red "error" displayed in the grid when you put some incorrect formula in.
13. Focusing on the error grid, the reason will be shown, e.g. recursive references between grids.

## React re-write version, implemented with React Hooks + Redux + Typescript

Github page link: https://foreverflying.github.io/spreadsheet/

## jQuery version

Github page link: https://foreverflying.github.io/spreadsheet/old.html