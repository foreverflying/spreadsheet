import React, { Dispatch, useEffect, useRef, useState } from 'react'
import { FC } from 'react'
import { Provider, useDispatch, useSelector } from 'react-redux'
import { createStore } from 'redux'
import './App.css'

enum FontStyleBit {
  bold = 1,
  italics = 2,
  underline = 4
}

enum ActionType {
  changeCurrentGrid,
  updateGridText,
  updateGridFontStyle,
  reCalculateGrid,
  rebuildPage,
}

type FormulaData = {
  text: string,
  formula: string,
  value: number,
  error: string,
  dependOn: KeyMap,
}

type GridData = {
  value: string,
  fontMask: number,
  formula?: FormulaData,
}

type Position = {
  column: number,
  row: number
}

type GridMap = {
  [key: string]: GridData
}

type PageState = {
  current: Position,
  gridMap: GridMap
}

type PageAction = (
  { type: ActionType.changeCurrentGrid, column: number, row: number } |
  { type: ActionType.updateGridText, text: string } |
  { type: ActionType.updateGridFontStyle, fontBit: FontStyleBit, checked: boolean } |
  { type: ActionType.reCalculateGrid, column: number, row: number }
)

type KeyMap = {
  [key: string]: true
}

type GridDependenciesMap = {
  [key: string]: KeyMap,
}

const gridDependenciesMap: GridDependenciesMap = {}
const numberChecker = /^[+-]?\d+(\.\d+)?$/

const __f_SUM = (...args: any[]) => {
  return args.reduce((s, n) => s + n)
}

const __f_MIN = (...args: any[]) => {
  return Math.min(...args)
}

const __f_MAX = (...args: any[]) => {
  return Math.max(...args)
}

const calculateFormulaValue = (gridMap: GridMap, formula: FormulaData) => {
  const __g = (key: string) => {
    var data = gridMap[key]
    return data && numberChecker.test(data.value) ? parseFloat(data.value) : 0
  }

  try {
    formula.value = eval(formula.formula)
    formula.error = ''
  } catch (err: any) {
    formula.value = 0
    formula.error = err.toString()
  }
}

const defaultGridData: GridData = {
  value: '',
  fontMask: 0
}

const getGridByKey = (gridMap: GridMap, key: string) => {
  return gridMap[key] || defaultGridData
}

const getGridAt = (gridMap: GridMap, column: number, row: number): GridData => {
  return getGridByKey(gridMap, `${column}_${row}`)
}

const getColumnName = (x: number): string => {
  var arr = []
  while (true) {
    x--
    var n = x % 26
    arr.unshift(String.fromCharCode(65 + n))    // 65 is ASCII of 'A'
    x = x - n
    if (x === 0) {
      return arr.join('')
    }
    x /= 26
  }
}

const gridKeyToName = (key: string) => {
  const [columnStr, rowStr] = key.split('_')
  return getColumnName(parseInt(columnStr)) + rowStr
}

const getPositionByGridName = (name: string): Position => {
  var column = 0, row = 0
  for (var i = 0; i < name.length; ++i) {
    var n = name.charCodeAt(i)
    if (n >= 65) {
      n -= 65     // 65 is ASCII of 'A'
      column = column * 26 + n + 1
    } else {
      n -= 48     // 48 is ASCII of '0'
      row = row * 10 + n
      if (row === 0) {
        throw `invalid grid name: ${name}`
      }
    }
  }
  return { column, row }
}

const regexLegal = /^[ A-Z0-9\(\)\+\-\*\/%\^:,\.]+$/

function checkFormula(formulaText: string) {
  if (!formulaText || !regexLegal.test(formulaText)) {
    throw 'illegal formular'
  }
  return formulaText
}

const regexColonPair = /([A-Z]+[0-9]+) *: *([A-Z]+[0-9]+)/g

const replaceFormulaColonPairs = (formulaText: string) => {
  regexColonPair.lastIndex = 0
  return formulaText.replace(regexColonPair, function (str, grid1, grid2) {
    let { column: startCol, row: startRow } = getPositionByGridName(grid1)
    let { column: endCol, row: endRow } = getPositionByGridName(grid2)
    if (startCol > endCol) {
      [startCol, endCol] = [endCol, startCol]
    }
    if (startRow > endRow) {
      [startRow, endRow] = [endRow, startRow]
    }
    const arr = []
    while (true) {
      var row = startRow
      while (true) {
        arr.push(`__g('${startCol}_${row}')`)
        if (row === endRow) break
        row++
      }
      if (startCol === endCol) break
      startCol++
    }
    return arr.join()
  })
}

const regexFunctionName = /[A-Z][A-Z0-9_]+ *\(/g

const replaceFormulaFunctionNames = (formulaText: string) => {
  regexFunctionName.lastIndex = 0
  return formulaText.replace(regexFunctionName, '__f_$&')
}

const regexGridName = /[A-Z]+[0-9]+/g

const replaceFormulaGridNames = (formulaText: string, dependOn: KeyMap) => {
  regexGridName.lastIndex = 0
  return formulaText.replace(regexGridName, function (name) {
    const { column, row } = getPositionByGridName(name)
    const key = `${column}_${row}`
    dependOn[key] = true
    return `__g('${key}')`
  })
}

function checkDependencies(gridMap: GridMap, gridKey: string, checkingKey: string, dependOn: KeyMap) {
  for (const key in dependOn) {
    if (key === checkingKey) {
      throw `circular reference - ${gridKeyToName(gridKey)} already refered this grid ${gridKeyToName(checkingKey)}`
    }
    var data = getGridByKey(gridMap, key)
    if (data.formula) {
      checkDependencies(gridMap, key, checkingKey, data.formula.dependOn)
    }
  }
}

const updateGridDependencies = (add: boolean, gridKey: string, dependOn: KeyMap) => {
  for (const dependOnKey in dependOn) {
    let item = gridDependenciesMap[dependOnKey]
    if (!item) {
      item = {}
      gridDependenciesMap[dependOnKey] = item
    }
    if (add) {
      item[gridKey] = true
    } else {
      delete item[gridKey]
    }
  }
}

const buildFormula = (gridMap: GridMap, gridKey: string, formulaText: string): FormulaData => {
  const dependOn = {}
  try {
    let converted = formulaText.substring(1)
    converted = checkFormula(converted)
    converted = replaceFormulaFunctionNames(converted)
    converted = replaceFormulaColonPairs(converted)
    converted = replaceFormulaGridNames(converted, dependOn)
    checkDependencies(gridMap, gridKey, gridKey, dependOn)
    const ret: FormulaData = {
      text: formulaText,
      formula: converted,
      dependOn: dependOn,
      value: 0,
      error: '',
    }
    updateGridDependencies(true, gridKey, ret.dependOn)
    calculateFormulaValue(gridMap, ret)
    return ret
  } catch (err: any) {
    const ret: FormulaData = {
      text: formulaText,
      formula: '',
      dependOn: {},
      value: 0,
      error: err.toString(),
    }
    return ret
  }
}

const calculateFormula = (gridMap: GridMap, column: number, row: number, formulaText: string): FormulaData | undefined => {
  const gridData = getGridAt(gridMap, column, row)
  const old = gridData.formula
  const gridKey = `${column}_${row}`
  if (formulaText[0] !== '=') {
    if (old) {
      updateGridDependencies(false, gridKey, old.dependOn)
    }
    return
  }
  formulaText = formulaText.toUpperCase()
  if (old) {
    if (old.text === formulaText) {
      const ret = { ...old }
      calculateFormulaValue(gridMap, ret)
      return ret.value === old.value && ret.error === old.error ? old : ret
    }
    updateGridDependencies(false, gridKey, old.dependOn)
  }
  const ret = buildFormula(gridMap, gridKey, formulaText)
  return ret
}

const pageReducer = (state: PageState | undefined, action: PageAction): PageState => {
  if (!state) {
    return {
      current: {
        column: 1,
        row: 1,
      },
      gridMap: {},
    }
  }
  switch (action.type) {
    case ActionType.changeCurrentGrid:
      if (state.current.column === action.column && state.current.row === action.row) {
        return state
      }
      return {
        ...state,
        current: {
          column: action.column,
          row: action.row,
        },
      }
    case ActionType.reCalculateGrid:
      const gridMap = state.gridMap
      const gridKey = `${action.column}_${action.row}`
      const grid = getGridByKey(gridMap, gridKey)
      if (!grid.formula) {
        return state
      }
      const formula = calculateFormula(state.gridMap, action.column, action.row, grid.formula.text)
      const value = '' + formula!.value
      if (formula === grid.formula && value === grid.value) {
        return state
      }
      return {
        ...state,
        gridMap: {
          ...gridMap,
          [gridKey]: {
            ...grid,
            formula: formula,
            value: value,
          },
        },
      }
    default:
      break
  }
  const { current, gridMap } = state
  const currentKey = `${current.column}_${current.row}`
  const currentGrid = getGridAt(gridMap, current.column, current.row)
  switch (action.type) {
    case ActionType.updateGridText:
      const formula = calculateFormula(gridMap, current.column, current.row, action.text)
      const value = formula ? '' + formula.value : action.text
      if (formula === currentGrid.formula && value === currentGrid.value) {
        return state
      }
      return {
        ...state,
        gridMap: {
          ...gridMap,
          [currentKey]: {
            ...currentGrid,
            formula: formula,
            value: value,
          },
        },
      }
    case ActionType.updateGridFontStyle:
      const newFontMask = action.checked ? (currentGrid.fontMask | action.fontBit) : (currentGrid.fontMask & ~action.fontBit)
      return {
        ...state,
        gridMap: {
          ...gridMap,
          [currentKey]: {
            ...currentGrid,
            fontMask: newFontMask,
          },
        },
      }
    default:
      return state
  }
}

const store = createStore(pageReducer, (window as any).__REDUX_DEVTOOLS_EXTENSION__?.())
// const store = createStore(pageReducer)

const getGridHtmlText = (data: GridData): string => {
  const { value, fontMask, formula } = data
  if (!value) {
    return ''
  }
  let text
  if (formula) {
    if (formula.error) {
      text = `<span class="error">Error</span>`
    } else {
      text = `<span class="formula">${value}</span>`
    }
  } else {
    text = `<span>${value}</span>`
  }
  if (fontMask & FontStyleBit.bold) {
    text = `<b>${text}</b>`
  }
  if (fontMask & FontStyleBit.italics) {
    text = `<i>${text}</i>`
  }
  if (fontMask & FontStyleBit.underline) {
    text = `<u>${text}</u>`
  }
  return text
}

let lastScrollPosition = {
  scrollTop: 0,
  scrollLeft: 0,
}

const onScroll = (e: React.UIEvent<HTMLElement>) => {
  const { scrollTop, scrollLeft } = e.currentTarget
  if (lastScrollPosition.scrollTop !== scrollTop) {
    lastScrollPosition.scrollTop = scrollTop
    document.getElementById('RowNumberBar')!.scrollTop = scrollTop
  }
  if (lastScrollPosition.scrollLeft !== scrollLeft) {
    lastScrollPosition.scrollLeft = scrollLeft
    document.getElementById('SheetHeader')!.scrollLeft = scrollLeft
  }
}

const GridName: FC<{ value: string }> = (prop) => {
  const { value } = prop
  return <input id='GridName' readOnly={true} value={value} />
}

const GridValueInput: FC<{ text: string, current: Position, maxPos: Position }> = ({ text, current, maxPos }) => {
  const [editing, setEditing] = useState<string>()
  const [lastCurrent, setLastCurrent] = useState<Position>(current)
  const dispatch = useDispatch<Dispatch<PageAction>>()
  const onChange = (e: React.FormEvent<HTMLInputElement>) => {
    setEditing(e.currentTarget.value)
  }
  const onBlur = (e: React.FormEvent<HTMLInputElement>) => {
    if (editing !== undefined) {
      dispatch({ type: ActionType.updateGridText, text: editing })
      setEditing(undefined)
    }
  }
  const { column, row } = current
  const onKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    const action: PageAction = { type: ActionType.changeCurrentGrid, column, row }
    switch (e.key) {
      case 'ArrowLeft':
        action.column = column === 1 ? column : column - 1
        break
      case 'ArrowUp':
        action.row = row === 1 ? row : row - 1
        break
      case 'ArrowRight':
        action.column = column === maxPos.column ? column : column + 1
        break
      case 'ArrowDown':
      case 'Enter':
        action.row = row === maxPos.row ? row : row + 1
        break
      default:
        return
    }
    if (e.key !== 'Enter' && !e.shiftKey) {
      return
    }
    e.preventDefault()
    e.stopPropagation()
    if (editing !== undefined) {
      dispatch({ type: ActionType.updateGridText, text: editing })
      setEditing(undefined)
    }
    dispatch(action)
  }
  const ref = useRef<HTMLInputElement>(null)
  useEffect(() => {
    ref.current?.focus()
  }, [current])

  const value = editing !== undefined ? editing : text
  return <input id='GridValue' type='text' value={value} ref={ref} onKeyDown={onKeyDown} onChange={onChange} onBlur={onBlur} />
}

const FontStyleCheck: FC<{ htmlText: string, fontMask: number, fontBit: FontStyleBit }> = (prop) => {
  const { htmlText, fontMask, fontBit } = prop
  const checked = !!(fontMask & fontBit)
  const id = `font-style-${fontBit}`
  const dispatch = useDispatch<Dispatch<PageAction>>()
  const onChange = () => {
    dispatch({ type: ActionType.updateGridFontStyle, fontBit, checked: !checked })
  }
  return <>
    <input id={id} type='checkbox' checked={checked} onChange={onChange} />
    <label htmlFor={id} dangerouslySetInnerHTML={{ __html: htmlText }} />
  </>
}

const ErrorMessage: FC<{ errorText: string }> = ({ errorText }) => {
  return <div id='ErrorMessage'>
    <span className='error'>{errorText}</span>
  </div>
}

const SheetToolbar: FC<{ columnCount: number, rowCount: number }> = ({ columnCount, rowCount }) => {
  const current = useSelector((state: PageState) => state.current)
  const gridData = useSelector((state: PageState) => getGridAt(state.gridMap, current.column, current.row))
  const { formula, value, fontMask } = gridData
  const currentColumnName = getColumnName(current.column)
  return <>
    <div id='SheetToolbar' className='flexRow'>
      <GridName value={`${currentColumnName}${current.row}`} />
      <GridValueInput text={formula ? formula.text : value} current={current} maxPos={{ column: columnCount, row: rowCount }} />
      <FontStyleCheck htmlText='<b>bold</b>' fontMask={fontMask} fontBit={FontStyleBit.bold} />
      <FontStyleCheck htmlText='<i>italics</i>' fontMask={fontMask} fontBit={FontStyleBit.italics} />
      <FontStyleCheck htmlText='<u>underline</u>' fontMask={fontMask} fontBit={FontStyleBit.underline} />
    </div>
    <ErrorMessage errorText={gridData.formula?.error || ''} />
  </>
}

const SheetHeader: FC<{ columnCount: number }> = ({ columnCount }) => {
  const column = useSelector((state: PageState) => state.current.column)
  const headerGrids = []
  for (let i = 1; i <= columnCount; i++) {
    const columnName = getColumnName(i)
    const className = i === column ? 'current' : undefined
    headerGrids.push(
      <div key={`cn${i}`} className={className}>
        {columnName}
      </div>
    )
  }
  return <div className='flexRow fitFlexRowToParent'>
    <div id='SheetHeader' className='flexRow childrenShiftLeft'>
      {headerGrids}
      <div key='h_p' className='corner' />
    </div>
  </div>
}

const RowNumberColumn: FC<{ rowCount: number }> = ({ rowCount }) => {
  const row = useSelector((state: PageState) => state.current.row)
  const rowGrids = []
  for (let i = 1; i <= rowCount; i++) {
    const className = i === row ? 'current' : undefined
    rowGrids.push(
      <div key={`rn${i}`} className={className}>{i}</div>
    )
  }
  return <div id='RowNumberBar' className='childrenShiftUp'>
    {rowGrids}
    <div key='r_p' className='corner' />
  </div>
}

const SheetGrid: FC<{ column: number, row: number, current: boolean }> = ({ column, row, current }) => {
  const gridData = useSelector((state: PageState) => getGridAt(state.gridMap, column, row))
  const dispatch = useDispatch<Dispatch<PageAction>>()
  const onClick = () => {
    dispatch({ type: ActionType.changeCurrentGrid, column, row })
  }
  useEffect(() => {
    const dependedBy = gridDependenciesMap[`${column}_${row}`]
    if (dependedBy) {
      for (const key in dependedBy) {
        const [columnStr, rowStr] = key.split('_')
        dispatch({ type: ActionType.reCalculateGrid, column: parseInt(columnStr), row: parseInt(rowStr) })
      }
    }
  }, [gridData.value])
  const className = current ? 'current' : undefined
  const htmlText = getGridHtmlText(gridData)
  return <div className={className} onClick={onClick} dangerouslySetInnerHTML={{ __html: htmlText }} />
}

const SheetContent: FC<{ columnCount: number, rowCount: number }> = ({ columnCount, rowCount }) => {
  const current = useSelector((state: PageState) => state.current)
  const { column, row } = current
  const gridLines = []
  for (let i = 1; i <= rowCount; i++) {
    const gridLine = []
    for (let j = 1; j <= columnCount; j++) {
      gridLine.push(
        <SheetGrid key={`${j}_${i}`} column={j} row={i} current={j === column && i === row} />
      )
    }
    gridLines.push(
      <div key={`r_${i}`} className='flexRow childrenShiftLeft'>
        {gridLine}
      </div>
    )
  }
  return <div id='SheetContent' onScroll={onScroll}>
    <div className='childrenShiftUp'>
      {gridLines}
    </div>
  </div>
}

const SheetBody: FC<{ columnCount: number, rowCount: number }> = (prop) => {
  const { columnCount, rowCount } = prop
  return <div id='SheetBody' className='flexRow fitFlexRowToParent childrenShiftLeft'>
    <div className='flexColumn fitFlexColumnToParent childrenShiftUp'>
      <div className='corner' />
      <RowNumberColumn rowCount={rowCount} />
    </div>
    <div className='flexColumn fitFlexColumnToParent childrenShiftUp'>
      <SheetHeader columnCount={columnCount} />
      <SheetContent columnCount={columnCount} rowCount={rowCount} />
    </div>
  </div>
}

const App: FC<{ columnCount: number, rowCount: number }> = (prop) => {
  const { columnCount, rowCount } = prop
  return (
    <Provider store={store} >
      <div id='App' className='flexColumn fitFlexColumnToParent'>
        <SheetToolbar columnCount={columnCount} rowCount={rowCount} />
        <SheetBody columnCount={columnCount} rowCount={rowCount} />
      </div>
    </Provider>
  )
}

export default App
