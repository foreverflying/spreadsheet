
var _sheetDataMap
var _focusX, _focusY
var _rowCount, _columnCount
var _gridName, _gridValue, _formatBold, _formatItalics, _formatUnderline, _errorInfo
var _sideBar, _rowNumber, _sheetHeader, _sheetContent

function __getGridValue(key) {
    var data = _sheetDataMap.get(key)
    return !data || !data.value || isNaN(data.value) ? 0 : parseFloat(data.value)
}

function __f_SUM(...args) {
    return args.reduce((s, n) => s + n)
}

function __f_MIN(...args) {
    return Math.min(...args)
}

function __f_MAX(...args) {
    return Math.max(...args)
}

function getColumnName(x) {
    var arr = []
    while(true) {
        x--
        var n = x % 26
        arr.unshift(String.fromCharCode(65 + n))    // 65 is ASCII of 'A'
        x = x - n
        if (x === 0) {
            return arr.join("")
        }
        x /= 26
    }
}

function getGridKey(x, y) {
    return x + "_" + y
}

function getGridName(x, y) {
    return getColumnName(x) + y
}

function getGridValue(x, y) {
    var data = _sheetDataMap.get(getGridKey(x, y))
    return data ? data.value : ""
}

function getGridData(x, y) {
    var data = _sheetDataMap.get(getGridKey(x, y))
    return data ? data : {}
}

function getRowNumberGridByY(y) {
    return $("#r" + y)
}

function getColumnGridByX(x) {
    return $("#c" + x)
}

function getPositionFromGridName(name) {
    var x = 0, y = 0
    for (var i = 0; i < name.length; ++i) {
        var n = name.charCodeAt(i)
        if (n >= 65) {
            n -= 65     // 65 is ASCII of 'A'
            x = x * 26 + n + 1
        } else {
            n -= 48     // 48 is ASCII of '0'
            y = y * 10 + n
        }
    }
    return x < _columnCount && y < _rowCount ? { x: x, y: y } : null
}

function getDomGridByPosition(x, y) {
    return _sheetContent.children().eq(y - 1).children().eq(x - 1)
}

function ensureGetGridData(key) {
    var data = _sheetDataMap.get(key)
    if (!data) {
        data = {}
        _sheetDataMap.set(key, data)
    }
    return data
}

function updateGridValue(x, y, val, bold, italics, underline) {
    var key = getGridKey(x, y)
    var data = ensureGetGridData(key)
    data.bold = bold
    data.italics = italics
    data.underline = underline
    if (data.formulaStr) {
        if (data.formulaStr === val) {
            return data.value
        }
        clearFormulaGridData(key, data)
    } else if (data.value === val) {
        return data.value
    }
    data.value = val
    if (val[0] === "=") {
        data.formulaStr = val.toUpperCase()
        val = computeFormulaGridValue(key)
    }
    notifyGridDataUpdated(data)
    return val
}

function clearFormulaGridData(selfKey, selfData) {
    selfData.formulaStr = undefined
    selfData.formulaErr = undefined
    selfData.formula = undefined
    if (selfData.refGridKeyMap) {
        for (var key of selfData.refGridKeyMap.keys()) {
            var data = ensureGetGridData(key)
            data.listenerMap.delete(selfKey)
        }
    }
}

function listenGridUpdateNotify(selfKey, selfData) {
    for (var key of selfData.refGridKeyMap.keys()) {
        var data = ensureGetGridData(key)
        if (!data.listenerMap) {
            data.listenerMap = new Map()
        }
        data.listenerMap.set(selfKey, true)
    }
}

function notifyGridDataUpdated(data) {
    if (!data.listenerMap) {
        return
    }
    for (var key of data.listenerMap.keys()) {
        var val = computeFormulaGridValue(key)
        var arr = key.split("_")
        var grid = getDomGridByPosition(arr[0], arr[1])
        grid.text(val)
    }
}

function computeFormulaGridValue(key) {
    var data = ensureGetGridData(key)
    if (!data.formula) {
        parseFormulaGrid(key, data)
    }
    if (!data.formulaErr) {
        try {
            data.value = eval(data.formula)
        } catch (err) {
            data.formulaErr = "run time error"
            data.formula = "_"
            data.value = "<b style='color:red;'>error</b>"
        }
    }
    return data.value
}

const regexLegal = /^[ A-Z0-9\(\)\+\-\*\/%\^:,\.]+$/
const regexGridName = /[A-Z]+[0-9]+/g
const regexColonPair = /([A-Z]+[0-9]+) *: *([A-Z]+[0-9]+)/g
const regexFunctionName = /[A-Z][A-Z0-9_]+\(/g

function checkFormula(formula) {
    if (!formula || !regexLegal.test(formula)) {
        throw "illegal formular"
    }
    return formula
}

function replaceFormulaColonPairs(formula) {
    regexColonPair.lastIndex = 0
    return formula.replace(regexColonPair, function(str, grid1, grid2) {
        var p1 = getPositionFromGridName(grid1)
        var p2 = getPositionFromGridName(grid2)
        if (!p1 || !p2) {
            throw "invalid grid name: " + !p1 ? grid1 : grid2
        }
        var smallX = p1.x < p2.x ? p1.x : p2.x
        var bigX = p1.x < p2.x ? p2.x : p1.x
        var smallY = p1.y < p2.y ? p1.y : p2.y
        var bigY = p1.y < p2.y ? p2.y : p1.y
        var arr = []
        while (true) {
            var tmpY = smallY
            while (true) {
                arr.push(getGridName(smallX, tmpY))
                if (tmpY === bigY) break
                tmpY++
            }
            if (smallX === bigX) break
            smallX++
        }
        return arr.join()
    })
}

function replaceFormulaFunctionNames(formula) {
    regexFunctionName.lastIndex = 0
    return formula.replace(regexFunctionName, "__f_$&")
}

function replaceFormulaGridNames(formula, refGridKeyMap) {
    regexGridName.lastIndex = 0
    return formula.replace(regexGridName, function(str) {
        var p = getPositionFromGridName(str)
        if (!p) {
            throw "invalid grid name: " + str
        }
        var key = getGridKey(p.x, p.y)
        refGridKeyMap.set(key, true)
        return "__getGridValue(\"" + key + "\")"
    })
}

function checkFormulaRef(selfKey, checkingKey, refGridKeyMap) {
    for (var key of refGridKeyMap.keys()) {
        if (key === selfKey) {
            var arr = checkingKey.split("_")
            throw "circular reference: " + getGridName(arr[0], arr[1]) + " already refered this grid"
        }
        var data = ensureGetGridData(key)
        if (data.refGridKeyMap) {
            checkFormulaRef(selfKey, key, data.refGridKeyMap)
        }
    }
}

function parseFormulaGrid(selfKey, data) {
    var formula = data.formulaStr.substring(1)
    var refGridKeyMap = new Map()
    try {
        formula = checkFormula(formula)
        formula = replaceFormulaColonPairs(formula)
        formula = replaceFormulaFunctionNames(formula)
        formula = replaceFormulaGridNames(formula, refGridKeyMap)
        checkFormulaRef(selfKey, selfKey, refGridKeyMap)
    } catch (err) {
        data.formulaErr = err
        data.formula = "_"
        data.value = "<b style='color:red;'>error</b>"
        return
    }
    data.formula = formula
    data.refGridKeyMap = refGridKeyMap
    listenGridUpdateNotify(selfKey, data)
}

function getFormattedHtml(text, bold, italics, underline) {
    if (text === undefined) {
        return ""
    }
    if (bold) {
        text = "<b>" + text + "</b>"
    }
    if (italics) {
        text = "<i>" + text + "</i>"
    }
    if (underline) {
        text = "<u>" + text + "</u>"
    }
    return text
}

function commitGridEdit() {
    var val = _gridValue.val()
    var bold = _formatBold.prop("checked")
    var italics = _formatItalics.prop("checked")
    var underline = _formatUnderline.prop("checked")
    val = updateGridValue(_focusX, _focusY, val, bold, italics, underline)
    var grid = getDomGridByPosition(_focusX, _focusY)
    grid.html(getFormattedHtml(val, bold, italics, underline))
}

function addGridFocus(x, y) {
    _focusX = x
    _focusY = y
    var focusGrid = getDomGridByPosition(_focusX, _focusY)
    focusGrid.addClass("gridFocus")
    var focusRowNumber = getRowNumberGridByY(_focusY)
    focusRowNumber.addClass("gridFocus")
    var focusColumn = getColumnGridByX(_focusX)
    focusColumn.addClass("gridFocus")
}

function removeGridFocus() {
    var focusColumn = getColumnGridByX(_focusX)
    focusColumn.removeClass("gridFocus")
    var focusRowNumber = getRowNumberGridByY(_focusY)
    focusRowNumber.removeClass("gridFocus")
    var focusGrid = getDomGridByPosition(_focusX, _focusY)
    focusGrid.removeClass("gridFocus")
}

function setGridDisplay(x, y) {
    var name = getGridName(x, y)
    var data = getGridData(x, y)
    _gridName.val(name)
    _gridValue.val(data.formulaStr ? data.formulaStr : data.value)
    _gridValue.focus()
    _formatBold.prop("checked", data.bold ? true : false)
    _formatItalics.prop("checked", data.italics ? true : false)
    _formatUnderline.prop("checked", data.underline ? true : false)
    _errorInfo.html(data.formulaErr ? "Formula Error: " + data.formulaErr : "&nbsp;")
}

function switchFocusGrid(x, y) {
    if (_focusX) {
        removeGridFocus()
        commitGridEdit()
    }
    addGridFocus(x, y)
    setGridDisplay(x, y)
}

function switchNeighborGridUp(x, y) {
    switchFocusGrid(x, y === 1 ? y : y - 1)
}

function switchNeighborGridDown(x, y) {
    switchFocusGrid(x, y === 100 ? y : y + 1)
}

function switchNeighborGridLeft(x, y) {
    switchFocusGrid(x === 1 ? x : x - 1, y)
}

function switchNeighborGridRight(x, y) {
    switchFocusGrid(x === 100 ? x : x + 1, y)
}

function setSheetBodyHeight() {
    var rect = _sheetContent[0].getBoundingClientRect()
    var pageHeight = $(window).height()
    var height = pageHeight - 30 - rect.top
    _sheetContent.css("maxHeight", "" + height + "px")
    _sideBar.css("height", "" + (height + 21) + "px")
}

function initRowNumber() {
    for (var i = 1; i <= _rowCount; ++i) {
        var grid = document.createElement("div")
        grid.id = "r" + i
        grid.innerHTML = i
        _rowNumber.append(grid)
    }
}

function initSheetHeader() {
    for (var i = 1; i <= _columnCount; ++i) {
        var grid = document.createElement("div")
        grid.id = "c" + i
        grid.innerHTML = getColumnName(i)
        _sheetHeader.append(grid)
    }
}

function initSheetContent() {
    for (var y = 1; y <= _rowCount; ++y) {
        var row = document.createElement("div")
        row.className = "gridRow"
        for (var x = 1; x <= _columnCount; ++x) {
            var grid = document.createElement("div")
            grid.xIndex = x
            var data = getGridData(x, y)
            grid.innerHTML = getFormattedHtml(data.value, data.bold, data.italics, data.underline)
            row.appendChild(grid)
        }
        row.yIndex = y
        row.appendChild(grid)
        _sheetContent.append(row)
    }
}

function initFocusGrid() {
    switchFocusGrid(1, 1)
}

function addSheetBodyResizeListener() {
    $(window).resize(function() {
        setSheetBodyHeight()
    })
}

function addSheetContentScrollListener() {
    var scrollTop = _sheetContent.scrollTop()
    var scrollLeft = _sheetContent.scrollLeft()

    _sheetContent.scroll(function onSheetScroll(event) {
        var curScrollTop = _sheetContent.scrollTop()
        var curScrollLeft = _sheetContent.scrollLeft()
        var diffTop = curScrollTop - scrollTop
        var diffLeft = curScrollLeft - scrollLeft
        scrollTop = curScrollTop
        scrollLeft = curScrollLeft
        if (diffLeft !== 0) {
            var offsetLeft = _sheetHeader.offset()
            offsetLeft.left -= diffLeft
            _sheetHeader.offset(offsetLeft)
        }
        if (diffTop !== 0) {
            var offsetTop = _rowNumber.offset()
            offsetTop.top -= diffTop
            _rowNumber.offset(offsetTop)
        }
    })
}

function addSheetContentClickListener() {
    $(".gridRow >div").click(function() {
        var grid = $(this)
        var x = grid[0].xIndex
        var y = grid.parent()[0].yIndex
        commitGridEdit()
        switchFocusGrid(x, y)
    })
}

function addInputListener() {
    _gridValue.keydown(function(event) {
        if (event.keyCode === 37 ) {
            switchNeighborGridLeft(_focusX, _focusY)
        } else if (event.keyCode === 38 ) {
            switchNeighborGridUp(_focusX, _focusY)
        } else if (event.keyCode === 39 ) {
            switchNeighborGridRight(_focusX, _focusY)
        } else if (event.keyCode === 40 || event.keyCode === 13 ) {
            switchNeighborGridDown(_focusX, _focusY)
        }
    })
    var commit = function(event) {
        commitGridEdit()
        setGridDisplay(_focusX, _focusY)
    }
    _gridValue.blur(commit)
    _formatBold.change(commit)
    _formatItalics.change(commit)
    _formatUnderline.change(commit)
}

function addRefreshButtonClickListener() {
    $("#refreshButton").click(function(event) {
        _rowNumber.empty()
        _sheetHeader.empty()
        _sheetContent.empty()
        alert("Sheet already destroyed, click OK to rebuild.")
        initSheetBody(_columnCount, _rowCount, _sheetDataMap)
    })
}

function initSheetBody(columnCount, rowCount, dataMap) {
    _sheetDataMap = dataMap
    _columnCount = columnCount
    _rowCount = rowCount
    _gridName = $("#gridName")
    _gridValue = $("#gridValue")
    _sideBar = $("#sideBar")
    _rowNumber = $("#rowNumber")
    _sheetHeader = $("#sheetHeader")
    _sheetContent = $("#sheetContent")
    _formatBold = $("#formatBold")
    _formatItalics = $("#formatItalics")
    _formatUnderline = $("#formatUnderline")
    _errorInfo = $("#errorInfo")

    setSheetBodyHeight()
    initRowNumber(_rowCount)
    initSheetHeader(_columnCount)
    initSheetContent(_columnCount, _rowCount)
    initFocusGrid()
    addSheetBodyResizeListener(_sheetContent)
    addSheetContentScrollListener()
    addSheetContentClickListener()
    addInputListener()
    addRefreshButtonClickListener()
}

function createSpreadSheet(columnCount, rowCount) {
    initSheetBody(columnCount, rowCount, new Map())
}