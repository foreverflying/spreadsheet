
var _sheetDataMap
var _focusGridName
var _rowCount, _columnCount
var _gridName, _gridValue, _formatBold, _formatItalics, _formatUnderline, _errorInfo
var _sideBar, _rowNumberBar, _sheetHeader, _sheetContent

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

function getGridNameByPos(x, y) {
    return getColumnName(x) + y
}

function getPosByGridName(name) {
    var x = 0, y = 0
    for (var i = 0; i < name.length; ++i) {
        var n = name.charCodeAt(i)
        if (n >= 65) {
            n -= 65     // 65 is ASCII of 'A'
            x = x * 26 + n + 1
        } else {
            n -= 48     // 48 is ASCII of '0'
            y = y * 10 + n
            if (y === 0) {
                return null
            }
        }
    }
    return x <= _columnCount && y <= _rowCount ? [x, y] : null
}

function getGridData(name) {
    var data = _sheetDataMap.get(name)
    return data ? data : {}
}

function getGridDataEnsure(name) {
    var data = _sheetDataMap.get(name)
    if (!data) {
        data = {}
        _sheetDataMap.set(name, data)
    }
    return data
}

function getRowNumberBarGridByY(y) {
    return $("#r" + y)
}

function getSheetHeaderGridByX(x) {
    return $("#s" + x)
}

function getDomGridByName(name) {
    return $("#" + name)
}

function getDomGridByPos(x, y) {
    return $(_sheetContent[0].children[(y - 1) * _columnCount + x - 1])
}

function updateGridValue(name, val, bold, italics, underline) {
    var data = getGridDataEnsure(name)
    data.bold = bold
    data.italics = italics
    data.underline = underline
    if (data.formulaStr) {
        if (data.formulaStr === val) {
            return data
        }
        clearFormulaGridData(name, data)
    } else if (data.value === val) {
        return data
    }
    if (val[0] === "=") {
        data.formulaStr = val.toUpperCase()
        computeFormulaGridValue(name, data)
    } else {
        data.value = val
    }
    notifyGridDataUpdated(data)
    return data
}

function listenGridUpdateNotify(selfName, selfData) {
    for (var name of selfData.refGridKeyMap.keys()) {
        var data = getGridDataEnsure(name)
        if (!data.listenerMap) {
            data.listenerMap = new Map()
        }
        data.listenerMap.set(selfName, true)
    }
}

function notifyGridDataUpdated(data) {
    if (!data.listenerMap) {
        return
    }
    for (var name of data.listenerMap.keys()) {
        var notifyData = getGridDataEnsure(name)
        computeFormulaGridValue(name, notifyData)
        var grid = getDomGridByName(name)
        grid.html(getGridFormattedHtml(notifyData))
        notifyGridDataUpdated(notifyData)
    }
}

function clearFormulaGridData(selfName, selfData) {
    selfData.formulaStr = undefined
    selfData.formulaErr = undefined
    selfData.formula = undefined
    if (selfData.refGridKeyMap) {
        for (var name of selfData.refGridKeyMap.keys()) {
            var data = getGridDataEnsure(name)
            data.listenerMap.delete(selfName)
        }
    }
}

const regexLegal = /^[ A-Z0-9\(\)\+\-\*\/%\^:,\.]+$/

function checkFormula(formula) {
    if (!formula || !regexLegal.test(formula)) {
        throw "illegal formular"
    }
    return formula
}

const regexColonPair = /([A-Z]+[0-9]+) *: *([A-Z]+[0-9]+)/g

function replaceFormulaColonPairs(formula) {
    regexColonPair.lastIndex = 0
    return formula.replace(regexColonPair, function(str, grid1, grid2) {
        var [x1, y1] = getPosByGridName(grid1)
        var [x2, y2] = getPosByGridName(grid2)
        if (!x1 || !x2) {
            throw "invalid grid name: " + !x1 ? grid1 : grid2
        }
        if (x1 > x2) {
            [x1, x2] = [x2, x1]
        }
        if (y1 > y2) {
            [y1, y2] = [y2, y1]
        }
        var arr = []
        while (true) {
            var y = y1
            while (true) {
                arr.push(getGridNameByPos(x1, y))
                if (y === y2) break
                y++
            }
            if (x1 === x2) break
            x1++
        }
        return arr.join()
    })
}

const regexFunctionName = /[A-Z][A-Z0-9_]+ *\(/g

function replaceFormulaFunctionNames(formula) {
    regexFunctionName.lastIndex = 0
    return formula.replace(regexFunctionName, "__f_$&")
}

const regexGridName = /[A-Z]+[0-9]+/g

function replaceFormulaGridNames(formula, refGridKeyMap) {
    regexGridName.lastIndex = 0
    return formula.replace(regexGridName, function(name) {
        if (!getPosByGridName(name)) {
            throw "invalid grid name: " + name
        }
        refGridKeyMap.set(name, true)
        return "__getGridValue(\"" + name + "\")"
    })
}

function checkFormulaRef(selfName, checkingName, refGridKeyMap) {
    for (var name of refGridKeyMap.keys()) {
        if (name === checkingName) {
            throw "circular reference - " + selfName + " already refered this grid " + checkingName
        }
        var data = getGridDataEnsure(name)
        if (data.refGridKeyMap) {
            checkFormulaRef(name, checkingName, data.refGridKeyMap)
        }
    }
}

function parseFormulaGrid(selfName, data) {
    var formula = data.formulaStr.substring(1)
    var refGridKeyMap = new Map()
    try {
        formula = checkFormula(formula)
        formula = replaceFormulaColonPairs(formula)
        formula = replaceFormulaFunctionNames(formula)
        formula = replaceFormulaGridNames(formula, refGridKeyMap)
        checkFormulaRef(selfName, selfName, refGridKeyMap)
    } catch (err) {
        data.formulaErr = err
        data.formula = "_"
        data.value = "error"
        return
    }
    data.formula = formula
    data.refGridKeyMap = refGridKeyMap
    listenGridUpdateNotify(selfName, data)
}

function computeFormulaGridValue(name, data) {
    if (!data.formula) {
        parseFormulaGrid(name, data)
    }
    if (!data.formulaErr) {
        try {
            data.value = eval(data.formula)
        } catch (err) {
            data.formulaErr = "run time error"
            data.formula = "_"
            data.value = "error"
        }
    }
}

function getGridFormattedHtml(data) {
    var val = data.value
    if (val === undefined) {
        return ""
    }
    if (data.formulaStr) {
        var color = data.formulaErr ? "red" : "dodgerblue"
        val = "<span style='color:" + color + ";'>" + val + "</span>"
    }
    if (data.bold) {
        val = "<b>" + val + "</b>"
    }
    if (data.italics) {
        val = "<i>" + val + "</i>"
    }
    if (data.underline) {
        val = "<u>" + val + "</u>"
    }
    return val
}

function commitGridEdit() {
    var val = _gridValue.val()
    var bold = _formatBold.prop("checked")
    var italics = _formatItalics.prop("checked")
    var underline = _formatUnderline.prop("checked")
    var data = updateGridValue(_focusGridName, val, bold, italics, underline)
    var grid = getDomGridByName(_focusGridName)
    grid.html(getGridFormattedHtml(data))
}

function addGridFocus(x, y) {
    var focusGrid = getDomGridByPos(x, y)
    focusGrid.addClass("gridFocus")
    var focusRowNumber = getRowNumberBarGridByY(y)
    focusRowNumber.addClass("gridFocus")
    var focusColumn = getSheetHeaderGridByX(x)
    focusColumn.addClass("gridFocus")
}

function removeGridFocus(x, y) {
    var focusColumn = getSheetHeaderGridByX(x)
    focusColumn.removeClass("gridFocus")
    var focusRowNumber = getRowNumberBarGridByY(y)
    focusRowNumber.removeClass("gridFocus")
    var focusGrid = getDomGridByPos(x, y)
    focusGrid.removeClass("gridFocus")
}

function setGridToDisplay(name) {
    var data = getGridData(name)
    _gridName.val(name)
    _gridValue.val(data.formulaStr ? data.formulaStr : data.value)
    _gridValue.focus()
    _formatBold.prop("checked", data.bold ? true : false)
    _formatItalics.prop("checked", data.italics ? true : false)
    _formatUnderline.prop("checked", data.underline ? true : false)
    _errorInfo.html(data.formulaErr ? "Formula Error: " + data.formulaErr : "&nbsp;")
}

function switchFocusGrid(name, x, y) {
    if (_focusGridName) {
        commitGridEdit(_focusGridName)
        var [oldX, oldY] = getPosByGridName(_focusGridName)
        removeGridFocus(oldX, oldY)
    }
    _focusGridName = name
    addGridFocus(x, y)
    setGridToDisplay(name)
}

function initGlobals(columnCount, rowCount, dataMap) {
    _sheetDataMap = dataMap
    _columnCount = columnCount
    _rowCount = rowCount
    _gridName = $("#gridName")
    _gridValue = $("#gridValue")
    _sideBar = $("#sideBar")
    _rowNumberBar = $("#rowNumberBar")
    _sheetHeader = $("#sheetHeader")
    _sheetContent = $("#sheetContent")
    _formatBold = $("#formatBold")
    _formatItalics = $("#formatItalics")
    _formatUnderline = $("#formatUnderline")
    _errorInfo = $("#errorInfo")
}

function setSheetBodyHeight() {
    var rect = _sheetContent[0].getBoundingClientRect()
    var pageHeight = $(window).height()
    var height = pageHeight - 30 - rect.top
    _sheetContent.css("maxHeight", "" + height + "px")
    _sideBar.css("height", "" + (height + 21) + "px")
}

function initRowNumberBar() {
    for (var i = 1; i <= _rowCount; ++i) {
        var grid = document.createElement("div")
        grid.id = "r" + i
        grid.innerHTML = i
        _rowNumberBar.append(grid)
    }
}

function initSheetHeader() {
    for (var i = 1; i <= _columnCount; ++i) {
        var grid = document.createElement("div")
        grid.id = "s" + i
        grid.innerHTML = getColumnName(i)
        _sheetHeader.append(grid)
    }
}

function initSheetContent() {
    _sheetContent.css("grid-template-columns", "repeat(" + _columnCount + ", 1fr)")
    for (var y = 1; y <= _rowCount; ++y) {
        for (var x = 1; x <= _columnCount; ++x) {
            var grid = document.createElement("div")
            grid.id = getGridNameByPos(x, y)
            var data = getGridData(grid.id)
            grid.innerHTML = getGridFormattedHtml(data)
            _sheetContent.append(grid)
        }
    }
}

function initFocusGrid() {
    var name = "A1"
    var [x, y] = getPosByGridName(name)
    switchFocusGrid(name, x, y)
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
            var offsetTop = _rowNumberBar.offset()
            offsetTop.top -= diffTop
            _rowNumberBar.offset(offsetTop)
        }
    })
}

function addSheetContentClickListener() {
    $("#sheetContent >div").click(function() {
        var grid = $(this)[0]
        var name = grid.id
        var [x, y] = getPosByGridName(name)
        commitGridEdit()
        switchFocusGrid(name, x, y)
    })
}

function addInputListener() {
    _gridValue.keydown(function(event) {
        var [x, y] = getPosByGridName(_focusGridName)
        switch (event.keyCode) {
        case 37:    // left
            x = x === 1 ? x : x - 1
            break
        case 38:    // up
            y = y === 1 ? y : y - 1
            break
        case 39:    // right
            x = x === 100 ? x : x + 1
            break
        case 40:    // down
        case 13:    // enter
            y = y === 100 ? y : y + 1
            break
        default:
            return
        }
        var name = getGridNameByPos(x, y)
        switchFocusGrid(name, x, y)
    })
    var commit = function(event) {
        commitGridEdit()
        setGridToDisplay(_focusGridName)
    }
    _gridValue.blur(commit)
    _formatBold.change(commit)
    _formatItalics.change(commit)
    _formatUnderline.change(commit)
}

function addRefreshButtonClickListener() {
    $("#refreshButton").click(function(event) {
        _rowNumberBar.empty()
        _sheetHeader.empty()
        _sheetContent.empty()
        alert("Sheet already destroyed, click OK to rebuild.")
        initSheetBody(_columnCount, _rowCount, _sheetDataMap)
    })
}

function initSheetBody(columnCount, rowCount, dataMap) {
    initGlobals(columnCount, rowCount, dataMap)
    setSheetBodyHeight()
    initRowNumberBar()
    initSheetHeader()
    initSheetContent()
    initFocusGrid()
    addSheetBodyResizeListener()
    addSheetContentScrollListener()
    addSheetContentClickListener()
    addInputListener()
    addRefreshButtonClickListener()
}

function createSpreadSheet(columnCount, rowCount) {
    initSheetBody(columnCount, rowCount, new Map())
}