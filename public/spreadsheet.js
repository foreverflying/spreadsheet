
var sheetData = {}
var focusX = 1
var focusY = 1

const ROW_COUNT = 100
const COLUMN_COUNT = 100

var gridName, gridValue, sideBar, rowNumber, sheetHeader, sheetContent

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

function getGridName(x, y) {
    return getColumnName(x) + y
}

function getGridValue(x, y) {
    var data = sheetData["v" + x + "_" + y]
    return data ? data.value : ""
}

function setGridValue(x, y, val) {
    var key = "v" + x + "_" + y
    var data = sheetData[key]
    if (data) {
        data.value = val
    } else {
        sheetData[key] = { value: val }
    }
}

function getGridByXY(x, y) {
    return sheetContent.children().eq(y - 1).children().eq(x - 1)
}

function getRowNumberGridByY(y) {
    return $("#r" + y)
}

function getColumnGridByX(x) {
    return $("#c" + x)
}

function translateGridNameToKey(name) {
    var patt = /[A-Z]+[0-9]+/
    if (!patt.test(name)) {
        return ""
    }
    var x = 0, y = 0
    for (var char in name) {
        var n = char.charCodeAt(0)
        if (n > 65) {
            x = x * 26 + n
        } else {
            y = y * 10 + n
        }
    }
    return "v" + (x + 1) + "_" + (y + 1)
}

function parseFomulas(val) {
    if (val[0] !== "=") {
        return null;
    }
    // TODO:
}

function commitEdit() {
    var val = gridValue.val()
    setGridValue(focusX, focusY, val)
    var grid = getGridByXY(focusX, focusY)
    grid.text(val)
}

function switchFocusGrid(x, y) {
    commitEdit()
    var focusColumn = getColumnGridByX(focusX)
    focusColumn.removeClass("gridFocus")
    var focusRowNumber = getRowNumberGridByY(focusY)
    focusRowNumber.removeClass("gridFocus")
    var focusGrid = getGridByXY(focusX, focusY)
    focusGrid.removeClass("gridFocus")
    focusX = x
    focusY = y
    var focusGrid = getGridByXY(focusX, focusY)
    focusGrid.addClass("gridFocus")
    var focusRowNumber = getRowNumberGridByY(focusY)
    focusRowNumber.addClass("gridFocus")
    var focusColumn = getColumnGridByX(focusX)
    focusColumn.addClass("gridFocus")
    var name = getGridName(x, y)
    gridName.val(name)
    gridValue.val(getGridValue(x, y))
    gridValue.focus()
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
    var rect = sheetContent[0].getBoundingClientRect()
    var pageHeight = $(window).height()
    var height = pageHeight - 30 - rect.top
    sheetContent.css("maxHeight", "" + height + "px")
    sideBar.css("height", "" + (height + 21) + "px")
}

function initRowNumber(rowCount) {
    for (var i = 1; i <= rowCount; ++i) {
        var grid = document.createElement("div")
        grid.id = "r" + i
        grid.innerHTML = i
        rowNumber.append(grid)
    }
}

function initSheetHeader(columnCount) {
    for (var i = 1; i <= columnCount; ++i) {
        var grid = document.createElement("div")
        grid.id = "c" + i
        grid.innerHTML = getColumnName(i)
        sheetHeader.append(grid)
    }
}

function initSheetContent(columnCount, rowCount) {
    for (var y = 1; y <= rowCount; ++y) {
        var row = document.createElement("div")
        row.className = "gridRow"
        for (var x = 1; x <= columnCount; ++x) {
            var grid = document.createElement("div")
            grid.xIndex = x
            grid.innerHTML = getGridValue(x, y)
            row.appendChild(grid)
        }
        row.yIndex = y
        row.appendChild(grid)
        sheetContent.append(row)
    }
}

function initFocusGrid() {
    switchFocusGrid(focusX, focusY)
}

function addSheetBodyResizeListener(sheetContent) {
    $(window).resize(function() {
        setSheetBodyHeight()
    })
}

function addSheetContentScrollListener() {
    var scrollTop = sheetContent.scrollTop()
    var scrollLeft = sheetContent.scrollLeft()

    sheetContent.scroll(function onSheetScroll(event) {
        var curScrollTop = sheetContent.scrollTop()
        var curScrollLeft = sheetContent.scrollLeft()
        var diffTop = curScrollTop - scrollTop
        var diffLeft = curScrollLeft - scrollLeft
        scrollTop = curScrollTop
        scrollLeft = curScrollLeft
        if (diffLeft !== 0) {
            var offsetLeft = sheetHeader.offset()
            offsetLeft.left -= diffLeft
            sheetHeader.offset(offsetLeft)
        }
        if (diffTop !== 0) {
            var offsetTop = rowNumber.offset()
            offsetTop.top -= diffTop
            rowNumber.offset(offsetTop)
        }
    })
}

function addSheetContentClickListener() {
    $(".gridRow >div").click(function() {
        var grid = $(this)
        var x = grid[0].xIndex
        var y = grid.parent()[0].yIndex
        switchFocusGrid(x, y)
    })
}

function addInputListener() {
    gridValue.keydown(function(event) {
        if (event.keyCode === 37 ) {
            switchNeighborGridLeft(focusX, focusY)
        } else if (event.keyCode === 38 ) {
            switchNeighborGridUp(focusX, focusY)
        } else if (event.keyCode === 39 ) {
            switchNeighborGridRight(focusX, focusY)
        } else if (event.keyCode === 40 || event.keyCode === 13 ) {
            switchNeighborGridDown(focusX, focusY)
        }
    })
}

function addRefreshButtonClickListener() {
    $("#refreshButton").click(function(event) {
        sheetContent.empty()
        alert("Sheet already destroyed, click OK to rebuild.")
        initSheetBody(100, 100)
    })
}

function initSheetBody(columnCount, rowCount) {
    gridName = $("#gridName")
    gridValue = $("#gridValue")
    sideBar = $("#sideBar")
    rowNumber = $("#rowNumber")
    sheetHeader = $("#sheetHeader")
    sheetContent = $("#sheetContent")

    setSheetBodyHeight()
    initRowNumber(rowCount)
    initSheetHeader(columnCount)
    initSheetContent(columnCount, rowCount)
    initFocusGrid()
    addSheetBodyResizeListener(sheetContent)
    addSheetContentScrollListener(sheetHeader, sheetContent)
    addSheetContentClickListener()
    addInputListener()
    addRefreshButtonClickListener()
}

function createSpreadSheet(columnCount, rowCount) {
    initSheetBody(columnCount, rowCount)
}