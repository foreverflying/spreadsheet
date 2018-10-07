
var sheetData = {}

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

function setSheetBodyHeight(sideBar, sheetContent) {
    var rect = sheetContent[0].getBoundingClientRect()
    var pageHeight = $(window).height()
    var height = pageHeight - 30 - rect.top
    sheetContent.css("maxHeight", "" + height + "px")
    sideBar.css("height", "" + (height + 21) + "px")
}

function setSheetBodySizeListener(sideBar, sheetContent) {
    $(window).resize(function() {
        setSheetBodyHeight(sideBar, sheetContent)
    })
}

function setSheetContentScrollListener(rowNumber, sheetHeader, sheetContent) {
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

function initSheetBody() {
    var sideBar = $("#sideBar")
    var rowNumber = $("#rowNumber")
    var sheetHeader = $("#sheetHeader")
    var sheetContent = $("#sheetContent")
    setSheetBodyHeight(sideBar, sheetContent)
    setSheetBodySizeListener(sideBar, sheetContent)
    setSheetContentScrollListener(rowNumber, sheetHeader, sheetContent)
}


function initRowNumber(rowCount) {
    var sheetHeader = $("#rowNumber")
    for (var i = 1; i <= rowCount; ++i) {
        var grid = document.createElement("div")
        grid.innerHTML = i
        sheetHeader.append(grid)
    }
}

function initSheetHeader(columnCount) {
    var sheetHeader = $("#sheetHeader")
    for (var i = 1; i <= columnCount; ++i) {
        var grid = document.createElement("div")
        grid.innerHTML = getColumnName(i)
        sheetHeader.append(grid)
    }
}

function initSheetContent(columnCount, rowCount) {
    var sheetContent = $("#sheetContent")
    for (var i = 1; i <= rowCount; ++i) {
        var row = document.createElement("div")
        row.className = "gridRow"
        for (var j = 1; j <= columnCount; ++j) {
            var grid = document.createElement("div")
            row.appendChild(grid)
        }
        row.appendChild(grid)
        sheetContent.append(row)
    }
}

function createSpreadSheet(columnCount, rowCount) {
    initSheetBody()
    initRowNumber(rowCount)
    initSheetHeader(columnCount)
    initSheetContent(columnCount, rowCount)
}