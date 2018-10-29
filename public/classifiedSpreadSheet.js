
class NormalGridValue {
    constructor(strRaw) {
        this._strRaw = strRaw
    }

    get strRaw() {
        return this._strRaw
    }

    get strDisplay() {
        return this._strRaw
    }

    get strError() {
        return ""
    }

    get value() {
        return 0
    }

    onOtherGridValueChange(spreadSheet, gridName) {

    }

    discard(spreadSheet, gridName) {

    }
}

class NumberGridValue extends NormalGridValue {
    get value() {
        return parseFloat(this._strRaw)
    }
}

class FormularGridValue extends NormalGridValue {
    constructor(strRaw) {
        super(strRaw.toUpperCase())
    }

    get strDisplay() {
        return this._strRaw
    }

    get strError() {
        return ""
    }

    get value() {
        return 0
    }

    onOtherGridValueChange(spreadSheet, gridName) {

    }

    discard(spreadSheet, gridName) {
        if (this._refGridNameMap) {
            for (var name of this._refGridNameMap.keys()) {
                var data = spreadSheet.getGridData(name)
                data.unsubscribeValueChange(gridName)
            }
        }
    }
}

class GridValueFactory {
    static createGridValue(strRaw) {
        if (strRaw[0] === "=") {
            return new FormularGridValue(strRaw)
        } else if (isNaN(strRaw)) {
            return new NormalGridValue(strRaw)
        } else {
            return new NumberGridValue(strRaw)
        }
    }
}

class Grid {
    constructor(spreadSheet, gridName) {
        this._spreadSheet = spreadSheet
        this._gridName = gridName
    }

    update(strRaw, isBold, isItalics, isUnderline) {
        this._updateStrRaw(strRaw)
        this._isBold = isBold
        this._isItalics = isItalics
        this._isUnderline = isUnderline
    }

    subscribeValueChange(notifyGridName) {
        if (!this._listenerMap) {
            this._listenerMap = new Map()
        }
        this._listenerMap.set(notifyGridName, true)
    }

    unsubscribeValueChange(notifyGridName) {
        this._listenerMap.delete(notifyGridName)
    }

    onOtherGridValueChange(otherGridName) {
        this._gridValue.onOtherGridValueChange(this._spreadSheet, this._gridName)
        this._notifyGridValueChange()
    }

    get isBold() {
        return this._isBold
    }

    get isItalics() {
        return this._isItalics
    }

    get isUnderline() {
        return this._isUnderline
    }

    get strRaw() {
        return this._gridValue.strRaw
    }

    get strDisplay() {
        return this._gridValue.strDisplay
    }

    get strError() {
        return this._gridValue.strError
    }

    get value() {
        return this._gridValue.value
    }

    get strFormated() {
        let ret = this.strDisplay
        if (this._isBold) {
            ret = "<b>" + ret + "</b>"
        }
        if (this._isItalics) {
            ret = "<i>" + ret + "</i>"
        }
        if (this._isUnderline) {
            ret = "<u>" + ret + "</u>"
        }
        return ret
    }

    _updateStrRaw(strRaw) {
        if (this._gridValue) {
            this._gridValue.discard(this._spreadSheet, this._gridName)
        }
        this._gridValue = GridValueFactory.createGridValue(strRaw)
        if (this._listenerMap) {
            this._notifyGridValueChange()
        }
    }

    _notifyGridValueChange() {
        for (var name of this._listenerMap.keys()) {
            var notifyData = this._spreadSheet.getGridData(name)
            notifyData.onOtherGridValueChange(this._gridName)
            this._spreadSheet.onGridValueUpdate(name)
        }
    }
}

class SpreadSheet {
    constructor(columnCount, rowCount, dataMap) {
        this._columnCount = columnCount
        this._rowCount = rowCount
        this._dataMap = dataMap
        this._onGridValueUpdate = function() {}
        this._emptyGridData = new Grid()
        this._emptyGridData.update("", false, false, false)
    }

    get columnCount() {
        return this._columnCount
    }
    
    get rowCount() {
        return this._rowCount
    }

    getColumnName(x) {
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
    
    getGridNameByPos(x, y) {
        return getColumnName(x) + y
    }
    
    getPosByGridName(name) {
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
        return x <= this._columnCount && y <= this._rowCount ? [x, y] : null
    }
    
    setGridValueUpdateHandler(handler) {
        this._onGridValueUpdate = handler
    }

    onGridValueUpdate(gridName) {
        this._onGridValueUpdate(gridName)
    }

    getGridData(name) {
        var data = this._dataMap.get(name)
        return data ? data : this._emptyGridData
    }
}

class SpreadSheetUI {
    constructor(spreadSheet) {
        this._spreadSheet = spreadSheet
        this._spreadSheet.setGridValueUpdateHandler(this._onGridValueUpdate)
        this._gridName = $("#gridName")
        this._gridValue = $("#gridValue")
        this._sideBar = $("#sideBar")
        this._rowNumberBar = $("#rowNumberBar")
        this._sheetHeader = $("#sheetHeader")
        this._sheetContent = $("#sheetContent")
        this._formatBold = $("#formatBold")
        this._formatItalics = $("#formatItalics")
        this._formatUnderline = $("#formatUnderline")
        this._errorInfo = $("#errorInfo")
    }

    init() {
        this._setSheetBodyHeight()
        this._initRowNumberBar()
        this._initSheetHeader()
        this._initSheetContent()
        this._initFocusGrid()
        this._addSheetBodyResizeListener()
        this._addSheetContentScrollListener()
        this._addSheetContentClickListener()
        this._addInputListener()
        this._addRefreshButtonClickListener()
    }

    _setSheetBodyHeight() {
        var rect = this._sheetContent[0].getBoundingClientRect()
        var pageHeight = $(window).height()
        var height = pageHeight - 30 - rect.top
        this._sheetContent.css("maxHeight", "" + height + "px")
        this._sideBar.css("height", "" + (height + 21) + "px")
    }
    
    _initRowNumberBar() {
        for (var i = 1; i <= this._rowCount; ++i) {
            var grid = document.createElement("div")
            grid.id = "r" + i
            grid.innerHTML = i
            this._rowNumberBar.append(grid)
        }
    }
    
    _initSheetHeader() {
        for (var i = 1; i <= this._columnCount; ++i) {
            var grid = document.createElement("div")
            grid.id = "s" + i
            grid.innerHTML = this._spreadSheet.getColumnName(i)
            this._sheetHeader.append(grid)
        }
    }
    
    _initSheetContent() {
        this._sheetContent.css("grid-template-columns", "repeat(" + this._columnCount + ", 1fr)")
        for (var y = 1; y <= this._rowCount; ++y) {
            for (var x = 1; x <= this._columnCount; ++x) {
                var grid = document.createElement("div")
                grid.id = this._spreadSheet.getGridNameByPos(x, y)
                var data = this._spreadSheet.getGridData(grid.id)
                grid.innerHTML = data.strFormated
                this._sheetContent.append(grid)
            }
        }
    }
    
    _initFocusGrid() {
        var name = "A1"
        var [x, y] = this._spreadSheet.getPosByGridName(name)
        this._switchFocusGrid(name, x, y)
    }
    
    _addSheetBodyResizeListener() {
        $(window).resize(function() {
            this._setSheetBodyHeight()
        })
    }
    
    _addSheetContentScrollListener() {
        var scrollTop = this._sheetContent.scrollTop()
        var scrollLeft = this._sheetContent.scrollLeft()
    
        this._sheetContent.scroll(function onSheetScroll(event) {
            var curScrollTop = this._sheetContent.scrollTop()
            var curScrollLeft = this._sheetContent.scrollLeft()
            var diffTop = curScrollTop - scrollTop
            var diffLeft = curScrollLeft - scrollLeft
            scrollTop = curScrollTop
            scrollLeft = curScrollLeft
            if (diffLeft !== 0) {
                var offsetLeft = this._sheetHeader.offset()
                offsetLeft.left -= diffLeft
                this._sheetHeader.offset(offsetLeft)
            }
            if (diffTop !== 0) {
                var offsetTop = this._rowNumberBar.offset()
                offsetTop.top -= diffTop
                this._rowNumberBar.offset(offsetTop)
            }
        })
    }
    
    _addSheetContentClickListener() {
        $("#sheetContent >div").click(function() {
            var grid = $(this)[0]
            var name = grid.id
            var [x, y] = this._spreadSheet.getPosByGridName(name)
            this._commitGridEdit()
            this._switchFocusGrid(name, x, y)
        })
    }
    
    _addInputListener() {
        this._gridValue.keydown(function(event) {
            var [x, y] = this._spreadSheet.getPosByGridName(this._focusGridName)
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
            var name = this._spreadSheet.getGridNameByPos(x, y)
            this._switchFocusGrid(name, x, y)
        })
        var onCommit = function(event) {
            this._commitGridEdit()
            this._setGridToDisplay(_focusGridName)
        }
        this._gridValue.blur(onCommit)
        this._formatBold.change(onCommit)
        this._formatItalics.change(onCommit)
        this._formatUnderline.change(onCommit)
    }
    
    _addRefreshButtonClickListener() {
        $("#refreshButton").click(function(event) {
            this._rowNumberBar.empty()
            this._sheetHeader.empty()
            this._sheetContent.empty()
            alert("Sheet already destroyed, click OK to rebuild.")
            this._init()
        })
    }

    _getRowNumberBarGridByY(y) {
        return $("#r" + y)
    }
    
    _getSheetHeaderGridByX(x) {
        return $("#s" + x)
    }
    
    _getDomGridByName(name) {
        return $("#" + name)
    }
    
    _getDomGridByPos(x, y) {
        return $(this._sheetContent[0].children[(y - 1) * this._spreadSheet.columnCount + x - 1])
    }
        
    _commitGridEdit() {
        var val = this._gridValue.val()
        var bold = this._formatBold.prop("checked")
        var italics = this._formatItalics.prop("checked")
        var underline = this._formatUnderline.prop("checked")
        var data = this._spreadSheet.getGridDataEnsure(this._focusGridName)
        data.update(val, bold, italics, underline)
    }
    
    _addGridFocus(x, y) {
        var focusGrid = this._getDomGridByPos(x, y)
        focusGrid.addClass("gridFocus")
        var focusRowNumber = this._getRowNumberBarGridByY(y)
        focusRowNumber.addClass("gridFocus")
        var focusColumn = this._getSheetHeaderGridByX(x)
        focusColumn.addClass("gridFocus")
    }
    
    _removeGridFocus(x, y) {
        var focusColumn = _getSheetHeaderGridByX(x)
        focusColumn.removeClass("gridFocus")
        var focusRowNumber = _getRowNumberBarGridByY(y)
        focusRowNumber.removeClass("gridFocus")
        var focusGrid = _getDomGridByPos(x, y)
        focusGrid.removeClass("gridFocus")
    }
    
    _setGridToDisplay(name) {
        var data = this._spreadSheet.getGridData(name)
        this._gridName.val(name)
        this._gridValue.val(data.strFormated)
        this._gridValue.focus()
        this._formatBold.prop("checked", data.isBold)
        this._formatItalics.prop("checked", data.isItalics)
        this._formatUnderline.prop("checked", data.isUnderline)
        this._errorInfo.html(data.strError)
    }

    _onGridValueUpdate(name) {
        var grid = this._getDomGridByName(name)
        var data = this._spreadSheet.getGridData(name)
        grid.html(data.strFormated)
    }
    
    _switchFocusGrid(name, x, y) {
        if (this._focusGridName) {
            this._commitGridEdit(this._focusGridName)
            var [oldX, oldY] = this._spreadSheet.getPosByGridName(this._focusGridName)
            this._removeGridFocus(oldX, oldY)
        }
        this._focusGridName = name
        this._addGridFocus(x, y)
        this._setGridToDisplay(name)
    }
}

class Test {
    constructor() {
        this._a = 1;
    }

    get a() {
        return this._a
    }
}

function createSpreadSheet(columnCount, rowCount) {
    let spreadSheet = new SpreadSheet(columnCount, rowCount, new Map())
    let spreadSheetUI = new SpreadSheetUI(spreadSheet)
    spreadSheetUI.init()
}
