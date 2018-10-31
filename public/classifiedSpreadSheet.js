
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
        return "&nbsp;"
    }

    get value() {
        return 0
    }

    onOtherGridValueChange(spreadSheet, selfName) {

    }

    discard(spreadSheet, selfName) {

    }
}

class NumberGridValue extends NormalGridValue {
    get value() {
        return parseFloat(this._strRaw)
    }
}

let __spreadSheet = null

function __getGridValue(name) {
    return __spreadSheet.getGridData(name).value
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

const regexLegal = /^[ A-Z0-9\(\)\+\-\*\/%\^:,\.]+$/
const regexColonPair = /([A-Z]+[0-9]+) *: *([A-Z]+[0-9]+)/g
const regexFunctionName = /[A-Z][A-Z0-9_]+ *\(/g
const regexGridName = /[A-Z]+[0-9]+/g

class FormularGridValue extends NormalGridValue {
    constructor(spreadSheet, selfName, strRaw) {
        super(strRaw.toUpperCase())
        this._computeFormulaGridValue(spreadSheet, selfName)
    }

    get strDisplay() {
        var color = this._formulaError ? "red" : "dodgerblue"
        return "<span style='color:" + color + ";'>" + this._value + "</span>"
    }

    get strError() {
        return this._formulaError ? this._formulaError : "&nbsp;"
    }

    get value() {
        return isNaN(this._value) ? 0 : this._value
    }

    onOtherGridValueChange(spreadSheet, selfName) {
        this._computeFormulaGridValue(spreadSheet, selfName)
    }

    discard(spreadSheet, selfName) {
        let grid = spreadSheet.getGridData(selfName)
        grid.unlistenReferedGridValueChange()
    }

    _checkFormula(formula) {
        if (!formula || !regexLegal.test(formula)) {
            throw "illegal formular"
        }
        return formula
    }

    _replaceFormulaColonPairs(spreadSheet, formula) {
        regexColonPair.lastIndex = 0
        return formula.replace(regexColonPair, function(str, grid1, grid2) {
            let [x1, y1] = spreadSheet.getPosByGridName(grid1)
            let [x2, y2] = spreadSheet.getPosByGridName(grid2)
            if (!x1 || !x2) {
                throw "invalid grid name: " + !x1 ? grid1 : grid2
            }
            if (x1 > x2) {
                [x1, x2] = [x2, x1]
            }
            if (y1 > y2) {
                [y1, y2] = [y2, y1]
            }
            let arr = []
            while (true) {
                let y = y1
                while (true) {
                    arr.push(spreadSheet.getGridNameByPos(x1, y))
                    if (y === y2) break
                    y++
                }
                if (x1 === x2) break
                x1++
            }
            return arr.join()
        })
    }

    _replaceFormulaFunctionNames(formula) {
        regexFunctionName.lastIndex = 0
        return formula.replace(regexFunctionName, "__f_$&")
    }

    _replaceFormulaGridNames(spreadSheet, formula, refGridNameMap) {
        regexGridName.lastIndex = 0
        return formula.replace(regexGridName, function(name) {
            if (!spreadSheet.getPosByGridName(name)) {
                throw "invalid grid name: " + name
            }
            refGridNameMap.set(name, true)
            return "__getGridValue(\"" + name + "\")"
        })
    }

    _parseFormulaGrid(spreadSheet, selfName) {
        let formula = this._strRaw.substring(1)
        let refGridNameMap = new Map()
        let grid = spreadSheet.getGridData(selfName)
        try {
            formula = this._checkFormula(formula)
            formula = this._replaceFormulaColonPairs(spreadSheet, formula)
            formula = this._replaceFormulaFunctionNames(formula)
            formula = this._replaceFormulaGridNames(spreadSheet, formula, refGridNameMap)
            grid.checkReferedGrid(selfName, selfName, refGridNameMap)
        } catch (err) {
            this._formulaError = err
            this._formula = "_"
            this._value = "error"
            return
        }
        this._formula = formula
        grid.listenReferedGridValueChange(refGridNameMap)
    }
    
    _computeFormulaGridValue(spreadSheet, name) {
        if (!this._formula) {
            this._parseFormulaGrid(spreadSheet, name)
        }
        if (!this._formulaError) {
            try {
                this._value = eval(this._formula)
            } catch (err) {
                this._formulaError = err //"run time error"
                this._formula = "_"
                this._value = "error"
            }
        }
    }
}

class GridValueFactory {
    static createGridValue(spreadSheet, selfName, strRaw) {
        if (strRaw[0] === "=") {
            return new FormularGridValue(spreadSheet, selfName, strRaw)
        } else if (strRaw === "" || isNaN(strRaw)) {
            return new NormalGridValue(strRaw)
        } else {
            return new NumberGridValue(strRaw)
        }
    }
}

class Grid {
    constructor(spreadSheet, selfName) {
        this._spreadSheet = spreadSheet
        this._selfName = selfName
    }

    updateData(strRaw, isBold, isItalics, isUnderline) {
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

    listenReferedGridValueChange(refGridNameMap) {
        if (this._refGridNameMap) {
            this.unlistenReferedGridValueChange()
        }
        this._refGridNameMap = refGridNameMap
        for (let name of this._refGridNameMap.keys()) {
            let data = this._spreadSheet.getGridDataEnsure(name)
            data.subscribeValueChange(this._selfName)
        }
    }

    unlistenReferedGridValueChange() {
        if (this._refGridNameMap) {
            for (let name of this._refGridNameMap.keys()) {
                let data = this._spreadSheet.getGridDataEnsure(name)
                data.unsubscribeValueChange(this._selfName)
            }
            this._refGridNameMap = null
        }
    }

    checkReferedGrid(selfName, checkingName, refGridNameMap) {
        if (!refGridNameMap) {
            refGridNameMap = this._refGridNameMap
            if (!refGridNameMap) {
                return
            }
        }
        for (let name of refGridNameMap.keys()) {
            if (name === checkingName) {
                throw "circular reference - " + selfName + " already refered this grid " + checkingName
            }
            let data = this._spreadSheet.getGridDataEnsure(name)
            data.checkReferedGrid(name, checkingName)
        }
    }

    onOtherGridValueChange(otherGridName) {
        let orginalValue = this._gridValue.value
        this._gridValue.onOtherGridValueChange(this._spreadSheet, this._selfName)
        let newValue = this._gridValue.value
        if (orginalValue !== newValue) {
            this._notifyGridValueChange()
        }
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
            if (this._gridValue.strRaw === strRaw) {
                return
            }
            this._gridValue.discard(this._spreadSheet, this._selfName)
        }
        this._gridValue = GridValueFactory.createGridValue(this._spreadSheet, this._selfName, strRaw)
        this._notifyGridValueChange()
    }

    _notifyGridValueChange() {
        if (this._selfName) {
            this._spreadSheet.onGridValueUpdate(this._selfName)
        }
        if (this._listenerMap) {
            for (let name of this._listenerMap.keys()) {
                let notifyData = this._spreadSheet.getGridData(name)
                notifyData.onOtherGridValueChange(this._selfName)
            }
        }
    }
}

class SpreadSheet {
    constructor(columnCount, rowCount, dataMap) {
        this._columnCount = columnCount
        this._rowCount = rowCount
        this._dataMap = dataMap
        this._onGridValueUpdate = function() {}
        this._emptyGridData = new Grid(this, "")
        this._emptyGridData.updateData("", false, false, false)
    }

    getColumnName(x) {
        let arr = []
        while(true) {
            x--
            let n = x % 26
            arr.unshift(String.fromCharCode(65 + n))    // 65 is ASCII of 'A'
            x = x - n
            if (x === 0) {
                return arr.join("")
            }
            x /= 26
        }
    }
    
    getGridNameByPos(x, y) {
        return this.getColumnName(x) + y
    }
    
    getPosByGridName(name) {
        let x = 0, y = 0
        for (let i = 0; i < name.length; ++i) {
            let n = name.charCodeAt(i)
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
        let data = this._dataMap.get(name)
        return data ? data : this._emptyGridData
    }

    getGridDataEnsure(name) {
        let data = this._dataMap.get(name)
        if (!data) {
            data = new Grid(this, name)
            data.updateData("", false, false, false)
            this._dataMap.set(name, data)
        }
        return data
    }

    updateGridData(name, strRaw, isBold, isItalics, isUnderline) {
        let data = this.getGridDataEnsure(name)
        data.updateData(strRaw, isBold, isItalics, isUnderline)
        this._onGridValueUpdate(name)
    }
}

class SpreadSheetUI {
    constructor(spreadSheet, columnCount, rowCount) {
        this._spreadSheet = spreadSheet
        this._spreadSheet.setGridValueUpdateHandler(this._onGridValueUpdate.bind(this))
        this._columnCount = columnCount
        this._rowCount = rowCount
        this._selfName = $("#gridName")
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
        let rect = this._sheetContent[0].getBoundingClientRect()
        let pageHeight = $(window).height()
        let height = pageHeight - 30 - rect.top
        this._sheetContent.css("maxHeight", "" + height + "px")
        this._sideBar.css("height", "" + (height + 21) + "px")
    }
    
    _initRowNumberBar() {
        for (let i = 1; i <= this._rowCount; ++i) {
            let grid = document.createElement("div")
            grid.id = "r" + i
            grid.innerHTML = i
            this._rowNumberBar.append(grid)
        }
    }
    
    _initSheetHeader() {
        for (let i = 1; i <= this._columnCount; ++i) {
            let grid = document.createElement("div")
            grid.id = "s" + i
            grid.innerHTML = this._spreadSheet.getColumnName(i)
            this._sheetHeader.append(grid)
        }
    }
    
    _initSheetContent() {
        this._sheetContent.css("grid-template-columns", "repeat(" + this._columnCount + ", 1fr)")
        for (let y = 1; y <= this._rowCount; ++y) {
            for (let x = 1; x <= this._columnCount; ++x) {
                let grid = document.createElement("div")
                grid.id = this._spreadSheet.getGridNameByPos(x, y)
                let data = this._spreadSheet.getGridData(grid.id)
                grid.innerHTML = data.strFormated
                this._sheetContent.append(grid)
            }
        }
    }
    
    _initFocusGrid() {
        let name = "A1"
        let [x, y] = this._spreadSheet.getPosByGridName(name)
        this._switchFocusGrid(name, x, y)
    }
    
    _addSheetBodyResizeListener() {
        $(window).resize(function() {
            this._setSheetBodyHeight()
        })
    }
    
    _addSheetContentScrollListener() {
        let scrollTop = this._sheetContent.scrollTop()
        let scrollLeft = this._sheetContent.scrollLeft()
        let that = this
        this._sheetContent.scroll(function onSheetScroll(event) {
            let curScrollTop = that._sheetContent.scrollTop()
            let curScrollLeft = that._sheetContent.scrollLeft()
            let diffTop = curScrollTop - scrollTop
            let diffLeft = curScrollLeft - scrollLeft
            scrollTop = curScrollTop
            scrollLeft = curScrollLeft
            if (diffLeft !== 0) {
                let offsetLeft = that._sheetHeader.offset()
                offsetLeft.left -= diffLeft
                that._sheetHeader.offset(offsetLeft)
            }
            if (diffTop !== 0) {
                let offsetTop = that._rowNumberBar.offset()
                offsetTop.top -= diffTop
                that._rowNumberBar.offset(offsetTop)
            }
        })
    }
    
    _addSheetContentClickListener() {
        let that = this
        $("#sheetContent >div").click(function() {
            let grid = $(this)[0]
            let name = grid.id
            let [x, y] = that._spreadSheet.getPosByGridName(name)
            that._switchFocusGrid(name, x, y)
        })
    }
    
    _addInputListener() {
        let that = this
        this._gridValue.keydown(function(event) {
            let [x, y] = that._spreadSheet.getPosByGridName(that._focusGridName)
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
            let name = that._spreadSheet.getGridNameByPos(x, y)
            that._switchFocusGrid(name, x, y)
        })
        let onCommit = function(event) {
            that._commitGridEdit()
            that._setGridToDisplay(that._focusGridName)
        }
        this._gridValue.blur(onCommit)
        this._formatBold.change(onCommit)
        this._formatItalics.change(onCommit)
        this._formatUnderline.change(onCommit)
    }
    
    _addRefreshButtonClickListener() {
        let that = this
        $("#refreshButton").click(function(event) {
            that._rowNumberBar.empty()
            that._sheetHeader.empty()
            that._sheetContent.empty()
            alert("Sheet already destroyed, click OK to rebuild.")
            that.init()
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
        return $(this._sheetContent[0].children[(y - 1) * this._columnCount + x - 1])
    }
        
    _commitGridEdit() {
        let val = this._gridValue.val()
        let bold = this._formatBold.prop("checked")
        let italics = this._formatItalics.prop("checked")
        let underline = this._formatUnderline.prop("checked")
        this._spreadSheet.updateGridData(this._focusGridName, val, bold, italics, underline)
    }
    
    _addGridFocus(x, y) {
        let focusGrid = this._getDomGridByPos(x, y)
        focusGrid.addClass("gridFocus")
        let focusRowNumber = this._getRowNumberBarGridByY(y)
        focusRowNumber.addClass("gridFocus")
        let focusColumn = this._getSheetHeaderGridByX(x)
        focusColumn.addClass("gridFocus")
    }
    
    _removeGridFocus(x, y) {
        let focusColumn = this._getSheetHeaderGridByX(x)
        focusColumn.removeClass("gridFocus")
        let focusRowNumber = this._getRowNumberBarGridByY(y)
        focusRowNumber.removeClass("gridFocus")
        let focusGrid = this._getDomGridByPos(x, y)
        focusGrid.removeClass("gridFocus")
    }
    
    _setGridToDisplay(name) {
        let data = this._spreadSheet.getGridData(name)
        this._selfName.val(name)
        this._gridValue.val(data.strRaw)
        this._gridValue.focus()
        this._formatBold.prop("checked", data.isBold)
        this._formatItalics.prop("checked", data.isItalics)
        this._formatUnderline.prop("checked", data.isUnderline)
        this._errorInfo.html(data.strError)
    }

    _onGridValueUpdate(name) {
        let grid = this._getDomGridByName(name)
        let data = this._spreadSheet.getGridData(name)
        grid.html(data.strFormated)
    }
    
    _switchFocusGrid(name, x, y) {
        if (this._focusGridName) {
            this._commitGridEdit()
            let [oldX, oldY] = this._spreadSheet.getPosByGridName(this._focusGridName)
            this._removeGridFocus(oldX, oldY)
        }
        this._focusGridName = name
        this._addGridFocus(x, y)
        this._setGridToDisplay(name)
    }
}

function createSpreadSheet(columnCount, rowCount) {
    let spreadSheet = new SpreadSheet(columnCount, rowCount, new Map())
    __spreadSheet = spreadSheet
    let spreadSheetUI = new SpreadSheetUI(spreadSheet, columnCount, rowCount)
    spreadSheetUI.init()
}
