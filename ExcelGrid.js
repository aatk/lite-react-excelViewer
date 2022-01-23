import React from 'react';
import "./excelGrid.css";

function _26to10(value26) {
    //3А7. 7*16^0+10*16^1+3*16^2=7+160+768=935.
    let result10 = 0;

    let chars = value26.split('');
    for (let index in chars) {
        let indexI = parseInt(index);
        let charInChars = chars[chars.length - indexI - 1];
        let value = charInChars.charCodeAt(0) - 64
        result10 = result10 + value * Math.pow(26, indexI);
    }

    return result10;
}
function _10to26(value10) {
    let result26 = '';

    let mod = 0;
    while (value10 > 26) {
        mod = value10 % 26;
        result26 = String.fromCharCode(mod + 64) + result26;

        value10 = Math.round(value10 / 26);
    }
    if (value10 > 0) {
        result26 = String.fromCharCode(value10 + 64) + result26;
    }

    return result26;
}
function getRange(excelRange) {
    const regex = /([A-Z]+)([\d]+):([A-Z]+)([\d]+)/gm;
    let result = [...excelRange.matchAll(regex)];

    let range = {
        left: result[0][1],
        top: result[0][2],
        width: result[0][3],
        height: result[0][4]
    }
    return range;
}
function getCell(excelRange) {
    const regex = /([A-Z]+)([\d]+)/gm;
    let result = [...excelRange.matchAll(regex)];

    let cell = {
        col: result[0][1],
        row: result[0][2]
    }
    return cell;
}

function h(type, props, children) {
    return React.createElement(type, props, children)
}
function f(children) {
    return React.createElement(React.Fragment, {}, children)
}


export class SheetsList extends React.Component {

    constructor(props) {
        super(props);

        this.state = {
            nowSheet: props.content[0]
        }

        this.onSheetNameClick = this.onSheetNameClick.bind(this);
    }

    onSheetNameClick(e) {
        let value = e.target.getAttribute("data-value")
        if (this.props.onChange !== undefined) {
            this.props.onChange(value);
        }
        this.setState({nowSheet: value});
    }

    render() {
        let SheetNames = this.props.content;

        let Block = SheetNames.map((value, key) => {
            let className = "btnSheetName";
            if (value === this.state.nowSheet) {
                className = className + " btnSheetNameActive";
            }
            return h("div", {key: key, className: className, onClick: this.onSheetNameClick, "data-value": value}, value);
        })

        return h("div", {className: "SheetsNameList"}, Block);
    }

}

export class SheetsGrid extends React.Component {

    constructor(props) {
        super(props);

        this.state = {
            table: React.createElement( React.Fragment, {})
        }

        this.getCellContent = this.getCellContent.bind(this);
        this.createTable = this.createTable.bind(this);
    }

    getCellContent(col, row) {
        let result = '';
        let excelCol = _10to26(col);
        let excelCellName = excelCol + row;

        let value = this.props.content[excelCellName];
        if (value !== undefined) {
            result = value.w;
        }

        return result;
    }

    createTable() {
        let ref = this.props.content["!ref"];
        let excelRange = getRange(ref);

        let jsRange = {
            left: _26to10(excelRange.left),
            top: excelRange.top,
            width: _26to10(excelRange.width),
            height: excelRange.height
        }

        let thBlock = [];
        thBlock.push( h("th", {key:"col0 row0", className:"sheetsGridTH col0 row0"}) );
        for (let i = jsRange.left; i <= jsRange.width; i++) {
            let thTitle = _10to26(i);
            thBlock.push( h("th", {key:"col" + i + "_row0", className:"sheetsGridTH col" + i + " row0"}, thTitle) );
        }


        let tbodyBlock = [];
        for (let j = jsRange.top; j <= jsRange.height; j++) {
            let trBlock = [];
            trBlock.push( h("td", {key:"col0_row" + j, className: "sheetsGridTD col0 row" + j }, j) );

            for (let i = jsRange.left; i <= jsRange.width; i++) {
                let contentCell = this.getCellContent(i, j);
                trBlock.push( h ( "td", {key:'col' + i + '_row' + j, className:"sheetsGridTD col" + i + " row" + j }, contentCell) );
            }

            tbodyBlock.push( h("tr", {key:"row" + j },trBlock) );
        }

        return f( [
            h("thead", {key:"thead"},
                h("tr", {}, thBlock)
            ),
            h("tbody", {key:"tbody"}, tbodyBlock)
        ]);
    }

    componentDidMount() {
        let newTable = this.createTable();
        this.setState({table: newTable});
    }

    componentDidUpdate(prevProps, prevState, snapshot) {
        if (prevProps.content !== this.props.content) {
            let newTable = this.createTable();
            this.setState({table: newTable});
        }
    }


    render() {
        return h("table", {className: "sheetsGrid"}, this.state.table);
    }
}

export class ExcelGrid extends React.Component {

    constructor(props) {
        super(props);

        this.state = {
            content: undefined,
            nowSheet: ''
        }

        this.sheetChange = this.sheetChange.bind(this);
    }

    componentDidMount() {
        let content = this.props.excelJson;
        this.setState({
            content: content,
            nowSheet: content.SheetNames[0]
        });
    }

    sheetChange(nowSheet) {
        this.setState({nowSheet: nowSheet});
    }

    render() {
        // let ExcelGridClassNames = {
        //     ExcelGrid: {
        //         Container: {
        //             css: "gridContainer"
        //         }
        //     },
        //     SheetsList: {},
        //     SheetsGrid: {}
        // }
        //
        // let containerProps = {};
        // if (ExcelGridClassNames.ExcelGrid.Container.style !== undefined) {
        //     containerProps["style"] = ExcelGridClassNames.ExcelGrid.Container.style
        // } else if (ExcelGridClassNames.ExcelGrid.Container.css !== undefined) {
        //     containerProps["className"] = ExcelGridClassNames.ExcelGrid.Container.css
        // }

        let sheetsListClassName = {};
        let sheetsGridClassName = {};

        let Block = "";

        if (this.state.content !== undefined) {
            Block = [];
            Block.push( h(SheetsList, { key: "SheetsList", className: sheetsListClassName, content: this.state.content.SheetNames, onChange: this.sheetChange } ) );
            Block.push( h(SheetsGrid, { key: "SheetsGrid", className: sheetsGridClassName, content: this.state.content.Sheets[this.state.nowSheet] } ) );
        }

        return h("div", {className: "excelGridContainer"},  Block);
    }
}

