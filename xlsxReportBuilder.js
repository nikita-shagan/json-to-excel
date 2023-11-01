const XLSX = require("xlsx-js-style")


// You can define styles as json object
const styles = {
	headerGrey: {
		fill: {
			type: "pattern",
			patternType: "solid",
			fgColor: "FFDEE6EF",
		},
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
		font: {
			color: "FF000000",
			name: "Arial",
			size: 10,
			bold: true,
			underline: false,
		},
		alignment: {
			vertical: "top",
			horizontal: "center",
		},
		wrapText: true,
	},
	cellNum: {
		numberFormat: "#,##0",
		font: { name: "Arial", size: 10 },
		alignment: {
			vertical: "top",
			horizontal: "right",
		},
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
	},
	cellPercent: {
		numberFormat: "0.0%",
		font: { name: "Arial", size: 10 },
		alignment: {
			vertical: "top",
			horizontal: "right",
		},
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
	},
	cellCenter: {
		alignment: {
			vertical: "top",
			horizontal: "center",
		},
		numberFormat: "0",
		font: { name: "Arial", size: 10 },
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
	},
	cellQuantity: {
		alignment: {
			vertical: "top",
			horizontal: "center",
		},
		numberFormat: "0.0##",
		font: { name: "Arial", size: 10 },
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
	},
	cellDate: {
		alignment: {
			vertical: "top",
			horizontal: "center",
		},
		numberFormat: "dd.mm.yy",
		font: { name: "Arial", size: 10 },
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
	},
	cellDateTime: {
		alignment: {
			vertical: "top",
			horizontal: "center",
		},
		numberFormat: "yyyy-mm-dd hh:mm",
		font: { name: "Arial", size: 10 },
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
	},
	cellDefault: {
		alignment: {
			vertical: "top",
		},
		font: { name: "Arial", size: 10 },
		border: {
			top: { style: "thin", color: "#404040" },
			bottom: { style: "thin", color: "#404040" },
			left: { style: "thin", color: "#404040" },
			right: { style: "thin", color: "#404040" },
		},
	},
};


function getXLSX(data) {
	//Array of objects representing heading rows (very top)
	// const heading = [
	// 	[
	// 		{ value: "Заголовок 1", style: styles.headerDark },
	// 		{ value: "b1", style: styles.headerDark },
	// 		{ value: "c1", style: styles.headerDark },
	// 	],
	// 	["Заголовок 2", "пояснение", "еще"], // <-- It can be only values
	// ];
	// The data set should have the following shape (Array of Objects)
	// <- Notice that this is an array. Pass multiple sheets to create multi sheet report
	//     {
	//       name: "mshop", // <- Specify sheet name (optional)
	//       heading: heading, // <- Raw heading array (optional)
	//       specification: specification1, // <- Report specification
	//       data: data.sheet1, // <-- Report data
	//     },
	//   ]);
	const report = buildExport(data.sheets);
	if (report) return XLSX.write(report, { type: 'buffer' })
}


function convertColor(color) {
  if (typeof color === "string") {
    return color.length === 7
      ? {rgb: color.slice(1)}
      : {rgb: color.slice(2)}
  }
  return color
}


function convertStyle(style) {
  if (!style) return
  const newStyle = {}
  Object.keys(style).forEach(key => {
    if (key === "alignment") {
      newStyle.alignment = {}
      Object.keys(style.alignment).forEach(alignmentProp => {
        if (["horizontal", "vertical"].includes(alignmentProp)) {
          if (["top", "center", "bottom"].includes(style.alignment[alignmentProp])) {
            newStyle.alignment[alignmentProp] = style.alignment[alignmentProp]
          }
        }
        if (["wrapText", "textRotation"].includes(alignmentProp)) {
          newStyle.alignment[alignmentProp] = style.alignment[alignmentProp]
        }
      })
    }

    if (key === "border") {
      newStyle.border = {}
      Object.keys(style.border).forEach(borderProp => {
        if (["top", "bottom", "left", "right", "diagonal"].includes(borderProp)) {
          newStyle.border[borderProp] = {...style.border[borderProp]}
          newStyle.border[borderProp].color = convertColor(style.border[borderProp].color)
        }
      })
    }

    if (key === "fill") {
      newStyle.fill = {}
      Object.keys(style.fill).forEach(fillProp => {
        if (fillProp === "patternType") {
          newStyle.fill.patternType = style.fill.patternType
        }
        if (["fgColor", "bgColor"].includes(fillProp)) {
          newStyle.fill[fillProp] = convertColor(style.fill[fillProp])
        }
      })
    }

    if (key === "font") {
      newStyle.font = {}
      Object.keys(style.font).forEach(fontProp => {
        if (["bold", "italic", "name", "strike", "underline"].includes(fontProp)) {
          newStyle.font[fontProp] = style.font[fontProp]
        }
        if (fontProp === "color") {
          newStyle.font.color = convertColor(style.font.color)
        }
        if (fontProp === "size") {
          newStyle.font.sz = String(style.font.size)
        }
        if (fontProp === "style") {
          newStyle.font.bold = style.font.style === "bold"
        }
      })
    }

    if (key === "numberFormat") {
      newStyle.numFmt = style.numberFormat
    }
  })
  return newStyle
}


function getMergesObjects(merges){
  const mergesObjects = []
  Object.keys(merges).forEach(colNum => {
    const rowNums = merges[colNum];
    let mergingRowNumStart = rowNums[0].rowNum
    for (let i = 1; i < rowNums.length; i++) {
      if (rowNums[i].rowNum - rowNums[i - 1].rowNum > 1 || i === rowNums.length - 1) {
        mergesObjects.push({
          s: { r: mergingRowNumStart, c: Number(colNum) },
          e: { r: mergingRowNumStart + rowNums[i - 1].offset, c: Number(colNum) }
        })
        mergingRowNumStart = rowNums[i].rowNum
      }
    }
    mergesObjects.push({
      s: { r: mergingRowNumStart, c: Number(colNum) },
      e: { r: mergingRowNumStart + rowNums[rowNums.length - 1].offset, c: Number(colNum) }
    })
  })
  return mergesObjects
}


function getCell({ value, baseStyle, styleFromFunction, beforeWriteStyle }) {
  baseStyle = convertStyle(baseStyle)
  styleFromFunction = convertStyle(styleFromFunction)
  beforeWriteStyle = convertStyle(beforeWriteStyle)
  let v = value || '';
  let t = (typeof value)[0]
  let s = baseStyle || {}

  if (styleFromFunction) {
    s = {...s, ...styleFromFunction}
  }
  if (beforeWriteStyle) {
    s = {...s, ...beforeWriteStyle}
  }
  if (typeof v.getMonth === 'function') {
    v = 25569 + ((value.getTime()) / (1000 * 60 * 60 * 24))
    t = 'n'
  }

  return { v, t, s }
}


function buildExport(sheets) {
  const workbook = XLSX.utils.book_new();
  sheets.forEach(sheet => {
		if (!sheet.specification) return;
    let sheetTable = [];
    let columnsWidths = [];
    let rowHeights = [];

    // Заполняем заголовок страницы из аттрибута "heading"
    let heading = sheet.heading || [];
		heading.forEach((row) => {
			if (row instanceof Array) {
        let headingRow = []
				row.forEach((value) => {
          const cellData = {}
					if (value && typeof value === "object") {
            cellData.value = value.value
            cellData.baseStyle = value.style
					} else {
            cellData.value = value
					}
          const cell = getCell(cellData)
					headingRow.push(cell)
				});
        sheetTable.push(headingRow)
			}
		});

    // Устанавливаем заголовки столбцов
    const headerRow = []
    Object.keys(sheet.specification).forEach((colName) => {
			const spec = sheet.specification[colName];
			const cell = getCell({
        value: spec.displayName,
        baseStyle: styles[spec.headerStyle]
      });
			if (spec.width) columnsWidths.push({wch: Math.floor(Number(spec.width) * 1.1)});
      headerRow.push(cell)
		});
    sheetTable.push(headerRow)

    // Заполняем таблицу данными из аттрибута "data"
    const merges = {}
    const headingsHeight = sheetTable.length
    sheet.data.forEach((row, rowNum) => {
      let rowData = []
      rowHeights.push({hpt: 12.85})
      Object.keys(sheet.specification).forEach((colName, colNum) => {
        let value = row[colName];
        const spec = sheet.specification[colName]
				const styleName = spec.cellStyle
        const baseStyle = styleName && styles[styleName] || styles.cellDefault
        let beforeWriteStyle;
				let styleFromFunction;
        if (spec.styleFunc && typeof spec.styleFunc === "function") {
					styleFromFunction = spec.styleFunc(value, row)
				}
				if (spec.beforeWrite && typeof spec.beforeWrite === "function") {
					const res = spec.beforeWrite(value, {
						dataset: sheet.data,
						row,
						rowno: rowNum,
						colname: colName,
					});
          beforeWriteStyle = res.style;
					value = res.newvalue;
          if (res.merges) {
            if (!merges.hasOwnProperty(colNum)) {
              merges[colNum] = []
            }
            merges[colNum].push({rowNum: rowNum + headingsHeight - 1, offset: res.merges.up})
          }
				}
        const cell = getCell({ value, baseStyle, styleFromFunction, beforeWriteStyle })
        rowData.push(cell)
      })
      sheetTable.push(rowData)
      rowData = []
    })

    // Создаем и форматируем страницу
    const worksheet = XLSX.utils.aoa_to_sheet(sheetTable);
    worksheet['!cols'] = columnsWidths;
    worksheet['!rows'] = rowHeights;
    worksheet['!merges'] = getMergesObjects(merges);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
    sheetTable = []
  })

	return workbook;
}

module.exports = { getXLSX, styles };
