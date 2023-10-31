const XLSX = require("xlsx-js-style")
const { styles } = require("./styles")

const dateRegex = /\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d\.\d+([+-][0-2]\d:[0-5]\d|Z)/


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


function getCell({ value, baseStyle, styleFunc, beforeWriteStyle }) {
  let v = value || '';
  let t = (typeof value)[0]
  let s = baseStyle || {}
  if (styleFunc) {
    s = {...s, ...styleFunc}
  }
  if (beforeWriteStyle) {
    s = {...s, ...beforeWriteStyle}
  }
  if (dateRegex.test(value)) {
    const date = new Date(value);
    v = 25569 + ((date.getTime()) / (1000 * 60 * 60 * 24))
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
    const merges = []
    const headingsHeight = sheetTable.length
    let rowStart = sheetTable.length
    sheet.data.forEach((row, rowNum) => {
      let rowData = []
      rowHeights.push({hpt: 12.85})
      Object.keys(sheet.specification).forEach((colName) => {
        let value = row[colName];
        const spec = sheet.specification[colName]
				const styleName = spec.cellStyle
        const baseStyle = styleName && styles[styleName] || styles.cellDefault
        let res = {};
				let sf;
        if (spec.styleFunc && typeof spec.styleFunc === "function") {
					sf = spec.styleFunc(value, row);
				}
				if (spec.beforeWrite && typeof spec.beforeWrite === "function") {
					res = spec.beforeWrite(value, {
						dataset: sheet.data,
						row,
						rowNum,
						colName,
					});
					value = res.newvalue;
          if (res.allSame === false) {
            res.fields.forEach((field, index) => {
              merges.push({
                s: { r: rowStart, c: index },
                e: { r: rowNum + headingsHeight - 1, c: index }
              })
            })
            rowStart = rowNum + headingsHeight
          }
				}
        const cell = getCell({ value, baseStyle, styleFunc: sf, beforeWriteStyle: res.style})
        rowData.push(cell)
      })
      sheetTable.push(rowData)
      rowData = []
    })

    // Создаем и форматируем страницу
    const worksheet = XLSX.utils.aoa_to_sheet(sheetTable);
    worksheet['!cols'] = columnsWidths;
    worksheet['!rows'] = rowHeights;
    worksheet['!merges'] = merges;
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
    sheetTable = []
  })

	return workbook;
}

module.exports = { getXLSX, styles };
