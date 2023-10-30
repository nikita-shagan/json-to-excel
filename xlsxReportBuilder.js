const excel = require("excel4node");
const fs = require("fs");

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
	if (report) return report.writeToBuffer();
}

function set_cell({ ws, r1, c1, r2, c2, merged, value, baseStyle, styleFunc, beforeWriteStyle }) {
	let cell = ws.cell(r1, c1, r2, c2, merged);
	if (value == null) value = undefined;
	switch (typeof value) {
		case "number":
			cell.number(value);
			break;
		case "undefined":
			cell.string("");
			break;
		default:
			if (typeof value.getMonth === "function") {
				cell.date(value);
				break;
			}
			cell.string(String(value));
			break;
	}
	if (baseStyle) {
		cell.style(baseStyle);
	}
	if (styleFunc) {
		cell.style(styleFunc);
	}
	if (beforeWriteStyle) {
		cell.style(beforeWriteStyle); // beforeWrite идет в самом конце
	}
}

function buildExport(sheets) {
	let workbook = new excel.Workbook();
	let stylebook = {};
	Object.keys(styles).forEach(stylename => {
		let styledef = styles[stylename];
    stylebook[stylename] = workbook.createStyle(styledef);
	});
	sheets.forEach(sheet => {
		if (!sheet.specification) return;
		let worksheet = workbook.addWorksheet(sheet.name, {
			sheetView: { showGridLines: false },
			sheetFormat: { defaultRowHeight: 12.85 },
		});
		let heading = sheet.heading || [];
		let headrow = heading.length + 1;
		heading.forEach((r, rn) => {
			if (r instanceof Array) {
				r.forEach((val, cn) => {
					let m = { ws: worksheet, r1: rn + 1, c1: cn + 1 };
					if (val && typeof val === "object" && val.style) {
						m.value = val.value;
						m.style = val.style;
					} else {
						m.value = val;
					}
					set_cell(m);
				});
			}
		});
		Object.keys(sheet.specification).forEach((colname, colno) => {
			let spec = sheet.specification[colname];
			let cell = worksheet.cell(headrow, colno + 1).string(spec.displayName);
			if (stylebook[spec.headerStyle]) {
				cell.style(stylebook[spec.headerStyle]);
			}
			if (spec.width) worksheet.column(colno + 1).setWidth(Number(spec.width));
		});
		let merges = {}; // объединить ячейки в excel4node можно только 1 раз, поэтому будем копить объединения тут
		// в привязке к верхней левой ячейке
		sheet.data.forEach((row, rowno) => {
			row._row_number = rowno + 1;
			Object.keys(sheet.specification).forEach((colname, colno) => {
				let value = row[colname];
				let spec = sheet.specification[colname];
				let res = {};
				let sf;
				if (spec.styleFunc && typeof spec.styleFunc === "function") {
					sf = spec.styleFunc(value, row);
				}
				if (spec.beforeWrite && typeof spec.beforeWrite === "function") {
					res = spec.beforeWrite(value, {
						dataset: sheet.data,
						row,
						rowno,
						colname,
					});
					value = res.newvalue;
				}
				let m = {
					ws: worksheet,
					value,
					baseStyle: stylebook[spec.cellStyle],
					styleFunc: sf,
					beforeWriteStyle: res.style,
				};
				if (res.merges) {
					// если есть объединения, то откладываем на потом
					m.r1 = headrow + rowno + 1 - res.merges.up;
					m.c1 = colno + 1 - res.merges.left;
					m.r2 = headrow + rowno + 1;
					m.c2 = colno + 1;
					m.merged = true;
					merges["R" + m.r1 + "C" + m.c1] = m;
				} else {
					m.r1 = headrow + rowno + 1;
					m.c1 = colno + 1;
					set_cell(m);
				}
			});
		});
		// закончили все строчки, теперь надо пройтись по объединенным ячейкам
		// если объединения постепенно расширялись, то более поздние затирали более ранние
		// никакой проверки мы не проводим, кроме того, что объединения можно возвращать только влево и вверх из
		// beforeWrite
		Object.keys(merges).forEach(topLeftCell => {
			let m = merges[topLeftCell];
			set_cell(m);
		});
		worksheet.row(headrow).freeze();
	});
	return workbook;
}

module.exports = { getXLSX, styles };
