const fs = require("fs/promises");
const reportbuilder = require("./xlsxReportBuilder");
test();

async function test() {
	const filename = `test.xlsx`;
	let data = JSON.parse(await fs.readFile("data.json"));
	let storage = {};
	const fields = ["supplier_id", "supplier", "shops", "shops2", "selling_arts1", "revenue1", "sow1", "modified_at"];

	reportbuilder.styles.cellMoney = {
		numberFormat: "0",
		font: { name: "Arial", sz: 10 },
		alignment: {
			vertical: "top",
			horizontal: "right",
		},
		border: {
			top: { style: "thin", color: {rgb: "404040"} },
			bottom: { style: "thin", color: {rgb: "404040"} },
			left: { style: "thin", color: {rgb: "404040"} },
			right: { style: "thin", color: {rgb: "404040"} },
		},
	};

	const beforeWrite = function (value, { dataset, row, rowNum, colName }) {
		let nextstyle = {
			fill: {
				patternType: "solid",
				fgColor: {rgb: "d9d9d9"},
			},
		};
		let style = {
			fill: {
				patternType: "solid",
				fgColor: {rgb: "ffffff"},
			},
		};
		let important_color = {
			fill: {
				patternType: "solid",
				fgColor: {rgb: "d9d9d9"},
			},
			font: { color: {rgb: "ff0000"}, bold: true },
		};
		let important_white = {
			fill: {
				patternType: "solid",
				fgColor: {rgb: "ffffff"},
			},
			font: { color: {rgb: "ff0000"}, bold: true },
		};

		let curr = storage[rowNum];
		let important_row;
		if (curr && curr.row && !curr.row.selling_arts1 && curr.row.shops2 > 3) important_row = true;
		let currstyle;
		if (curr) {
			currstyle = important_row
				? curr.style.fill.fgColor.rgb === "ffffff"
					? important_white
					: important_color
				: curr.style;
		}
		if (curr && curr.repeat) {
			if (fields.includes(colName)) {
				return { newvalue: value, style: currstyle, merges: { left: 0, up: curr.repeat } };
			} else {
				return { newvalue: value, style: currstyle };
			}
		}
		if (curr && curr.repeat === 0) return { newvalue: value, style: currstyle };
		// если мы тут, значит это колонка А, начало строки
		storage[rowNum] = curr || { row, repeat: 0 };
		curr = storage[rowNum];
		// заново определяем, важная ли это строка
		if (curr && curr.row && !curr.row.selling_arts1 && curr.row.shops2 > 3) important_row = true;
		let prevno = rowNum - 1;
		if (!(prevno in storage)) {
			curr.repeat = 0;
			curr.style = style;
			curr.nextstyle = nextstyle;
			currstyle = important_row
				? curr.style.fill.fgColor.rgb === "ffffff"
					? important_white
					: important_color
				: curr.style;
			return { newvalue: value, style: currstyle }; // по умолчанию - ничего не делаем
		}
		let prev = storage[prevno];
		// если мы тут, значит это не первая строка
		// let fields = Object.keys(spec).filter(f => typeof spec[f].beforeWrite === "function");
		let all_same = true;
		for (const f of fields) {
			if (prev.row[f] !== row[f]) {
				if (
					row[f] &&
					typeof row[f].valueOf === "function" &&
					prev.row[f] &&
					typeof prev.row[f].valueOf === "function"
				) {
					if (prev.row[f].valueOf() !== row[f].valueOf()) {
						all_same = false;
						break;
					}
				} else {
					all_same = false;
					break;
				}
			}
		}
		if (all_same) {
			curr.repeat = prev.repeat + 1;
			curr.style = prev.style;
			curr.nextstyle = prev.nextstyle;
			currstyle = important_row
				? curr.style.fill.fgColor.rgb === "ffffff"
					? important_white
					: important_color
				: curr.style;
			if (fields.includes(colName)) {
				return { newvalue: value, style: currstyle, merges: { left: 0, up: curr.repeat } };
			} else {
				return { newvalue: value, style: currstyle };
			}
		}
		curr.style = prev.nextstyle;
		curr.nextstyle = prev.style;
		curr.repeat = 0;
		currstyle = important_row
			? curr.style.fill.fgColor.rgb === "ffffff"
				? important_white
				: important_color
			: curr.style;
		return { newvalue: value, style: currstyle, fields, allSame: all_same };
	};

	let spec = {
		supplier_id: {
			displayName: "id",
			cellStyle: "cellCenter",
			width: 4,
		},
		supplier: {
			displayName: "Поставщик", // <- Here you specify the column header
			cellStyle: "cellDefault",
			width: 20,
		},
		shops: {
			displayName: "Ликвид",
			cellStyle: "cellCenter",
			width: 7,
		},
		shops2: {
			displayName: "Магазинов",
			cellStyle: "cellCenter",
			width: 12,
		},
		selling_arts1: {
			displayName: "Артикулов 1",
			cellStyle: "cellCenter",
			width: 11,
		},
		revenue1: {
			displayName: "Выручка 1",
			cellStyle: "cellNum",
			width: 11,
		},
		sow1: {
			displayName: "Доля 1",
			cellStyle: "cellPercent",
			width: 7,
		},
		modified_at: {
			displayName: "Обновлено!",
			cellStyle: "cellDateTime",
			width: 15,
		},
		shop_id2: {
			displayName: "маг2",
			cellStyle: "cellCenter",
			width: 10,
		},
		selling_arts2: {
			displayName: "Артикулов 2",
			cellStyle: "cellCenter",
			width: 11,
		},
		revenue2: {
			displayName: "Выручка 2",
			cellStyle: "cellNum",
			width: 11,
		},
		sow2: {
			displayName: "Доля 2",
			cellStyle: "cellPercent",
			width: 7,
		},
	};
	let spec2 = {
		sid: {
			displayName: "sid",
			cellStyle: "cellCenter",
			width: 4,
		},
		sname: {
			displayName: "Поставщик", // <- Here you specify the column header
			cellStyle: "cellDefault",
			width: 15,
		},
		rootbc: {
			displayName: "Штрихкод",
			cellStyle: "cellCenter",
			width: 20,
		},
		name: {
			displayName: "Товар",
			cellStyle: "cellDefault",
			width: 35,
		},
		minprice: {
			displayName: "Цена мин.",
			cellStyle: "cellMoney",
			width: 10,
		},
		maxprice: {
			displayName: "Цена макс.",
			cellStyle: "cellMoney",
			width: 10,
		},
		lastsupp: {
			displayName: "Посл. поставщик",
			cellStyle: "cellDefault",
			width: 20,
		},
		lastprice: {
			displayName: "Посл. цена",
			cellStyle: "cellMoney",
			width: 10,
			styleFunc: (value, row) => {
				if (row && value > row.maxprice)
					return {
						fill: {
							patternType: "solid",
							fgColor: {rgb: "ffc773"},
						},
					};
			},
		},
		avgsell: {
			displayName: "Продажи шт.",
			cellStyle: "cellMoney",
			width: 12,
		},
		sumbought: {
			displayName: "Закупка",
			cellStyle: "cellMoney",
			width: 10,
		},
		shopnum: {
			displayName: "Магазинов",
			cellStyle: "cellCenter",
			width: 11,
		},
		aliases: {
			displayName: "Магазины",
			cellStyle: "cellDefault",
			width: 35,
		},
		category: {
			displayName: "Категория",
			cellStyle: "cellDefault",
			width: 20,
		},
	};
	Object.keys(spec).forEach(field => {
		spec[field].headerStyle = "headerGrey";
		spec[field].beforeWrite = beforeWrite;
	});
	Object.keys(spec2).forEach(field => {
		spec2[field].headerStyle = "headerGrey";
		// spec[field].beforeWrite = beforeWrite;
	});
	data.sheets[0].specification = spec;
	data.sheets[1].specification = spec2;
	console.log("data.json loaded");
	let buffer = await reportbuilder.getXLSX(data);
	console.log("buffer received");
	await fs.writeFile(filename, buffer);
	console.log("xlsx written");
}
