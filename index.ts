import fs from "fs"
import path from "path"

import XLSX from "xlsx"

console.error = function () {}

const dict = [
	"Январь",
	"Февраль",
	"Март",
	"Апрель",
	"Май",
	"Июнь",
	"Июль",
	"Август",
	"Сентябрь",
	"Октябрь",
	"Ноябрь",
	"Декабрь",
]
interface IJsonOutput {
	[name: string]: string
}

;(function CountXlsx(): void {
	try {
		const sumFinalArray: (string | number)[] = []
		fs.readdirSync("./tables", {
			encoding: "utf8",
		})
			.filter((file) => {
				return new Promise((resolve, reject) => {
					fs.stat("./tables/" + file, (err, stats) => {
						resolve(new Date(stats.birthtime).getTime())
					})
				})
			})
			.forEach((file) => {
				if (path.extname(file).toLowerCase() !== ".xlsx") return
				const JsonTable = parseExcel("./tables/" + file)
				let concatArray
				const sumArray: number[] = []
				JsonTable.forEach((row) => {
					const filtered = Object.keys(row).filter(
						(isNum) => !isNaN(isNum as unknown as number)
					)
					for (let i = 0; i <= 23; i++) {
						if (isNaN(sumArray[i])) sumArray[i] = 0
						filtered.forEach((e) => {
							if (Number(e) !== i) return
							sumArray[i] += Number(row[e])
						})
					}
				})

				sumFinalArray.push(...sumArray, "")
				console.info(
					"Файл '" + path.parse(file).name + "' обработан успешно!"
				)
			})
		writeExcel("./output.xlsx", [sumFinalArray])
	} catch (e: any) {
		console.warn("Произошла ошибка при обработке файла")
		throw new Error(e)
	}
})()

function parseExcel(filePath: string): any[] {
	const workBook = XLSX.readFile(filePath)

	let name = workBook.SheetNames[0]

	return XLSX.utils.sheet_to_json(workBook.Sheets[name])
}
function writeExcel(
	filePath: string,
	list: any[],
	sheetName = "Выходные данные"
) {
	const workBook = XLSX.utils.book_new()

	XLSX.utils.book_append_sheet(
		workBook,
		XLSX.utils.aoa_to_sheet(list),
		sheetName
	)

	XLSX.writeFile(workBook, filePath)
}
