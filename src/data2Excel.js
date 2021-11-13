/* eslint-disable */
import ExcelJs from 'exceljs'
import { saveAs } from "file-saver"
import axios from 'axios'

class ClassExcelUtils {

	constructor(filename) {
		this.percentage = 0
		this.imageCache = {}
		this.wb = new ExcelJs.Workbook()
		this.filename = filename
		this.supportImgExts = ['jpeg', 'png', 'gif']
	}

	getPercentage() {
		return this.percentage
	}

	addSheet(sheetName, metaData, datas) {
		sheetName = sheetName || 'Sheet1'
		if (!this.wb) {
			throw new Error(`workbook's instance not exist`)
		}
		if (metaData == null || metaData.length == 0) {
			throw new Error('metaData should not be empty')
		}
		console.time(this.filename + 'excel执行耗时')
		return new Promise((resolve, reject) => {
			this.createSheet(this.wb, sheetName, metaData, datas).then(res => {
				console.timeEnd(this.filename + 'excel执行耗时')
				resolve(res)
			}).catch(error => {
				reject(error)
			})
		})
	}

	async exportExcel() {
		var filename = this.filename || ''
		const wb = this.wb
		var res = await wb.xlsx.writeBuffer()
		const blob = new Blob([res], {
			type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
		});
		// 回收内存
		this.imageCache = {}
		return saveAs(blob, filename + new Date().getHours() + ':' +
			new Date().getMinutes() + ':' + new Date().getSeconds() + '.xlsx');
	}

	async createSheet(workbook, sheetName, metaData, datas) {
		const ws = workbook.addWorksheet(sheetName)
		this.setCellWidth(ws, metaData)

		let allFieldItems = this.traverseTree({ children: metaData }, true).map(item => {
			if (item.excel) {
				return { ...item, excel_sort: item.excel.sort }
			}
		}).filter(item => item != null)
		let colfieldItems = allFieldItems.filter(item => !item.children || item.children.length == 0)
		let fields = colfieldItems.sort((a, b) => a.excel_sort - b.excel_sort).map(item => item.prop)
		var headerLabelArrs = this.getHeaderLabelArrs({ children: metaData })
		var headerArr = this.getHeaderArr(headerLabelArrs, allFieldItems)
		var titleHeight = headerArr.length

		for (var i = titleHeight - 1; i >= 0; i--) {
			var header = headerArr[i]
			datas.unshift(header)
		}
		let metaDataMap = {}
		allFieldItems.forEach(item => {
			var prop = item.prop
			metaDataMap[prop] = item
		})
		var that = this
		return new Promise((resolve, reject) => {
			try {
				that.insertDataToSheet(workbook, ws, titleHeight, fields, metaDataMap, datas).then(res => {
					that.setStyle(ws, fields, metaDataMap)
					that.setRowHeight(ws, datas.length, titleHeight, metaDataMap)
					that.mergeCells(ws, headerArr)
					resolve(true)
				})
			} catch (error) {
				reject(error)
			}
		})
	}

	/**
	 * 遍历树节点
	 * @param {*} node 根节点
	 * @param {*} needParent 输出结果是否需要父级
	 */
	traverseTree(node, needParent) {
		if (!node) {
			return;
		}
		let stack = [];
		let arr = [];
		let tmpNode;
		stack.push(node);
		while (stack.length) {
			tmpNode = stack.shift();
			if (tmpNode && needParent) arr.push(tmpNode)
			if (tmpNode && tmpNode.children && tmpNode.children.length) {
				tmpNode.children.reverse().map(item => stack.unshift(item))
			} else if (!needParent) {
				arr.push(tmpNode)
			}
		}
		return arr
	}

	async insertDataToSheet(workbook, sheet, titleHeight, fields, metaDataMap, datas) {
		var total = datas.length
		for (var i in datas) {
			await this.createCells(workbook, sheet, i, titleHeight, fields, metaDataMap, datas[i]).catch(error => console.log('创建单元格异常:' + error))
			var cal = (Number(i) + 1) / total * 100
			this.percentage = cal.toFixed(2)
		}
	}

	// 样式相关待补充
	setStyle(sheet, header, metaDataMap) {
		for (var i in header) {
			var key = header[i]
			sheet.getColumn(Number(i) + 1).alignment = { vertical: 'middle', horizontal: 'center' }
			var defineStyle = metaDataMap[key].excel
			if (defineStyle) {
				for (var config in defineStyle) {
					sheet.getColumn(Number(i) + 1)[config] = defineStyle[config]
				}
			}
		}
	}

	setCellWidth(sheet, metaData) {

	}

	setRowHeight(sheet, rowNum, titleHeight, metaDataMap) {
		var rowHeight = 30
		for (var field in metaDataMap) {
			var item = metaDataMap[field]
			if (item.excel && item.excel.height) {
				rowHeight = Math.max(rowHeight, item.excel.height)
			}
		}
		for (var i = titleHeight + 1; i <= rowNum; i++) {
			const row = sheet.getRow(i)
			row.height = Number(rowHeight)
		}
	}


	async createCells(workbook, sheet, rowNum, titleHeight, fields, metaDataMap, data) {
		var arr = []
		if (rowNum < titleHeight) {
			// 填充空白
			sheet.addRow(data)
			return
		}
		for (var i in fields) {
			var field = fields[i]
			var prop = metaDataMap[field]
			if (!prop) continue
			var val = data[field]
			var type = prop.type || 'default'
			if (prop.formatter && prop.formatter instanceof Function) {
				val = prop.formatter(data)
			}
			if (type === 'image') {
				await this.createImageCell(workbook, sheet, i, rowNum, prop.excel, val).catch(error => console.log('导出图片异常' + error))
				val = ''
			} else {

			}
			arr.push(val)
		}
		sheet.addRow(arr)
	}
	async createImageCell(workbook, sheet, colNum, rowNum, prop, val) {
		if (!val) return
		var imgObj = this.imageCache[val]
		if (!imgObj) {
			imgObj = await this.getImgObj(val)
			this.imageCache[val] = imgObj
		}
		if (imgObj) {
			const imageId = workbook.addImage(imgObj);
			var rowNum1 = Number(rowNum)
			var tmpProp = Object.assign({}, prop)
			if (tmpProp.width) {
				tmpProp.width = tmpProp.width / 2
			}

			sheet.addImage(imageId, {
				tl: { row: parseInt(rowNum1), col: parseInt(colNum) },
				br: { row: parseInt(rowNum1) + 1, col: parseInt(colNum) + 1 },
				ext: prop,
				editAs: 'undefined',
			});
			val = ''
		}
	}

	async createNumberCell(sheet, colNum, rowNum, prop, val) {

	}
	async createStringCell(sheet, colNum, rowNum, prop, val) {

	}

	mergeCells(sheet, headerArr) {
		var merges = this.calMerges(headerArr)
		merges.forEach(item => {
			sheet.mergeCells(item)
		})
	}

	// TODO 测试案例数量不够，可能存在bug
	calMerges(headerArr) {
		var height = headerArr.length
		var colorMap = []
		for (var h = 0; h < height; h++) {
			var initArr = Array(headerArr[0].length).fill(0)
			colorMap.push(initArr)
		}
		var merges = []
		for (let i = 0; i < height; i++) {
			var header = headerArr[i]
			var len = header.length
			for (var j = 0; j < len; j++) {
				if (!header[j]) continue
				if (colorMap[i][j] == 1) continue
				var startRow = i
				var startCol = j
				// 计算横向最大值
				var tempCol = startCol
				do {
					if (headerArr[startRow][tempCol] && tempCol != j) {
						break
					}
					tempCol++
				} while (tempCol < len)
				// 计算纵向最大值
				var tempRow = startRow
				do {
					if (headerArr[tempRow][startCol] && tempRow != i) {
						break
					}
					tempRow++
				} while (tempRow < height)
				var xy = this.rangeExist(startRow, tempRow, startCol, tempCol, headerArr, colorMap)
				// console.log('染色====>'+startRow+','+startCol+'-'+xy[0]+','+xy[1])
				this.dye(startRow, xy[0], startCol, xy[1], colorMap)
				if (startRow === xy[0] && startCol === xy[1]) {

				} else {
					if (xy[0] < height && xy[1] < len) {
						merges.push(this.getMergeByXy(startRow, startCol, xy[0], xy[1]))
					}
				}
			}
		}
		return merges
	}
	getMergeByXy(x1, y1, x2, y2) {
		// 仅支持 24 * 24 行
		var firstChar = this.calChar(y1, '')
		var secondChar = this.calChar(y2, '')
		var merge = `${firstChar}${x1 + 1}:${secondChar}${x2 + 1}`
		return merge
	}
	calChar(num, str = '') {
		let a = num % 26
		let b = Math.floor(num / 26) - 1
		str = String.fromCharCode(a + 65) + str
		if (b === -1) return str
		return calChar(b, str)
	}

	dye(startRow, endRow, startCol, endCol, colorMap) {
		for (var i = startRow; i <= endRow && i < colorMap.length; i++) {
			for (var j = startCol; j <= endCol && j < colorMap[i].length; j++) {
				colorMap[i][j] = 1
			}
		}
	}


	rangeExist(startRow, endRow, startCol, endCol, headers, colorMap) {
		var i = startCol
		var j = startRow
		for (var j = startRow; j < endRow; j++) {
			for (var i = startCol; i < endCol; i++) {
				if (i === startCol && j === startRow) continue
				if (headers[j][i] || colorMap[j][i] === 1) {
					// 判断每个轴有没有数据
					var tmpi = startCol
					var tmpj = startRow
					// x轴 
					do {
						if (headers[startRow][tmpi] && tmpi != i) {
							break
						}
						tmpi++
					} while (tmpi < i - 1)
					// 计算纵向最大值
					do {
						if (headers[tmpj][startCol] && tmpj != j) {
							break
						}
						tmpj++
					} while (tmpj < j - 1)
					return [tmpj, tmpi]
				}
			}
		}

		return [endRow - 1, endCol - 1]
	}
	// 层序生成每一层标题数组
	getHeaderLabelArrs(root) {
		let que = [], newque = [], result = [];
		if (root !== null) {
			que.push(root);
		} else {
			return [0];
		}
		do {
			let sum = [];
			que.forEach(function (item) {
				sum.unshift(item.label);
			})
			result.push(sum);
			while (que.length != 0) {
				let node = que.shift();
				var ch = node.children || []
				// newque.push('')
				ch.forEach(item => {
					newque.push(item);
				})
			}
			let temp = newque;
			newque = que;
			que = temp;
		} while (que.length != 0);
		result.shift()
		return result;
	}
	getHeaderArr(headerLabelArrs, allFieldItems) {
		var sorts = allFieldItems.map(item => item.excel_sort).sort((a, b) => b - a)
		var width = sorts[0]
		var fieldMap = {}
		allFieldItems.forEach(item => {
			var label = item.label
			fieldMap[label] = item
		})
		var headerDatas = []
		for (let i = 0; i < headerLabelArrs.length; i++) {
			// 必须填充 空字符串，否则exceljs会有问题
			var emptyArr = new Array(width).fill('')
			headerDatas.push(emptyArr)
		}
		for (var i in headerLabelArrs) {
			var headerValArr = headerLabelArrs[i]
			headerValArr.forEach(item => {
				if (fieldMap[item]) {
					var col = fieldMap[item].excel_sort
					var val = item || ''
					headerDatas[i][col - 1] = val
				} else {
					headerDatas[i][col - 1] = ''
				}
			})
		}
		return headerDatas
	}

	async getImgObj(url) {
		var res
		try {
			res = await axios.get(url, { responseType: 'arraybuffer' })
		} catch (Error) {
			return null
		}
		if (!res || res.status != 200) {
			return null
		}
		var ext = url.substr(url.lastIndexOf(".")).toLowerCase();
		ext = ext.split('?')[0].replaceAll('.', '')
		var buffer = res.data
		if (res.data) {
			return { buffer, extension: ext }
		}
		return null
	}

}

export { ClassExcelUtils }