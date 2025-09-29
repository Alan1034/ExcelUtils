/**
 * @format
 * @Author: 陈德立*******419287484@qq.com
 * @Date: 2025-09-29 18:42:40
 * @LastEditTime: 2025-09-29 18:55:20
 * @LastEditors: 陈德立*******419287484@qq.com
 * @Github: https://github.com/Alan1034
 * @Description:
 * @FilePath: /ExcelUtils/src/Excel2data.ts
 */

import xlsx from "xlsx";

class ClassExcelReadUtils {
  constructor() {}

  formatNum(s) {
    const l = `${s}`.split(".")[0].split("").reverse();
    let t = "";
    for (let i = 0; i < l.length; i += 1) {
      t += l[i] + ((i + 1) % 3 === 0 && i + 1 !== l.length ? "," : "");
    }
    return t.split("").reverse().join("");
  }

  readXlsx(fileBase64) {
    // 将base64转换为blob
    const dataURLtoBlob = (dataurl) => {
      const arr = dataurl.split(",");
      const mime = arr[0].match(/:(.*?);/)[1];
      const bstr = atob(arr[1]);
      let n = bstr.length;
      const u8arr = new Uint8Array(n); // eslint-disable-next-line no-plusplus
      while (n--) {
        u8arr[n] = bstr.charCodeAt(n);
      }
      return new Blob([u8arr], { type: mime });
    };
    const blob = dataURLtoBlob(fileBase64);
    const fileReader = new FileReader();
    fileReader.readAsBinaryString(blob);

    return fileReader;
  }

  dealData(fileBase64, getdata, type) {
    if (!fileBase64) {
      return;
    }
    const fileReader = this.readXlsx(fileBase64);
    fileReader.onload = (ev) => {
      let workbook = null;
      let datas = [];
      try {
        const data = ev.target.result;
        workbook = xlsx.read(data, {
          type: "binary",
        }); // 以二进制流方式读取得到整份excel表格对象 //  // 存储获取到的数据
      } catch (e) {
        // console.log('文件类型不正确');
      } // 表格的表格范围，可用于判断表头是否数量是否正确 // let fromTo = ''; // 遍历每张表读取
      for (const sheet in workbook.Sheets) {
        // if (workbook.Sheets.hasOwnProperty(sheet)) {
        if (Object.prototype.hasOwnProperty.call(workbook.Sheets, sheet)) {
          // fromTo = workbook.Sheets[sheet]['!ref'];
          // console.log(fromTo);
          datas = datas.concat(
            xlsx.utils.sheet_to_json(workbook.Sheets[sheet])
          ); // break; // 如果只取第一张表，就取消注释这行
        }
      }
      if (getdata && type) {
        // type为传入参数 // getdata为取值用的函数
        getdata(datas, type);
      }
      return datas;
    };
  }
}

export { ClassExcelReadUtils };
