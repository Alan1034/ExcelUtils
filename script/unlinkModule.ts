/*
 * @Author: 陈德立*******419287484@qq.com
 * @Date: 2023-11-10 09:39:59
 * @LastEditTime: 2025-08-25 19:27:02
 * @LastEditors: 陈德立*******419287484@qq.com
 * @Github: https://github.com/Alan1034
 * @Description: unlink处理流程
 * @FilePath: \VueNpmFrame\script\unlinkModule.ts
 * 
 */
import path from "path";
import fs from 'fs'
import chalk from 'chalk'
import { fileURLToPath } from 'url'
const __filenameNew = fileURLToPath(import.meta.url)
const __dirnameNew = path.dirname(__filenameNew)
const updateMain = () => {
  try {
    const packageTxt = fs.readFileSync(path.join(__dirnameNew, '../package.json'), 'utf8');
    const packageJson = JSON.parse(packageTxt);
    const umdDir = "./dist/index.umd.cjs"
    const esDir = "./dist/index.js"
    const types = "./dist/index.d.ts"
    packageJson.type = "module"
    packageJson.main = umdDir
    packageJson.module = esDir
    packageJson.exports = {
      ".": {
        "types": types,
        "import": esDir,
        "require": umdDir,
      },
      "./style": "./dist/index.css",
    }
    packageJson.typings = types
    packageJson.files = [
      "/dist"
    ]
    const versionData = JSON.stringify(packageJson, null, 2);
    fs.writeFileSync(path.join(__dirnameNew, '../package.json'), versionData, 'utf8');
    console.log(chalk.green.bold('修改入口配置成功！当前入口路径为为：' + umdDir));
  } catch (error) {
    console.log(chalk.red.bold('修改入口配置出错:', error.toString()));
  }

}

updateMain()