import path from "path";
import fs from 'fs'
import chalk from 'chalk'
import { fileURLToPath } from 'url'
const __filenameNew = fileURLToPath(import.meta.url)
const __dirnameNew = path.dirname(__filenameNew)
export const updateMain = (mainDir:string) => {
  try {
    const packageTxt = fs.readFileSync(path.join(__dirnameNew, '../package.json'), 'utf8');
    const packageJson = JSON.parse(packageTxt);

    delete packageJson.files
    delete packageJson.module
    delete packageJson.exports
    packageJson.typings = "./publish/index.d.ts"
    packageJson.main = mainDir
    const versionData = JSON.stringify(packageJson, null, 2);
    fs.writeFileSync(path.join(__dirnameNew, '../package.json'), versionData, 'utf8');
    console.log(chalk.green.bold('修改入口配置成功！当前入口路径为为：' + mainDir));
  } catch (error) {
    console.log(chalk.red.bold('修改入口配置出错:', error.toString()));
  }

}
