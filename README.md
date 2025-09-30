# ExcelUtils使用指南

## 简介 
一个前端表格处理组件
### 使用指南

#### 使用引入

```
import { ExcelUtils, ExcelReadUtils  } from 'excel-utils-bt'  
```

#### 简单使用

```
// 1、新建一个 实例
let el = new ExcelUtils('采购-拣货单')
// 2、调用实例函数，向实例添加表和数据
var data = []
el.addJsonToSheet('采购-拣货单', data)
// 3、调用导出函数
//  exportExcel()函数为异步调用，可用 el.exportExcel().then(res => {}).catch(err => {})
el.exportExcel()
```

#### 多个表格

```
// 1、新建一个 实例
let el = new ExcelUtils('采购-拣货单')
// 2、调用实例函数，向实例添加表和数据。多个表格重复调用即可，表格名重复则覆盖
var data = []
el.addJsonToSheet('采购-拣货单1', data)
el.addJsonToSheet('采购-拣货单2', data)
// 3、调用导出函数
el.exportExcel()
```

#### 导出模板

```
// 1、新建一个 实例
let el = new ExcelUtils('采购-拣货单')
// 2、调用实例函数，向实例添加表和数据
let aoa = ['国内物流公司', '国内运单号', '客户公司', 'SKU码', "计划数量"]
el.addAoaToSheet('采购-拣货单', aoa)
// 3、调用导出函数
el.exportExcel()
```

#### 按列表表单导出

```
// 1、新建一个 实例
let el = new ExcelUtils('采购-拣货单')
// 2、调用实例函数，向实例添加表和数据
let data = []
let metaData = [] // 你的表单定义，对于接口主要读取的是 label 和 prop 内容如 [{"label":"","prop":""}]，具体结构可以看metaData标题内容
el.addSheet('采购-拣货单', metaData, data) // 若 data为空，则为导出模板
// 3、调用导出函数
el.exportExcel()
```

#### 表格文件上传并读取

```
//react，使用异步函数
const upload = (event) => {
  var e = window.event || event;
  var File = e.target.files[0];
  let reader = new FileReader();
  const er = new ExcelReadUtils();

  reader.onload = async function () {
    const res = await er.dealBase64Data(reader.result);
    console.log(res, 'res');
  };
  reader.readAsDataURL(File);
};
<Input
  type="file"
  onChange={upload}
  accept=".xls,.xlsx"
></Input>

//vue，使用回调函数
const er = new ExcelReadUtils();

er.dealBase64Data(province, this.getdata, "省份区域" )

getdata = (arr, type) => {
   const { newMenuData } = this.state
   this.setState({
     newMenuData: newMenuData.map(ele => {
       const { key } = ele
       if (key === type) {
         ele.children = arr.map(item => {
           const { code, name } = item
           return ({
             ...item,
             key: code,
             title: name,
             search: code,
             icon: null,
           })
         })
       }
       return ele
     }),
   })
 }


```


#### 获取导出进度

```
let that = this
let excel = new ExcelUtils('fileName')
// 单个导出情况下可使用前端定时调用方式更新
that.timeId = setInterval(() => {
    var tmp = excel.getPercentage()
    that.percentage = tmp
    if(Number(tmp)>=100){
        clearInterval(that.timeId)
    }
}, 500)

var data = []
var metaData =[]
excel.addSheet('Sheet1', metaData , data )
excel.exportExcel().then(res => {

}).catch( e => {
    that.msgError("请重试")
}).finally( fin =>{
    clearInterval(that.timeId)
}) 
```

#### 批量导出

页面代码：

```
<template>
 <FileUploadProgress 
      v-if="dialogVisiable"
      ref="downLoadProcess"
    />
</template>
```

函数代码：

```
<script>
import FileUploadProgress from '../../components/file-upload-progress' // 模态框路径
 export default {
    name: "example",
    data() {
      return {
        percentage: 0,
        checked: [],
        dialogVisiable: false,
        timeIdMap: {}
      },
         components: {
            FileUploadProgress
        },

       methods: {
           async exportExcel( rows ){
            let that = this
            that.timeIdMap = {}
            if(rows.length > 10){
              that.msgInfo('最多只能选择10个表格进行批量下载')
              return
            }
            
            var confirm = await that.$confirm('是否确认导出已选外发加工单?', "提示", {
              confirmButtonText: "确定",
              cancelButtonText: "取消",
              type: "warning"
            }).then(res => res).catch(err => null)
            if(!confirm){
              return
            }
            
            that.dialogVisiable = true
            // 包装成模态框可读数据并初始化
            var checkedList = rows.map(item => {
              var fatoryName = item.factoryName
              var orderNo = item.uniqueId
              // 具体根据业务来定义
              var fileName = orderNo + '_' + fatoryName
              return {
                ...item,
                _fileName: fileName,
                _percentage: 0.0
              }
            })
            that.$nextTick(() => {
              that.$refs.downLoadProcess.init(checkedList)
            })
            // 执行循环操作
            for(var i in checkedList){
            
              var row = checkedList[i]
              var fileName = row._fileName
              let excel = new ExcelUtils(fileName)
            // 生成全局定时器
              var timeKey = fileName + '_' + i
              var timeId = setInterval(() => {
                var tmp = excel.getPercentage()
                that.$nextTick(() => {
                  if(that.$refs.downLoadProcess){
                    that.$refs.downLoadProcess.updatePercentage(fileName, Number(tmp).toFixed(1))
                  }
                })
              }, 1000)
              that.timeIdMap[timeKey] = timeId
              
            // 这部分业务代码，可考虑抽取到外部实现
              var [error, data] = await listManufactureOrderSku(row.uniqueId)
                                      .then(res => [null, res.data] ).catch(error => [error, null])
              var afterJson = detailJson.map(item => {
                if(item.prop==='supplier'){
                  item.label = row.factoryName
                }
                return item
              })
              excel.addSheet('Sheet1', afterJson, data)
              
              var [error1, res] = await excel.exportExcel().then(res => [null, res] ).catch(err => [err, null])
              var perc = 100.0
              if(error ||  error1){
                that.msgError(error)
                perc = -1
              }
              var tmpId = that.timeIdMap[timeKey]
              if(tmpId){
                clearInterval(tmpId)
              }
              that.$refs.downLoadProcess.updatePercentage(fileName, Number(perc).toFixed(1))
        }
      }
       }
    }
</script>
```

引入模态框

```
<template>
  <div class="allwh">
    <el-dialog
      :close-on-click-modal="false"
      :visible.sync="visible"
      width="40%"
      title="导出进度列表"
    >
    <el-table  :data="tableData"
      style="width: 100%" row-key="_fileName">
       <el-table-column label="序号"  align="center" width="80">
          <template slot-scope="scope">
          <div>{{ scope.$index + 1 }}</div>
        </template>
      </el-table-column>
      <el-table-column
        prop="_fileName"
        label="文件名"
        align="center">
      </el-table-column>

      <el-table-column
        prop="_percentage"
        label="进度"
        width="200"
        align="center">
         <template slot-scope="scope">
           <el-progress :stroke-width="15" text-inside :percentage="scope.row._percentage" :status="scope.row._status"></el-progress>
        </template>
      </el-table-column>
    </el-table>
    <span slot="footer" class="dialog-footer">
      <el-button @click="visible = false">取 消</el-button>
      <el-button type="primary" @click="visible = false">确 定</el-button>
    </span>
      
    </el-dialog>
  </div>
</template>

<script>
export default {
  components: {},
  data () {
    return {
      visible: false,
      tableData: [],
      percentage: 0
    }
  },
  created () {},
  methods: {
    init(data){
      var that = this
      that.visible = true
      that.tableData = []
      for(var i in data){
        data[i]._percentage = 0.0
         that.$set(that.tableData, i, data[i])
      }
    },
    updatePercentage(_fileName, percentage){
      var that = this
      for(var i in that.tableData){
        var item = that.tableData[i]
        if(_fileName === item._fileName){
          var tmp = item._percentage || 0
          if(Number(percentage) >= tmp){
            item._percentage = Number(percentage)
          }
          if(item._percentage === -1){
            item._status = 'exception'
          }
          if(item._percentage >= 100.0){
            item._status = 'success'
            item._percentage = 100.0
          }
          that.$set(that.tableData, i, item)
        }
      }
    }
    
  }
}
</script>
<style scoped></style>
```



### API

函数描述使用new ExcelUtils(String fileName)新建实例，后面API调用需要
通过返回的实例调用let el = new ExcelUtils("file.xlsx")addJsonToSheet(String sheetName, data)将返回列表数据直接转化el.addJsonToSheet("Sheet1",[])addSheet(String sheetName, Array metaData, Array data)根据前端表格组成返回列表el.addSheet("Sheet1",metaData, [])getPercentage()获取当前导出进度el.getPercentage()addAoaToSheet(String sheetName, Array headers)根据传入表头字符串数据生成表单el.addAoaToSheet('Sheet1',[])exportExcel()真正执行导出el.exportExcel()

### MetaData

定义：描述数据的数据，比如描述常见分页列表的组成

```
[
    {
          label: '前幅图片', // 表头命名
          prop: 'frontPicUrl', // 数据字段,必传,传入唯一值
          type: 'image', // 当类型为image时会自动导出为图片
          formatter: v => { // 当存在formatter时导出会执行该formatter函数
            return v.frontPicUrl ? v.frontPicUrl + '?x-oss-process=image/resize,w_70,m_mfit/quality,Q_90' : ''
          },
          children: [], // 支持多级表头，可相应合并表头
          excel: {
            // 设置了sort后必须所有字段都要设置，sort为excel自定义导出顺序
            sort: 6,
            // 单元格样式可根据 exceljs 相应扩展
            width: 20,
            height: 80
          }
        },
    ......
]
```

取消链接：yarn unlink<br/>
cancel link: yarn unlink 

发布：npm publish<br/>
publish: npm publish

迭代： npm version [patch,minor,major]，然后 npm publish<br/>
patch： 修复bug、微小改动，改变版本号第三位<br/>
minor： 上线新功能，并对当前版本已有功能模块不影响，改变版本号第二位<br/>
major： 上线多个新功能模块，并对当前版本已有功能有影响，改变版本号第一位<br/>
iteration: npm version [patch,minor,major], then npm publish<br/>
patch: fix bugs, make little changes, and change the third digit of the version number. <br/>
major: new functions will be launched, and the existing function modules of the current version will not be affected. The second digit of the version number will be changed.<br/>
major: several new function modules will be launched, which will affect the existing functions of current version. The first digit of the version number will be changed.


安装：npm i excel-utils-bt<br/>
install: npm i excel-utils-bt