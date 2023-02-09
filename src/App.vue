<template>
  <main>
    <div class="content">
      <div class="drops"
           @drop="handleDrop"
           @dragover="handleDragover"
           @dragenter="handleDragover"
           @click="clickInputFile">
        <input class="file-input" type="file" @change.prevent.stop="handleInput"/>
        <img class="xls-image" src="./assets/excel.svg">
        <div class="upload-text">{{ fileMessage }}</div>
      </div>
      <div v-if="errorMessage" class="error">{{ errorMessage }}</div>
    </div>
  </main>
</template>

<script>
import * as XLSX from 'xlsx'
import _ from 'lodash'

export default {
  data() {
    return {
      fileMessage: '将i18n Excel文件拖到此处 或点击上传',
      errorMessage: ''
    }
  },
  methods: {
    handleDrop(e) {
      e.stopPropagation()
      e.preventDefault()
      let file = e.dataTransfer.files[0];
      console.log("handleDrop", file)
      this.fileFilter(file)
    },

    handleInput(e) {
      let file = e.target.files[0]
      e.target.value = ""
      console.log(file)
      this.fileFilter(file)
    },

    fileFilter(file) {
      let name = file.name.toLowerCase()
      if (name.endsWith('xls') || name.endsWith('xlsx')) {
        this.fileMessage = name
        this.errorMessage = ''
        this.readXls(file)
      } else {
        alert('文件类型错误')
      }
    },

    clickInputFile() {
      document.querySelector('.file-input').click()
    },

    handleDragover(e) {
      e.stopPropagation()
      e.preventDefault()
      e.dataTransfer.dropEffect = 'copy'
    },

    readXls(file) {
      const fileReader = new FileReader()
      let that = this
      fileReader.onload = (event) => {
        console.log('ev', event)
        try {
          const data = event.target.result
          let workbook = XLSX.read(data, {
            type: 'binary'
          }) // 以二进制流方式读取得到整份excel表格对象
          let sheetJson = [] // 存储获取到的数据
          // 表格的表格范围，可用于判断表头是否数量是否正确
          // 遍历每张表读取
          for (const sheet in workbook.Sheets) {
            if (workbook.Sheets.hasOwnProperty(sheet)) {
              console.log('Sheets', workbook.Sheets)
              // let fromTo = workbook.Sheets[sheet]['!ref']
              // console.log('fromTo', fromTo)
              sheetJson = sheetJson.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]))
              break // 如果只取第一张表，就取消注释这行
            }
          }
          console.log(sheetJson)
          that.generateOriginData(sheetJson)
        } catch (e) {
          console.log('load error=>', e)
        }
      }
      // 以二进制方式打开文件
      fileReader.readAsBinaryString(file)
    },

    /**
     * {
     *     "key": "mainpage_junkclean",
     *     "中文": "垃圾清理",
     *     "英文": "Junk Clean",
     *     "日语": "ジャンククリーン",
     *     "法语(fr)": "Effacer les fichiers inutiles",
     *     "德语(de)": "Junk-Bereinigung",
     *     "西班牙语(es)": "Limpieza de basura",
     *     "葡萄牙语(pt)": "Limpeza de lixo",
     *     "意大利语(it)": "Sgombero rifiuti",
     *     "韩语(ko)": "쓰레기 치우기",
     *     "繁中": "垃圾清理"
     * }
     *
     * <resources xmlns:tools="http://schemas.android.com/tools">
     * <string name="app_name">Super Cleaner</string>
     * </resources>
     */
    generateOriginData(sheetJson) {
      let dataKeys = []
      let dataRes = {}
      const sheetKey = 'key'
      let tempValStr
      let tempSheetKeyStr
      let tempSheetValStr
      sheetJson.forEach((sheet, index) => {
        if (index === 0) {
          _.forIn(sheet, (_, key) => {
            dataKeys = dataKeys.concat(key.trim())
          })
        }
        if (sheet[sheetKey]) {
          dataKeys.forEach((key) => {
            if (key === sheetKey) return
            tempSheetKeyStr = sheet[sheetKey].trim()
            if (sheet[key]) {
              tempSheetValStr = sheet[key]
                  .replaceAll('&amp；', '&amp;')
                  .replaceAll(new RegExp('&(?!amp;)', 'gm'), "&amp;")
                  .trim()
            } else {
              this.errorMessage = (index + 2) + '行[' + key + ']列找不到内容'
              throw new DOMException(this.errorMessage)
            }
            tempValStr = '<string name="' + tempSheetKeyStr + '">' + tempSheetValStr + '</string>'
            if (dataRes[key]) {//有值就添加
              dataRes[key] = dataRes[key].concat(tempValStr)
            } else {//没有重新创建对象
              dataRes[key] = '<resources xmlns:tools="http://schemas.android.com/tools">' + tempValStr
            }
          })
        } else {
          this.errorMessage = '文档缺少第一行 （一般为字符串"key // 语言种类"）'
          throw new DOMException(this.errorMessage)
        }
      })
      console.log('dataKeys', dataKeys)
      let parser = new DOMParser()
      _.forIn(dataRes, (val, key) => {
        dataRes[key] = val.concat('</resources>')
        let xml = parser.parseFromString(dataRes[key], "text/xml")
        console.log(xml)
        let blob = new Blob([dataRes[key]], {type: 'text/xml'})
        let url = URL.createObjectURL(blob)
        let w = window.open(url)
        setTimeout(() => w.document.title = key, 1500)
        URL.revokeObjectURL(url)
      })
    }
  }
}

</script>

<style scoped>

main {
  display: flex;
  flex-direction: column;
  justify-content: center;
}

.content {
  margin-right: auto;
  margin-left: auto;
}

.file-input {
  display: none;
}

.drops {
  margin-top: 30px;
  width: 800px;
  height: 150px;
  background-color: rgba(85, 98, 112, 0.2);
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
}

.xls-image {
  width: 30px;
  height: 30px;
}

.error {
  color: red;
  margin-top: 30px;
}

.upload-text {
  margin-left: 30px;
  user-select: none;
}

</style>
