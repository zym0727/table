<template>
  <div class="my">
    <table id="outTable">
      <tr style="display:none">
        <th colspan="10">标题</th>
      </tr>

      <tr>
        <th>
          <div class="line">
            <span style="float:left">时间</span>
            <span style="float:right">项目</span>
          </div>
        </th>
        <th>地点</th>
        <th>人物</th>
        <th>事件</th>
        <th colspan="6">详情</th>
        <!-- <th>联系人</th>
        <th>记录人</th>
        <th>上级批示</th>
        <th>备注</th>
        <th>是否删除</th> -->
      </tr>

      <tbody>
        <tr>
          <td>wd</td>
          <td>ds</td>
          <td>fff</td>
          <td>dfd</td>
          <td>gf</td>
          <td>ee</td>
          <td>sd</td>
          <td>ggg</td>
          <td>sd</td>
          <td>fsf</td>
        </tr>

        <tr>
          <td>gd</td>
          <td>dfg</td>
          <td>hd</td>
          <td>hf</td>
          <td>se</td>
          <td>gf</td>
          <td>rt</td>
          <td>hj</td>
          <td>fdr</td>
          <td>ty</td>
        </tr>

        <tr>
          <td>gf</td>
          <td>kh</td>
          <td>hfhg</td>
          <td>kft</td>
          <td>tr</td>
          <td>we</td>
          <td>sf</td>
          <td>kh</td>
          <td>dr</td>
          <td>re</td>
        </tr>

        <tr>
          <td>jy</td>
          <td>dfg</td>
          <td>rew</td>
          <td>gdf</td>
          <td>gre</td>
          <td>gdr</td>
          <td>fdee</td>
          <td>gre</td>
          <td>gdr</td>
          <td>rw</td>
        </tr>
      </tbody>
    </table>
    <div>
      <button class="me" @click="exp1">导出</button>
    </div>
  </div>
</template>

<script>
import FileSaver from 'file-saver'
import XLSX from 'xlsx'
import XLSXStyle from 'xlsx-style'
export default {
  methods: {
    exp () {
      var xlsxParam = { raw: true }// 转换成excel时，使用原始的格式
      var wb = XLSX.utils.table_to_book(document.querySelector('#outTable'), xlsxParam)// outTable为列表id
      console.log('TCL: exp -> wb', wb)
      // var sheetName = wb.SheetNames
      // wb.Sheets[sheetName]['!merges'] = [{ s: { c: 0, r: 0 }, e: { c: 4, r: 0 } }]
      var wbout = XLSX.write(wb, {
        bookType: 'xlsx',
        bookSST: true,
        type: 'array'
      })
      try {
        FileSaver.saveAs(
          new Blob([wbout], { type: 'application/octet-stream;charset=utf-8' }),
          'aaaa.xlsx'
        )
      } catch (e) {
        if (typeof console !== 'undefined') console.log(e, wbout)
      }
      return wbout
    },
    exp1 () {
      var xlsxParam = { raw: true }// 转换成excel时，使用原始的格式
      var wb = XLSX.utils.table_to_book(document.querySelector('#outTable'), xlsxParam)// outTable为列表id
      var sheetName = wb.SheetNames
      var wsObj = wb.Sheets[sheetName]
      this.myFun(wsObj)
      this.addRangeBorder(wsObj['!merges'], wsObj)
      console.log('TCL: wsObj', wsObj)
      var wbout = XLSXStyle.write(wb, {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
      })
      FileSaver.saveAs(new Blob([this.s2ab(wbout)], { type: 'application/octet-stream;charset=utf-8' }), 'aabbcc.xlsx')
    },
    s2ab (s) {
      var buf = new ArrayBuffer(s.length)
      var view = new Uint8Array(buf)
      for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF
      return buf
    },
    myFun (ws) { // 自定义样式
      const sheetCols = []
      for (let i = 0; i < 10; i++) {
        sheetCols.push({ wch: 20 })
      }
      ws['!cols'] = sheetCols // 列宽
      ws['!rows'] = [{ hpx: 30 }, { hpx: 30 }] // 行高
      for (let item in ws) {
        switch (item) {
          case '!merges':
            break
          case '!ref':
            break
          case '!cols':
            break
          case '!rows':
            break
          case 'A1':
            ws['A1'].s = {
              font: {
                sz: 20
                // bold: true,
                // color: {
                //   rgb: 'FFFFAA00'
                // }
              },
              alignment: {
                horizontal: 'center',
                vertical: 'center'
              }
            }
            break
          default:
            ws[item].s = {
              font: {
                sz: 14
                // bold: true
              },
              border: {
                top: {
                  style: 'thin'
                },
                bottom: {
                  style: 'thin'
                },
                left: {
                  style: 'thin'
                },
                right: {
                  style: 'thin'
                }
              },
              alignment: {
                // 自动换行
                wrapText: 1,
                horizontal: 'center',
                vertical: 'center'
              }
            }
        }
      }
      ws['A2'].v = '日期                  项目'
      ws['A2'].s.border.diagonal = { style: 'thin' }
      ws['A2'].s.border.diagonalDown = true
    },
    /**
     *
     * @param range 合并单元格配置对象
     * @param ws  sheet配置对象
     * @returns {*}
     */
    addRangeBorder (range, ws) {
      let arr = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

      range.forEach(item => {
        let startRowNumber = Number(item.s.c)
        let endRowNumber = Number(item.e.c)

        for (let i = startRowNumber + 1; i <= endRowNumber; i++) {
          ws[arr[i] + (Number(item.e.r) + 1)] = { s: { border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } } }
        }
      })
      return ws
    }
  }
}
</script>

<style scoped>
.my {
  width: 1000px;
  height: 800px;
  background: wheat;
  margin: 0 auto;
  overflow: hidden;
}

table {
  border-collapse: collapse;
  width: 80%;
  margin: 40px auto;
}

th,
td {
  border: 1px solid red;
}

th {
  height: 50px;
}

td {
  height: 40px;
}

.me {
  margin: 10px;
  width: 100px;
  height: 50px;
  outline: none;
  background: rgba(23, 4, 2, 0.4);
  border: 1px solid green;
  border-radius: 50%;
  cursor: pointer;
}
.line {
  position: relative;
}
.line::after {
    content: '';
    width: 122px;
    height: 1px;
    display: inline-block;
    transform: rotate(25deg);
    position: absolute;
    background: black;
    left: -6px;
    top: 11px;
}
</style>
