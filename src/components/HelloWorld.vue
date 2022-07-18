<template>
  <div>
    <button @click="handleExport">导出</button>
    <table>
       <thead>
        <tr>
          <th>序号</th>
          <th>姓名</th>
          <th>年龄</th>
          <th>性别</th>
        </tr>
       </thead>
       <tbody>
         <tr v-for="item in tableData" :key="item.id">
          <td>{{item.id}}</td>
          <td>{{item.name}}</td>
          <td>{{item.age}}</td>
          <td>{{item.gender}}</td>
         </tr>
       </tbody>
    </table>
  </div>
</template>

<script>
import docxtemplater from 'docxtemplater'
import PizZip from 'pizzip'
import JSZipUtils from 'jszip-utils'
import { saveAs } from 'file-saver'

export default {
  name: 'HelloWorld',
  data() {
    return {
      tableData: [
        {
          id: 1,
          name: "cc",
          age: "22",
          gender: "man",
        },
        {
          id: 2,
          name: "marry",
          age: "25",
          gender: "woman",
        }
      ]
    }
  },
  methods: {
    handleExport() {
      console.log(this.tableData)
      const data = {
        titleDate: "2021-07-05 00:00 至 2022-07-18 00:00",
        titleName: "测试表格",
        table: this.tableData
      }
      this.exportDocx("Word/template.doc", data, "测试表格.doc");
    },
    exportDocx(tempDocxPath, data, fileName) {
      JSZipUtils.getBinaryContent(tempDocxPath, function (error, content) {
        if (error) {
          throw error;
        }
        console.log("content==", content)
        let zip = new PizZip(content);
        let doc = new docxtemplater().loadZip(zip);
        console.log("doc===", doc)
        doc.setData(data);
        try {
          doc.render();
        } catch (error) {
          let e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
          };
          console.log({
            error: e
          });
          throw error;
        }
        let out = doc.getZip().generate({
          type: "blob",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        console.log("out==", out)
        saveAs(out, fileName);
      });
    },
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
table {
    border-collapse: collapse;
    border-spacing: 0;
    margin-top:10px;
}
table td,
table th {
    text-align: center;
    padding: 5px 10px;
    border: 1px solid #ccc;
}
</style>
