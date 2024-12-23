<template>
  <div class="body">
    <div class="chat-container" id = 'chat-container'>
      <div v-for="(item,index) in messageList" style="display: flex;" :key="index">
        <div v-if="item.type === 'user'" class="message user-message">
          <div v-if="item.message.type === 'text'">
            <span>
              {{ item.message.data }}
            </span>
          </div>
        </div>
        <div v-loading="loading[index]" v-if="item.type === 'reply'" class="message reply-message">
          <div v-if="item.message.type === 'text'">
            <div >
              {{ item.message.data }}
            </div>
            <el-button v-if="item.message.button" type="primary" @click="addColumn">
              插入公式
            </el-button>
          </div>
          <div v-loading="loading[index]" v-if="item.message.type === 'table'">
            <el-table :data="item.message.data.columns" border style="width: 100%">
              <el-table-column v-for="(o,index)  in item.message.data.header" :prop="o" :key="index" :label="o">
              </el-table-column>
            </el-table>
            <el-button v-if="showApplyBtn" type="primary" @click="sortTable">
              应用操作
            </el-button>
          </div>
          <div v-loading="loading[index]" style="z-index: 999;" class="chart" v-if="item.message.type === 'chart'">
            <div v-html="item.message.data"></div>
            <el-button type="primary" @click="createChart">
              插入图片
            </el-button>
          </div>
        </div>
      </div>
    </div>
    <div class="input-container">
      <el-input style="margin-right: 15px;" v-model="message" placeholder="输入消息..." @keyup.enter="sendMessage" id="message-input"></el-input>
      <el-button id="send-btn" type="primary" @click="sendMessage">发送</el-button>
    </div>
  </div>
</template>


<script setup>
import {nextTick, ref} from "vue";
import * as echarts from 'echarts';

const loading = ref([]);
const showApplyBtn = ref(false);
const message = ref('');
const chart = ref(null);
const messageList = ref([]);

const scrollToBottom = () =>{
  const chatContainer = document.getElementById('chat-container');
  chatContainer.scrollTop = chatContainer.scrollHeight;
}

const sendMessage = () => {
  if (message.value !== '') {
    messageList.value.push({ message: {data:message.value,type:'text'}, type: 'user' });
    loading.value.push(false)

    // 这里对三种情况进行判断
    // 1. Sort the table by sales in descending order
    if (message.value === 'Sort the table by sales in descending order') {
      vueSortTable();
      loading.value.push(true)
    }else if (message.value === 'Create a scatter plot of sales and costs') {
      vueCreateChart();
      loading.value.push(true)
    }else if (message.value === 'Insert a column of profits') {
      vueAddColum();
      loading.value.push(true)
    }else{
      messageList.value.push({ message: {data:"目前暂不支持，请重新输入",type:'text'}, type: 'reply' });
      loading.value.push(false)
    }
    //   4. 延时
    let end = loading.value.length
    setTimeout(() => {
      loading.value[end - 1] = false
      showApplyBtn.value = true
    }, 3000);
    message.value = '';
    scrollToBottom();
  }
}

// 1. Sort the table by sales in descending order
const vueSortTable = async () => {
  //   1. 得到数据
  let header = []
  let columns ;
  let dataRange;
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let expensesTable = sheet.tables.getItemAt(0);
    dataRange = expensesTable.getDataBodyRange().load("values");
    // // Get data from the header row.
    let headerRange = expensesTable.getHeaderRowRange().load("values");
    await context.sync();
    header = headerRange.values[0]
  })
  //  2. 排序数据
  dataRange.values.sort((a,b)=> b[4] - a[4])
  columns = initTable(header,dataRange.values)


//   3. 添加到消息列表
  messageList.value.push({ message: {type:'table',data:{header,columns}}, type: 'reply' });

}


// 得到表格数据
const initTable = (header,data) =>{
  let columns = []
  data.forEach(item => {
    let column = {}
    let i = 0
    header.forEach(o =>{
      Object.defineProperty(column,o,{ value: item[i], writable: true, enumerable: true, configurable: true })
      i += 1
    })
    columns.push(column)
  });
  return columns
}

// 对excel表格执行排序
const sortTable = async () => {
  return await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let expensesTable = sheet.tables.getItemAt(0);

    let sortRange = expensesTable.getDataBodyRange().load("values");

    // // Get data from the header row.
    // let headerRange = expensesTable.getHeaderRowRange().load("values");
    // await context.sync();
    // console.log('sortRange.values------------' + sortRange.values)
    // console.log('sortRange val:', sortRange.values[0][4], ' ----', sortRange.values[3][4])
    // console.log('headerRange.values------------' + headerRange.values)
    // console.log('headerRange val:', headerRange.values[0][0], ' ----', headerRange.values[0][4])

    sortRange.sort.apply([
      {
        key: 4,
        ascending: false,
      },
    ]);
    await context.sync();
  });
}

// 2. Create a scatter plot of sales and costs
const vueCreateChart = async() =>{
  let data = [];
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    data = sheet.getRange("F7:G31").load("values");
    await context.sync();
  });
  
  // 使用 nextTick 确保 DOM 更新完成
  // 创建一个唯一的 DOM ID
  const chartId = `chart-${Date.now()}`;
  const domString = `<div id="${chartId}" class="chart" style="width: 100%; height: 250px;"></div>`

  messageList.value.push({ message: {type:'chart',data:domString}, type: 'reply' });
  
  await nextTick();
  initChart(data.values,chartId)
}

// 生成 echarts 图表
const initChart = async (data,domId) =>{
  const dom = document.getElementById(domId);
  let mychart = echarts.init(dom);
  // console.log(dom);
  // console.log(mychart);

  let option;
  option = {
    xAxis: {},
    yAxis: {},
    series: [
      {
        symbolSize: 5,
        data:data.slice(1),
        type: 'scatter'
      }
    ]
  }
  option && mychart.setOption(option)
  // console.log(mychart);

  await nextTick();
  mychart.resize()
}

// 执行 excel 插入图表
const createChart = async () => {
  return await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    // let expensesTable = sheet.tables.getItemAt(0);

    let dataRange = sheet.getRange("F7:G31")
    let chart = sheet.charts.add("XYScatter", dataRange, "auto");
    chart.setPosition("L7", "S21");
    chart.title.text = "'Costs' by 'Sales'";
    // chart.legend.position = "left";
    // chart.legend.format.fill.setSolidColor("white");
    chart.axes.valueAxis.title.text = "Costs";
    chart.axes.categoryAxis.title.text = "Sales";

    await context.sync();
  });
}

// 3. Insert a column of profits
const vueAddColum = () =>{
  messageList.value.push({ message: {type:'text',data:'=TEXT(INT([@Sales])-INT([@Costs]), "0")',button:true}, type: 'reply' });
}
const addColumn = async () => {
  return await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let expensesTable = sheet.tables.getItemAt(0);

    let profits = [["Profits"]];
    for (let i = 0; i < 24; i++) {
      profits.push(['=TEXT(INT([@Sales])-INT([@Costs]), "0")']);
      // profits.push(['=[@Sales]-[@Costs]']);
    }
    // 向表格中添加新列
    expensesTable.columns.add(null /*add columns to the end of the table*/, profits);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
  });
}

</script>

<style scoped>

.body {
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

/* 页面容器 */
.chat-container {
  flex: 1; /* 允许容器扩展以填充可用空间 */
  flex-direction: column;
  overflow-y: auto; /* 允许滚动 */
  padding: 10px;
  border: 1px solid #ddd;
  border-radius: 12px;
  background-color: #fff;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
  margin: 5px 0px 43px 0px;
  box-sizing: border-box; /* 确保 padding 不影响最终大小 */
}

/* 用户发送的消息 */
.user-message {
  align-items: flex-start; /* 内容从左到右排列 */
  right: 0%;
  display: inline-block; 
  margin-bottom: 15px;
  padding: 1rem;
  background-color: #007bff;
  color: white;
  text-align: right;
  margin-left: auto;
  border-radius: 15px 15px 15px 15px; /* 圆角边 */
  white-space: normal; /* 确保文本可以换行 */
  
}

/* 自动回复消息 */
.reply-message {
  display: inline-block; /* 或者使用 inline */
  margin-bottom: 15px;
  padding: 1rem;
  background-color: #e1f5fe;
  color: #007bff;
  text-align: left;
  border-radius: 15px 15px 15px 15px; /* 圆角边 */
}

/* 输入框容器 */
.input-container {
  display: flex;
  position: fixed;
  bottom: 0;
  left: 0;
  right: 0;
  padding: 10px 20px;
  background-color: #fff;
  border-top: 1px solid #ddd;
  box-shadow: 0 -2px 8px rgba(0, 0, 0, 0.1);
  z-index: 10;
}

/* 输入框样式 */
#message-input {
  flex: 1;
  padding: 10px;
  border: 1px solid #ccc;
  border-radius: 10px;
  margin-right: 10px;
  font-size: 16px;
  box-sizing: border-box; /* 确保 padding 不影响最终大小 */
}

/* 发送按钮样式 */
#send-btn {
  padding: 10px 15px;
  background-color: #007bff;
  color: white;
  border: none;
  border-radius: 10px;
  cursor: pointer;
  font-size: 16px;
}

#send-btn:hover {
  background-color: #0056b3;
}
.chart{
  width: 500px;
}
</style>