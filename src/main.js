import { createApp } from 'vue'
import App from './App.vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'  // 引入 Element Plus 样式


window.Office.onReady(() => {
    createApp(App).use(ElementPlus).mount('#app');
});