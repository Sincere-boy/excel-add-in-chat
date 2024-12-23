# 启动项目

### 在代码目录下载包

在 excel-copilot-code 路径中，执行 `npm install`命令，将在该目录下生成 `node_modules`包目录

### 启动 vue

在 excel-copilot-code 路径中，新建终端1，执行`npm run serve`启动 vue 服务

### 启动 excel-add-in 应用

在 excel-copilot-code 路径中，新建终端2，执行`npm start`启动 excel-add-in 应用

执行该命令后，会自动打开 excel 应用

### 使用加载项

打开期望使用 excel-add-in 加载项的 excel 文件，在文件中点击“开始”栏右侧的加载项，从中选择名为：“Show Task Pane”的加载项打开，即可使用

### 关闭加载项

 在终端 1 中使用`ctrl+C`或者`command+C`终止 vue 进程

在终端 2 中执行`npm stop`关闭 excel-add-in 应用