Excel 动态看板（本地版）

如果你的终端环境无法运行 Python/PowerShell，本方案用浏览器直接解析 Excel：
1) 打开 dashboard/index.html
2) 选择你的 Excel 文件（例如：数据源cursor.xlsx）
3) 选择工作表/字段，点击“应用筛选”即可生成图表

说明：
- 需要联网以加载 xlsx.js 与 echarts 的 CDN 资源
- 仅在本机浏览器运行，不会上传你的数据

