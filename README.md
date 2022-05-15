# saveToExcel
功能：导入json数据到excel
使用步骤：
+ 安装python
+ 安装模块
  + pip install xlrd -i https://pypi.tuna.tsinghua.edu.cn/simple
  + pip install xlwt -i https://pypi.tuna.tsinghua.edu.cn/simple
+ 编辑cmd.xls的sheet1
  + `json_path`对应的`value`写入json文件的路径
    + json读的是根节点第一个列表元素，以此为数据源，注意json格式
  + `target_path`对应的`value`写入excel要保存的路径
+ 运行
  + python saveToExcel.py
+ 测试用例
  + 已准备了order.json
  + 修改一下`target_path`为本地的目录
  + 执行
    + python saveToExcel.py