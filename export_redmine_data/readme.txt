环境配置：
安装python3.5
开始-运行-输入cmd 
依次输入
pip install xlrd
pip install xlwt
pip install python-redmine



参数设置：
1.用编辑器打开config.py  （右键该文件，选择edit with IDLE）
2.设置起始日期STARTDAT、截止日期ENDDATE 、月份MONTH,登陆名MYKEY


3.（可选）修改输出excel名
MANAGE_LIST: 运行Pro_Manage_Out.py所生成的excel，由系统直接导出的项目-项目经理对应名单，可能含有错误，校对后直接保存即可。
# 项目维度-按项目经理
FILENAME_MANAGE_OUT： 生成的结果excel名
# 项目维度-按部门
FILENAME_DEPARTMENT ： 生成的结果excel名


当然也可以选择都不设置就改日期参数



操作流程：
0.参数设置

1.开始-运行-输入cmd

2.输入 python 空格 然后将ProManage_Out.py 拖入 然后回车。如下：


3.等待运行完毕，在该文件目录下将会生成经理名单的excel文件，打开校对项目经理名单后直接保存（保存后不需要再次运行），
  建议修改后保存+另存，以便新项目添加时重新运行更方便校对，项目顺序无关，可以删除不想查看的项目。

4.在cmd中输入 python 空格 然后将run.py 拖入 然后回车。可得到项目维度-按部门，与项目维度-按经理两个结果
  若想分开运行，将dpark_department.py拖入 可得到 项目维度-按部门 的结果
               将dpark_promanage.py拖入 可得到 项目维度-按经理 的结果






注：将问题状态‘已分配’视作‘新问题’
注2：报错联系QQ：445839891
注2：还有啥问题需要再联系