# MypythonProject

1.使用方法
	通过CMD或Powershell导航到本文件夹
	执行bat脚本添加openpyxl模块，也可通过pip install openpyxl命令添加
	本方法支持单文件或文件夹处理
	切记处理前备份待处理文件
	
2.命令	
	-I, --items 	Excel文件路径或文件夹
	-C, --column 	删除第C列
	-R, --row		删除第R行
	
3.注意事项
	暂不支持.xls,请另存为.xlsx
	
python deleteExcel.py --I test.xlsx -C 1 -R 2
