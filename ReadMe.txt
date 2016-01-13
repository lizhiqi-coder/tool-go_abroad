作者：cafe lee 
日期：2016.1.13
本脚本使用范围：
	Android 中文版转国际版
	主要操作strings.xml 文件
	使用脚本生成xlsx 文件，将信息映射至 excel 文件中
	翻译在excel 根据相应中文翻译成英文，并写入对应的key 的列中
	生成翻译过后的string_en.xml 文件，即目标文件
	
使用方法：
	1,将strings.xml 文件copy 至本项目根目录中

	2，在根目录打开终端，运行 python run_to_generate_PM_file.py,生成strings.xlsx 文件，
	   将文件交予PM 翻译，并严格按照文件格式进行操作
		文件头分别为：android_key	CN_String_value 	EN_String_value

		2.1, 如果不是第一次使用，之前生成并翻译过string.xlsx，这时需将旧版文件翻译的内容
			copy 至新生成的strings.xlsx 文件的
			（注意：每次生成strings.xlsx 文件，其中没有任何被翻译的字段）
			更行操作：
				1,将旧版xlsx 文件更名为strings_old.xlsx
				2,在终端运行 python updata_PM_file.py

	3,翻译完后，将翻译后的文件copy 至根目录中，文件名问strings.xlsx 
		在终端运行 python run_to_generate_ANDROID_string_file.py
		程序结束后，在根目录中生成 string_en.xml文件，即为目的文件



===========================================================================================================================
注意事项：	1，程序运行时必须关闭xlsx 文件
		2，本脚本不能在xml 中解析 string-array 字段，考虑到该字段应用很少，不想大费周章，请开发者自行备份粘贴
		3，上述使用方法步骤顺序不可随意改变，会有意想不到的结果
		

		
									 