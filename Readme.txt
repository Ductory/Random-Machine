随手乱写的程序，如果有bug，请反馈至邮箱: Dangfer@qq.com

操作帮助(第一次使用时，请看这里)：
·关于列表文件（请确保是ANSI格式）：
	对于要读取的值必须按照以下格式书写：
		预抽取头
		项1
		...
		项n
	预抽取头：
	预抽取头用于加快文件读取速度并避免不必要的错误，包括以下可选参数:
	·K: 关键字(String)
	·W: 权重(Long)
	项：
	每个项必须按照以下格式书写：
		值[, 关键字][, 权重]
	项的每个参数相对位置不可改变，参数之间必须有且只有一个空格。
·关于CheckList：
	CheckList是一组CheckBox的集合
	单击左键，可以进行类似CheckBox的操作
	单击右键，可以选择“全选”和“取消全选”
	在本程序中，“列表预览”和“筛选器预览”使用了CheckList
·关于“列表预览”：
	列表预览展现当前待抽取项(即勾选项)
·关于设置：
	·允许加权：在此选项勾选时，按权重进行抽取，否则相等概率抽取
	·自动屏蔽：在此选项勾选时，将从待抽取项中去除当前抽取项
	·幂权：在此选项勾选时，将以“按幂加权”的方式进行加权抽取
	·批量抽取数量：一次名单批量抽取和随机数批量抽取的数量
	·组名：在“抽取一览”列表中显示的组的名字。若组名为空，则默认组名为“未命名组”
·关于扩展：
	RM提供了一定的外接程序拓展功能，用户可以使用扩展程序来扩展RM功能
	Extend 扩展程序使用方法：
		·在Config.ini文件的[Explorer]下新建En项，其中n必须是连续的
		·格式: En=[(PARAMS)]<EXE_TAG>:<COMMAND>
			PARAMS: 可选的，有以下参数：
				E: 在运行扩展程序时RM侦听返回值。当扩展程序终止时以弹窗形式给出退出代码。
				S: 在RM运行时自动启动
			EXE_TAG: 必需的，用于指定在“扩展”菜单中显示的内容
			COMMAND: 必需的，用于指定扩展程序的命令行
		·RM提供了以下几个环境参数：
			HWND: 给出RM主窗口的句柄
		可在COMMAND中使用环境参数。在使用环境参数时，必须用%%引起。
	例如：
		E1=(SA)扩展1:Ex1 %HWND%
		就设置了一个扩展1，并将RM窗口句柄作为参数传给了Ex目录下的Ex1.exe
·关于筛选器：
	对当前待抽取列表中选中的项进行筛选，提供三种方式的筛选
	·值
		对输入的 string 进行模式匹配
		?		任何单一字符
		*		零个或多个字符
		#		任何一个数字 (0–9)
		[charlist]	charlist 中的任何单一字符
		[!charlist]	不在 charlist 中的任何单一字符
		在中括号 ([ ]) 中，可以用由一个或多个字符 (charlist) 组成的 组 与 string 中的任一字符进行匹配，这个组几乎包括任何一个字符代码（用来表示一字符集中特定字符的数字，比如ANSI 字符集）以及数字。
		注意 为了与左括号 ([)、问号 (?)、数字符号 (#) 和星号 (*) 等特殊字符进行匹配，可以将它们用方括号括起来。惊叹号(!)不需要方括号括起来，在组外它匹配自身。不能在一个组内使用右括号 (]) 与自身匹配，但在组外可以作为个别字符使用。
		通过在范围的上、下限之间用连字符 (–)，charlist 可以指定字符的范围。例如， [A-Z] 意为匹配 A–Z 之间的任意字母（不区分大小写）。连字符可以出现在 charlist 的开头（如果使用惊叹号，则在惊叹号之后），也可以出现在 charlist 的结尾与自身匹配。在任何其它地方，连字符用来识别字符的范围。
	·关键字
		提供两种逻辑操作，“存在”和“包含”
		·存在：只要项的关键字中存在某一筛选器要求关键字就选中
		·包含：只有筛选器要求关键字在项的关键字全部存在才选中
	·权
		选择指定范围内的权的项
		例如，在文本框中输入">=1"即选择勾选的项中所有权重大于等于1的项
·关于随机数批量抽取
	升序实现用的是计数排序，唯一实现用的是哈希查找，因此请不要指定过大的范围。如要指定很大的范围，请使用扩展。
·关于RM脚本
	RM脚本允许批量地抽取名单并导出结果。
	脚本实质上是VBS，因此不支持中文变量、函数、过程名，但仍支持中文字符串。RM内置了一些脚本指令，脚本指令如下：
	·extract(次数, 是否加权, 组名)：抽取指令，将新建一个组，并将抽取结果加入组中
	·export(导出路径)：导出指令，将导出抽取结果
	·reset()：重置抽取列表
	注意 extract指令默认以“自动屏蔽”的方式抽取，因此在抽取列表空时，请自行调用reset指令
	此外，抽取结果保存在内置脚本参数RMBuff中。如果要调用内置的指令或参数，请使用 RM.指令或参数 的格式来调用
	例如：
		RM.extract 5, True, "组1"
		RM.export "C:\test.txt"
历史版本:
Ver 1.0.0-RELEASE
·添加了“自动保存修改后的列表文件”功能
·添加了“抽取结果导出”功能
·添加了RM脚本（RM Script Ver 1.0.0-BETA）
·更新了筛选器，优化了筛选算法
·优化了抽取算法
·优化了文件读取
·优化了历史记录（最近打开的历史记录会更新至第一条）
·优化了用户界面（“外接程序”现在改为“扩展”菜单）
·优化了“反转权”（现在改为“幂权”）
·移除了工具区和平滑移动算法
·移除了列表文件的注释功能和项标签
·取消了“批量抽取”的延时，移除了“闪抽”
·修复了“外接程序”写入错误的bug
·修复了“加权抽取”存在的bug
·修复了打开历史记录时可能出错的bug
·部分界面的改动

Ver 0.6.0
·支持了UNICODE(UTF-16 LE，为兼容VB6)，但程序可能更不稳定
·优化了“关键字匹配”
·优化了“随机数抽取”
·优化了抽取算法，添加二分加权抽取算法，移除了CEA算法
·修复了“加权和初始化错误”的bug
·修复了“随机数抽取”中“唯一”勾选时无效的bug
·修复了“随机数抽取”中“唯一”未勾选时无法抽取的bug
·修复了“抽取”中“允许重复”未勾选时列表检查缺失的bug
·修复了“加权抽取”中会抽到权重为0的项的bug
·取消了对"MSCOMCTL.OCX"和"COMDLG32.OCX"的部件引用
·改变了控件的样式(XP样式)
·部分界面的改动

Ver 0.5.0
·添加了“反转权”
·更新了筛选器
·更新了外接程序
·优化了用户界面

Ver 0.4.0
·添加了CEA算法(集合抽取算法)，移除了TEA算法
·添加了工具区及动画效果
·列表文件要求“预读取头”(详见list\sample.txt)
·添加了“外接程序”菜单(详见config.ini)
·优化了悬浮窗
·优化了抽取界面
·优化了部分代码
·修复了“自动屏蔽”未勾选时“允许重复”无效的bug
·部分界面的改动

Ver 0.3.0
·添加了TEA算法(表抽取算法)，移除了VTEA算法和LAEA2算法
·优化了悬浮窗，增加了悬浮窗“震动抽取”功能(上下震动，产生一个“随机数批量抽取”范围中的随机整数)
·支持“多关键字”功能
·修复了“加权抽取”错误的bug

Ver 0.2.0
·增加了“关键字”及相关功能(部分功能)
·增加了“悬浮窗”功能
·增加了“历史记录”功能
·修复了“允许加权”未勾选时仍然有效的bug(VTEA)
·扩大了列表文件的选择范围
·改动局部界面

Ver 0.1.0
·增加了“随机数批量抽取”功能
·修复了“允许重复”勾选时“自动屏蔽已抽”仍然有效的bug

Ver 0.0.0
·最初版本