# hoteru.AnExcelAdd-In
> 最终的加载宏文件适用于软件版本：`64位 Win7`上的`32位`的`Excel2007`，`Excel2010`，`Excel2013`

这是一个作者在某NGO的热线中心实习时设计的`Microsoft Excel2007加载宏`，当时的目的主要是为了自动化制作月报和阶段性报告的过程，减少不必要的重复性工作。
除了`README.md`，本仓库包含下列文件：
+ `customUI.xml`
+ `阶段性报告制作v1.bas`
+ `阶段性报告制作Shadow.bas`
+ `其他小功能.bas`
+ `月报自动制作脚本v1.bas`
+ `Inno.xlam`

其中的`Inno.xlam`文件即为可使用的加载宏文件。

由上述其余的5个文件即可组建出可用的`Inno.xlam`，该加载宏会为`Excel`增加新的`Ribbon`选项卡`Inno HC`，`customUI.xml`文件负责指定增加的`Inno HC`选项卡的样式。
为了加载宏的组装，可能需要编辑`Excel`文件中`xml`文件的工具：[CustomUIEditor](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx)。可以先新建一个空的`Excel`加载宏文件`Inno.xlam`（推荐`.xlam`后缀的加载宏），然后使用`CustomUIEditor`对其中的`xml`文件进行编辑，
即将本仓库中的`customUI.xml`中的内容复制到`CustomUIEditor`中编辑的`customUI.xml`文件中。

完成`xml`文件的编辑后，请在`CustomUIEditor`中保存该加载宏文件，然后关闭该软件。

在`CustomUIEditor`中编辑加载宏文件的情形可能如下：

![在CustomUiEditor中编辑customUI.xml的情景](http://ww3.sinaimg.cn/large/005BEzjejw1f851tt8y03j30sm0ihgvg.jpg)

现在可以双击打开该加载宏（如果被问到是否启用宏，请记住选择启用），打开`Visual Basic Editor`，依次导入`阶段性报告制作v1.bas`，`阶段性报告制作Shadow.bas`，`月报自动制作脚本v1.bas`。
现在我们已经完成了加载宏的制作过程。

安装加载宏的过程非常简单，可以考虑在`Excel`中使用`Alt+T+I`快捷键调出加载宏对话框，加载制作完成的加载宏文件。
详细过程还可以参考这篇文章：[What is an XLAM File?](http://pcsupport.about.com/od/fileextensions/f/xlamfile.htm)，这篇文章还包含有对`.xlam`文件的介绍。

最终在`Excel2007`上的效果可能是这样：

![加载宏效果](http://ww4.sinaimg.cn/large/005BEzjejw1f851ue330cj30qk0dogpn.jpg)
