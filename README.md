# RevitShortcutExpansion
An extension tool for Revit shortcut keys

- 以前在办公室总要去同事电脑排查问题，或者有时要用公用电脑进行Revit演示，但不同使用者的快捷键方案不一样

- 而Revit的快捷键管理只能通过手动导入导出的方式进行快捷键的批量替换

- 通过这个扩展工具，Revit可以保存多套快捷键方案，并可以很方便地进行切换

![image](https://user-images.githubusercontent.com/36910094/169762302-f8fd35c0-411e-4bc9-9bbc-cd4f05127921.png)

**使用方法：**

将`RevitShortcutExpansion.addin`和`RevitShortcutExpansion.dll`放到`C:\ProgramData\Autodesk\Revit\Addins\\<你的Revit版本号>\`里，

启动Revit后，使用快捷键KS可以弹出扩展后的快捷键修改窗口

**其他注意事项：**

创建的快捷键方案都在文件夹`C:\用户\<当前用户>\AppData\Roaming\Autodesk\Revit\<你的Revit版本>\Four\`里，删除对应的文件，窗口也会删掉对应的选项
