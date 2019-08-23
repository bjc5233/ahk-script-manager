# ahk-script-manager
> AHK脚本管理工具

### 说明
1. 启动时, 执行管理中的所有ahk脚本
2. 启动后, ahk脚本**不会展示托盘图标**, 使系统托盘更简洁


### 演示
<div align=center><img src="https://github.com/bjc5233/ahk-script-manager/raw/master/resources/demo.png"/></div>
<div align=center><img src="https://github.com/bjc5233/ahk-script-manager/raw/master/resources/demo2.png"/></div>



### 注意
1. 如果脚本需要在退出时执行操作, 请加入代码
    ```
    OnExit("ExitFunc")
    ExitFunc() {
        MsgBox, 清理资源
    }
    ```
2. ok

### TODO
1. 存在bug：启动后, 脚本托盘图标会显示，重新打开管理界面后托盘图标才消失
2. 新增功能：默认全部隐藏托盘图标，但增加配置，有的任务常需要进行托盘图标右键操作；增加contextmenu-显示图标\隐藏图标
3. 使用sqlite管理脚本信息