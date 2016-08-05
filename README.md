# KMS_Active KMS激活工具

KMS_Active是用于激活 Windows 系统的一个自动化程序

# 使用方法

1. 首先需要拥有一个 KMS 激活服务器
2. 命令范例
    ```
    激活 Windows (请确认安装的是VL版本)
    KMS_Active.exe /w /s active.server.com 

    显示 Windows 激活信息
    KMS_Active.exe /d /s active.server.com

    激活 Office (请确认安装的是VL版本)
    KMS_Active.exe /d /s active.server.com

    同时激活 Windows 和 Office
    KMS_Active.exe /w /o /s active.server.com
    ```

# 参数列表
*   /s  设置激活服务器地址（必选）
*   /d  显示 Windows 激活信息
*   /o  激活 Windows 
*   /w  激活 Office

# 参考连接
    Microsoft MSDN GetProductInfo function
    [GetProductInfo function](https://msdn.microsoft.com/zh-cn/library/windows/desktop/ms724358(v=vs.85).aspx)

    Microsoft KMS Client Setup Keys
    [Microsoft KMS Client Setup Keys](https://technet.microsoft.com/en-us/library/jj612867.aspx)