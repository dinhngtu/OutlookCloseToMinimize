Installation
------------

To use the addin, you'll first need to install the [VSTO runtime](https://www.microsoft.com/en-us/download/details.aspx?id=56961).
Afterwards, extract the zip content to an appropriate directory (not in a temporary directory, since the addin runs in place).

Next, you'll need to remove the "mark-of-the-Web" which blocks the addin's installation.
Open a PowerShell window in the downloaded file's folder (Shift-right click, then choose "Open PowerShell window here"), then run the following command:

```Unblock-File *```

Finally, simply double-click the .vsto file to install.

The addin works best when the "Hide when minimized" option is enabled (by right-clicking Outlook's tray icon):

![image](https://user-images.githubusercontent.com/1257909/134686359-b6df9c6f-364e-4c40-9d9a-ec67cb0fa3bd.png)

Addin keeps disabling itself
----------------------------

Outlook disables addins that cause it to start slowly. If you encounter this problem, apply the included `forceload.reg` Registry file.

If Outlook crashes
------------------

If Outlook crashes, simply run `outlook.exe /safe` and remove the addin manually from Options - Add-ins.

Notes
-----

I tested the addin with Outlook 365 and .NET 4.7.2, but in theory it should work with older Outlook versions with no problem.
