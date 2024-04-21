TrayFolder
==========

Windows 11 removed a very useful feature of prior versions: the ability to add a folder as a toolbar, so its contents would be displayed as a popup menu. Since the Windows start menu has long become useless for organizing shortcuts, I was using the folder-in-toolbar feature as a quick way to access my most-used programs with a mouse. I wrote this small tool to replace the missing toolbars for my own use.

What this does is to add an icon to the system tray, and by clicking on that icon, you can access the contents of any specific directory. I'm still using it like Win10: by putting many shortcuts in there (instead of e.g. on desktop) and have this program launched at startup.

Running
-------

You just need to pass the path to your folder as the required first command-line (or shortcut) argument to the program. As an optional second argument, you can specify an icon number from the list [here](https://learn.microsoft.com/en-us/windows/win32/api/shellapi/ne-shellapi-shstockiconid) (let's see how long it takes for Microsoft to break that URL).

You can run as many instances of this program as you want, and each will create their own tray icon and menu. Note that (what Win11 calls) "hidden icons menu" behavior will depend on the order in which instances are launched: if you tell Windows not to hide the icon for the first TrayFolder instance you launch, it will always be the first instance's icon that Windows will display in the system tray, *irrespective of the actual folder path*.

Building
--------

The code is short and in a single file. You can use the provided project file, or just compile on the command line (I recommend the latter, and don't forget C++20 and standard-conforming preprocessor flags). I've only tested with VS2022, but as long as your compiler supports C++20 (for designated initializers) and `__VA_OPT__`, you should be fine.

Dependency-wise, I've used nothing but the Windows standard libraries. I really didn't need this for anything other that Win11, so I haven't tested it on older Windows and some inadvertent minimum version requirement might be there.
