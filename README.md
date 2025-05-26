# CursorVSSync
A tool to quickly switch between Cursor and Visual Studio, activating the opposite window and opening the corresponding file.
Use case is for people that like to use Cursor, but need to run and debug in Visual Studio.

To use, setup a hotkey (using AutoHotkey for example) and put it on a key (I use F7).  Then press the key while in Visual Studio and wanting to switch to Cursor, or vice versa, and it will activate the opposite window and open the corresponding file.
IE:
- Install AutoHotkey
- Create a hotkey file CursorSync.ahk, fill it with:
F7::
Run, C:\projects\bin\CursorSync.exe
return

Then add that file to Startup Apps.

