<div align="center">

## Shell Execute Launch an Application and Open Files


</div>

### Description

You people are gonna LOVE ME LONG TIME for this!

This MODULE's function will load (on Call) any file name and launch it's associated application based on the file extension! I know you've see those SHELL commands that just take a couple of lines but they may or may not work in Windows 98, WindowsNT and Win2000 and guess what MINE DOES! And it does it EVERY SINGLE TIME! (You need to insert your own error traps in the Whatever_Event() of your code. *.txt, *.mdb, *.xls, *.html, *.doc, *.rtf, *.Anything Man! long as there is an associated application for the file extension!

Also has params for Specifying a working directory and you can also set the vbWhateverFocus like vbMinimized or whatever. vbNormailFocus is the default. But if you are a newbie then you don't have to set those params. Just call it like: Call Shell("C:\Wherever\YourFileIs.htm")

NOTE: THIS MOD DOES NOT SHELL TO THE WEB. IT LAUNCHES ANY FILE BUT ONLY FILES THAT ARE RESIDENT ON THE BOX OR NETWORK THAT THE PROGRAM CALLING THE FILE IS ON!!

EM Dixson

http://developer.ecorp.net
 
### More Info
 
Just paste the whole thing into one (1) new module and call name it SURETHING.BAS and call it from you form like this:

Private Sub Command_Click()

Call Shell("Whatever.txt") 'if the file you are calling is in the

'same folder as the progam (.EXE) that you made.

'if not then do it like this:

Call Shell("C:\Path\To\Yourfile.htm")'There are some optional params and if

'you know what you are doing you'll see them.

'Even if you don't know what you are doing and you are completely lame

'and "not all there" this MOD will still work for you. Now I'm beginning

'to sound like a get rich quick ad..(I need sleep).

'Come see me at:

'http://developer.ecorp.net

'FREE Visual Basic Source Code, Tips and Tricks

Clean as they come!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[EM Dixson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/em-dixson.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/em-dixson-shell-execute-launch-an-application-and-open-files__1-3735/archive/master.zip)

### API Declarations

```
'I'm just gonna paste the whole thing into one big enchilada.
'I suggest you do the same.
```


### Source Code

```
'#############################################################
'# This code was written by Emmett Dixson (c)1999. You may alter
'# this code, trade, steal, borrow, lend or give away this code.
'# However, this code has been regisered with the Library of
'# Congress as a literary acheivement and as such excludes it
'# from being known or proclaimed as "PUBLIC DOMAIN".
'#---------------You may NOT remove this header---------------
'#------------------You may NOT SELL this work----------------
'#----YES! You MAY use this work for commercial purposes------
'#---This code MAY NOT be sold or redistributed for profit----
'#-------- I wish you every success in your projects ---------
'#------------------------ Visit me at -----------------------
'#------------------http://developer.ecorp.net ---------------
'#-----------------FREE Visual Basic Source Code -------------
'##############################################################
'For best results paste everything into a NEW MODULE and be sure
'you SAVE the module to your project. I call the module...
'Surething.bas because it won't let you down.
'Works for Win3.x, Win95,Win98,WinNT and EVEN Win2000(don't ask!)
'Here it is and it is Soooo sweet!
'I mean it will call any file man and auto-launch it's
'associated application in any Windows OS.
'All you have to do is enter the path and the
'file-name and extension. It is totally awesome if I do say so
'my self.....LOL.
'Don't change anything...just paste all this code into ONE
'MODULE that you can add to a project.
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Function Shell(Program As String, Optional ShowCmd As Long = vbNormalNoFocus, Optional ByVal WorkDir As Variant) As Long
 Dim FirstSpace As Integer, Slash As Integer
 If Left(Program, 1) = """" Then
  FirstSpace = InStr(2, Program, """")
  If FirstSpace <> 0 Then
   Program = Mid(Program, 2, FirstSpace - 2) & Mid(Program, FirstSpace + 1)
   FirstSpace = FirstSpace - 1
  End If
 Else
  FirstSpace = InStr(Program, " ")
 End If
 If FirstSpace = 0 Then FirstSpace = Len(Program) + 1
 If IsMissing(WorkDir) Then
  For Slash = FirstSpace - 1 To 1 Step -1
   If Mid(Program, Slash, 1) = "\" Then Exit For
  Next
  If Slash = 0 Then
   WorkDir = CurDir
  ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
   WorkDir = Left(Program, Slash)
  Else
   WorkDir = Left(Program, Slash - 1)
  End If
 End If
 Shell = ShellExecute(0, vbNullString, _
 Left(Program, FirstSpace - 1), LTrim(Mid(Program, FirstSpace)), _
 WorkDir, ShowCmd)
 If Shell < 32 Then VBA.Shell Program, ShowCmd 'To raise Error
End Function
```

