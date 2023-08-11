#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
SetBatchLines, -1
CoordMode, Mouse, Screen

#Include ../Neutron.ahk


html =
( ; html
	<h1>现在可以录数据了！</h1>

	<p>
		这个小工具的原理是在你按完一个数字后自动按一个 Tab 跳到下一格，这样就不需要手动选到下一个格子，理论上可以少按一半的按键，从而提高一倍的效率。
	</p>

	<p>
		这个小工具只会在 Excel 和 WPS 里生效。
	</p>

	<h2>功能说明</h2>

	<p>
		1. 你可以用字母区上面的那一列数字键，可以用键盘最右侧的数字小键盘；
	</p>

	<p>
		2. 跳到下一行开头：数字键1左面的“`键”，或小键盘上的乘号“*键”；
	</p>

	<p>
		3. 被试没有填这个格子，要跳到下一个格子：Tab 键或小键盘上的 Enter 键；
	</p>

	<p>
		4. 输多了一位数要删除：选中要删除的那个格子，按减号键，会删掉这个数字并且将后面的格子自动前移一位。减号键用数字键旁边的或小键盘区的都可以；
	</p>

	<p>
		5. 输少了一位数要加进去：选中想要空出来的那个格子，按加号键，会把从这个格子开始的数字都后移一位，把这个格子空出来。加号键用数字键旁边的或小键盘区的都可以；
	</p>

	<p>
		注意：这个小工具有一个限制：一个格子只能输入一个数字，例如年龄这样的两位数，建议分成两个格子输入，在全部输完后，在 Excel 或 WPS 里用“=A1&B1”来合并。
	</p>

	<p>
		注意：如果功能失效，请关掉360再试！！！
	</p>
	
	<p>
		注意：在WPS中可能部分功能会偶尔失效！！！
	</p>

	<p>
		你们用到这个工具的时候作者应该已经毕业了，不是非常严重的问题，一般不会再修复，如果需要联系作者请找刘勇师门的学生，或者发邮件给我，邮箱是dongf1@qq.com。
	</p>	

	<p>
		也欢迎使用我的毕业论文模板。
	</p>	

)

css =
( ; css
	/* Make the title bar dark with light text */
	header {
		background: #333;
		color: white;
	}

	/* Make the content area dark with light text */
	.main {
		background: #444;
		color: white;
	}

	/* Make input boxes follow the dark theme */
	input {
		margin: 0.25em;
		padding: 0.5em;
		border: none;
		background: #333;
		color: white;
		border-radius: 0.25em;
	}
	:-ms-input-placeholder {
		color: silver;
	}
	button {
		background: slategray;
		border: none;
		color: white;
		padding: 0.25em 0.5em;
		border-radius: 0.25em;
	}
)

js =
( ; js
	// Write some JavaScript here
)

title = 心院快速录数据1.0 by 董繁一

; Create a Neutron Window with the given content and save a reference to it in
; the variable `neutron` to be used later.
neutron := new NeutronWindow(html, css, js, title)

; Use the Gui method to set a custom label prefix for GUI events. This code is
; equivalent to the line `Gui, name:+LabelNeutron` for a normal GUI.
neutron.Gui("+LabelNeutron")

; Show the GUI, with an initial size of 640 x 480. Unlike with a normal GUI
; this size includes the title bar area, so the "client" area will be slightly
; shorter vertically than if you were to make this GUI the normal way.
neutron.Show("w640 h480")

; Set up a timer to demonstrate making dynamic page updates every so often.
SetTimer, DynamicContent, 100
return

; The built in GuiClose, GuiEscape, and GuiDropFiles event handlers will work
; with Neutron GUIs. Using them is the current best practice for handling these
; types of events. Here, we're using the name NeutronClose because the GUI was
; given a custom label prefix up in the auto-execute section.
NeutronClose:
ExitApp
return


DynamicContent()
{
	; This function isn't called by Neutron, so we'll have to grab the global
	; Neutron window variable instead of using one from a Neutron event.
	global neutron
	
	; Get the mouse position
	MouseGetPos, x, y
	
	; Update the page with the new position
	neutron.doc.getElementById("ahk_x").innerText := x
	neutron.doc.getElementById("ahk_y").innerText := y
}


#if ((WinActive("ahk_exe EXCEL.EXE")) or (WinActive("ahk_exe et.exe")) or (WinActive("ahk_exe wps.exe")))

Numpad0 up::
SendInput 0{tab}
return

Numpad1 up::
SendInput 1{tab}
return

Numpad2 up::
SendInput 2{tab}
return

Numpad3 up::
SendInput 3{tab}
return

Numpad4 up::
SendInput 4{tab}
return

Numpad5 up::
SendInput 5{tab}
return

Numpad6 up::
SendInput 6{tab}
return

Numpad7 up::
SendInput 7{tab}
return

Numpad8 up::
SendInput 8{tab}
return

Numpad9 up::
SendInput 9{tab}
return


$0 up::
   sendinput 0{tab}
return

$1 up::
   sendinput 1{tab}
return

$2 up::
   sendinput 2{tab}
return

$3 up::
   sendinput 3{tab}
return

$4 up::
   sendinput 4{tab}
return

$5 up::
   sendinput 5{tab}
return

$6 up::
   sendinput 6{tab}
return

$7 up::
   sendinput 7{tab}
return

$8 up::
   sendinput 8{tab}
return

$9 up::
   sendinput 9{tab}
return

NumpadEnter::   ; enter向右空一格
	SendInput {tab}
return

`:: 
NumpadMult::   ;转跳到下一行开头
	SendInput {Down}
	Sleep 50
	SendInput {ctrl down}{Left}{ctrl up}
	tooltip("你回到了下一行开头")
return

; 少输入了一位，把那一位空出来
=::
NumpadAdd::
	SendInput, {CtrlDown}{ShiftDown}{Right}{CtrlUp}{ShiftUp}
	Sleep 350
	SendInput, ^c
	Sleep 50
	SendInput, {Right}
	Sleep 50
	SendInput, ^v
	Sleep 50
	SendInput, {left}
	Sleep 50
	SendInput, {Delete}
	tooltip("可以在这里输入漏输那一个数了")
return


;多输入了一位，选中多的那一位，把那一位删掉
-::
NumpadSub::
	SendInput, {Right}
	Sleep 50
	SendInput, {CtrlDown}{ShiftDown}{Right}{CtrlUp}{ShiftUp}
	Sleep 50
	SendInput, ^c
	Sleep 50
	SendInput, {left}
	Sleep 50
	SendInput, ^v
	Sleep 50
	SendInput, ^{Right}
	Sleep 50
	SendInput, {Delete}
	tooltip("你把多输的那个数删掉了")
return




#if

 tooltip(text := "", time := 1500) {
    Tooltip % text
    if text != ""
        SetTimer, % A_ThisFunc, % "-" time
 }
return