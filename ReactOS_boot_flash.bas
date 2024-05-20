$INCLUDE "RAPIDQ.INC"
$INCLUDE "QXpTheme.INC"
$INCLUDE "QButtonXP.inc"
$INCLUDE "QDrive2024.inc"
'$INCLUDE "QDrive.inc"
$include "Form.inc"

$include "QProcess.inc"
$include "QSHFileOperation.inc"

'$RESOURCE rRufus AS "Rufus01.bmp"
$RESOURCE rRufus AS "RufusEn.bmp"

Application.HintPause=100
Application.HintColor =&H7AFF59'' &H00FF00

Public Const INFINITE = &HFFFF      '  Infinite timeout


dim SHFile as QSHFileOperation

Declare Function GetConsoleCP Lib "kernel32" Alias "GetConsoleCP" () As Long
Declare Function SetConsoleCP Lib "kernel32" Alias "SetConsoleCP" (ByVal wCodePageID As Long) As Long
Declare Function GetConsoleOutputCP Lib "kernel32" Alias "GetConsoleOutputCP" () As Long
Declare Function SetConsoleOutputCP Lib "kernel32" Alias "SetConsoleOutputCP" (ByVal wCodePageID As Long) As Long

Declare Function ShellExecute lib "shell32" alias "ShellExecuteA" _
(ByVal hWnd as long, ByVal lpOperation as string, ByVal lpFile as string, ByVal lpParameters as integer, _
ByVal lpDirectory as string, ByVal nShowCmd as integer) as long


'print" 158: GetConsoleCP=";GetConsoleCP
'print"GetConsoleOutputCP=";GetConsoleOutputCP
CodePageID&=1251

'SetConsoleCP(CodePageID&)
'SetConsoleOutputCP(1251)

'print"420:консоль GetConsoleCP 1251=";GetConsoleCP
'print"421:консоль GetConsoleOutputCP 1251=";GetConsoleOutputCP

'SetConsoleCP(866)
'SetConsoleOutputCP(866)



'--- Declarations --- !!!
Declare Sub DelFiles
Declare Sub MoveFiles (Sender as QButtonXP)
Declare Sub GetFlashLetter
Declare Sub MainTabChange

Declare Sub CopyISOToFlash
Declare Sub CheckBoxOnClick
Declare Sub CheckBox7zOnClick
Declare Sub formOnShow
Declare Sub DownLoadISO
Declare Sub RunRufus
Declare Sub RunWizard
Declare Sub DloadISOGBoxRepaint
Declare function DownLoadURL2 (sUrl as string, dFileName as string, ChbIdx as long, TotalByt as long) as long
Declare Sub Timer1Over
Declare Sub CreateFreeldrIni

'Declare Sub CopyISO '(ISOFN$,FlashDriveN$)

'dim process as QProcess
'process.OnProcess=CopyISO '(ISOFN$,FlashDriveN$) 'process routine

ISOFN$=""

$OPTION ICON "ROS16w.ico"

dim drive as QDrive

DIM Timer1 AS QTimer

declare sub TimerOver 
Timer1.Interval = 2400 
Timer1.Enabled = 0 'True 
Timer1.OnTimer = Timer1Over 

'DIM TimerCopy AS QTimer

'declare sub TimerCopyOver 
'TimerCopy.Interval = 1000 
'TimerCopy.Enabled = 0 'True 
'TimerCopy.OnTimer = TimerCopyOver 


StartPath$ = COMMAND$(0)-Application.ExeName    
print"23:StartPath$ =";StartPath$ 

FreeldrFN$=StartPath$+"freeldr.ini.tpl"
arcPAth$=StartPath$+"archive"
RufusPath$=StartPath$+"Rufus"


FlashDrive$=""
g_driveCount=0 'drive.Count
'print"32:g_driveCount=";g_driveCount



'DIM ListBox AS QListBox
Declare Sub FlashDriveListBoxonclick
Declare Sub ShowFileInfo

dim DownGauge(10)  as QGauge
dim DownLAbel(10)  as QLAbel
Declare Sub OntimerSh

dim Timer11 as QTimer
Timer11.enabled=0
Timer11.Ontimer=OntimerSh

CREATE Form AS QFORM
	Caption = "ReactOS boot flash wizard"
	top=60
	Width = 350
	Height = 700
	left=1500
	top=Screen.Height/2-420
	ShowHint=1
	color =4444444
	font.name="Arial"
	
	'Center
	onshow=formOnShow
	
	
	CREATE PanelRufus AS QPANEL
		Left = 1
		Top = 1
		Caption = "Panel1"
		TabOrder = 8
		align=altop 'albottom 
		'color=cllb
		height=80
		
		CREATE LabelRufus AS QLABEL
			Caption = "Set Rufus properties as in the picture below"+ crlf+_
			"Boot selection=ReactOS"+ crlf+_
			"File system=FAT32"+ crlf+_
			"and format flash drive"
			Left = 1
			Top = 1
			height=45
			'Transparent = 1
			color=cly
			align=altop 'albottom 
			alignment=taCenter
			font.size=8
			font.name="Arial"
			autosize=1
			'visible=0
		END CREATE
		
		
		CREATE BtnRufus AS QButtonXP
			XP=1
			Caption = "Run Rufus and format bootable flash drive"
			Left = 100
			height=30
			Top = 2
			align=altop 'albottom
			color=clg
			font.size=11 '14
			font.name="Arial"
			font.bold=1
			font.color=clg
			OnClick=RunWizard 'RunRufus
			hint="Run Rufus,set it's properties as in the picture below and format flash drive"
			visible=1
			
		END CREATE
		
	END CREATE
	CREATE ImagePanel AS QPANEL
		Left = 1
		Top = 100
		'Caption = "Panel1"
		TabOrder = 8
		align=altop 'albottom 
		width=445
		Height = 690
		color=cllb
		CREATE ImageRufus AS QIMAGE
			Left = 10
			Top = 10
			align=alclient 'top'alleft
			width=445
			Height = 690
			BMPHandle = rRufus
			visible=1
			'Center=1
		END CREATE
	END CREATE
	
	CREATE PanelSpace AS QPANEL
		Left = 1
		Top = 20
		'Caption = "Panel1"
		TabOrder = 8
		align=altop 'albottom
		height=10
		color=clo
		visible=0
	END CREATE
	
	'===================================================================!!!! 
	
	'============================================ 
	
	CREATE FlashPane AS QPANEL
		Left = 1
		Top = 2000
		'Caption = "Panel1"
		TabOrder = 8
		align=alclient 'top 'albottom
		height=120
		color=cly '&HFFEFD8 '33333 'clBtnFace
		visible=1  
		
		CREATE CreateFldrBtn AS  QButtonXP
			XP=1
			Caption = "Create Freeldr Ini"
			'Left = 1
			'Top = 1
			height=30
			align=albottom 'top'
			OnClick=CreateFreeldrIni
			visible=1
			font.size=13
			font.bold=1
		END CREATE
		
		CREATE FlashLabel AS QLABEL
			Caption = "Flash Drive list"
			Top = 1
			align=altop 'albottom
			alignment=taCenter
			Transparent = 0
			font.size=13
			font.bold=1
			font.color=clb
			color=&HFFEFD8 '33333 'clBtnFace
			
		END CREATE
		
		CREATE FlashDriveListBox AS QLISTBOX
			Left = 1
			TabOrder = 8
			top=2
			Height = 40 'Form.ClientHeight
			align=altop'alclient
			font.name="Courier"
			onclick=FlashDriveListBoxonclick
			color=&HFFEFD8 '33333 'clBtnFace
			visible=1
		END CREATE
		
		CREATE FlashFileList AS QFileListBox
			ShowIcons = True
			Mask = "*.*"
			'OnDblClick = ExecuteApplication
			'onclick=ShowPrjTxt
			top=4
			Height = 100 'Form.ClientHeight
			align=alclient 'altop' 
			color=&HFFEFD8 '33333 'clBtnFace
			MultiSelect=1
			ExtendedSelect=1
			font.name="Courier"
			visible=1
			
		END CREATE
		
		
	END CREATE
	
	
	CREATE PanelDownISO AS QPANEL
		Left = 1
		Top = 20
		'Caption = "Panel1"
		TabOrder = 8
		align=altop 'albottom
		height=320
		color=clg 'clBtnFace
		visible=0  
		CREATE BtnDiowloadISO AS QButtonXP
			'XP=1
			Caption = "DownLoad ISO"
			Left = 100
			height=30
			Top = 1
			align=altop 'albottom
			color=clg
			font.size=13
			font.bold=1
			font.name="Arial"
			font.color=clr
			OnClick=DownLoadISO 
			'hint="Install selected core"
			visible=0
			
		END CREATE
		
		CREATE PanelCheckBox AS QPANEL
			align=altop 'albottom alclient '
			Left = 1
			Top = 10
			'Caption = "Panel1"
			height=150
			TabOrder = 8
			'color=cly
			
			
			CREATE CheckBox1 AS QCHECKBOX
				Caption = "LiveCD x64 dbg"
				Left = 10
				Top = 20
				width=100
				TabOrder = 8
				checked=0
				OnClick=CheckBoxOnClick
			END CREATE
			
			CREATE CheckBox2 AS QCHECKBOX
				Caption = "LiveCD x86 dbg"
				Left = 10
				Top = CheckBox1.top+20*1
				TabOrder = 8
			END CREATE
			
			CREATE CheckBox3 AS QCHECKBOX
				Caption = "LiveCD x86 release"
				Left = 10
				Top = CheckBox1.top+20*2
				TabOrder = 8
				width=110
				checked=1
			END CREATE
			
			CREATE CheckBox4 AS QCHECKBOX
				Caption = "BootCD x64 dbg" ' - to install on HDD!!! Be carefull!!!
				Left = 10
				Top = CheckBox1.top+20*3
				TabOrder = 8
				'width=250
				font.color=clr
				font.Underline=1 'fsUnderline
				checked=0
			END CREATE
			
			CREATE CheckBox5 AS QCHECKBOX
				Caption = "BootCD x86 dbg"
				Left = 10
				Top = CheckBox1.top+20*4
				TabOrder = 8
				font.Underline=1 'fsUnderline
				font.color=clr
				checked=0
				'width=250
			END CREATE
			
			CREATE CheckBox6 AS QCHECKBOX
				Caption = "BootCD x86 release"
				'width=250
				Left = 10
				Top = CheckBox1.top+20*5
				TabOrder = 8
				font.Underline=1 'fsUnderline
				font.color=clr
				checked=1
				width=115
			END CREATE
			
			
			' https://iso.reactos.org/bootcd/latest-x64-msvc-win-dbg
			' https://iso.reactos.org/bootcd/latest-x86-gcc-lin-rel
			' https://iso.reactos.org/bootcd/latest-x86-msvc-win-dbg
			' https://iso.reactos.org/bootcd/latest-x86-gcc-lin-dbg
			
			'!https://iso.reactos.org/livecd/latest-x64-msvc-win-dbg
			'!https://iso.reactos.org/livecd/latest-x86-gcc-lin-rel
			'!!https://iso.reactos.org/livecd/latest-x86-msvc-win-dbg
			
			'https://iso.reactos.org/livecd/latest-x86-gcc-lin-dbg
			'https://iso.reactos.org/livecd/latest-x86-gcc8.3-lin-dbg
			
			
			CREATE DloadISOGBox AS QCanvas ' red border  
				top=60
				height=150
				OnPaint=DloadISOGBoxRepaint
				align=altop
				visible=1
			END CREATE
			
			CREATE LabelDLoad AS QLABEL
				Caption = "Download required ISO"
				Left = PanelDownISO.width/2-90
				'Top = DloadISOGBox.top+DloadISOGBox.height-29
				Top = 0 'PanelSpace1.top+8
				'Transparent = 1
				'color=clw
				'align=altop 'albottom 
				alignment=taCenter
				font.size=10
				font.name="Arial"
				'visible=0
			END CREATE
			
		END CREATE
		
		
		CREATE ISOListLabel AS QLABEL
			Caption = "Downloaded ISO List"
			Left = 1
			Top = 41
			Transparent = 0
			align=altop
			font.size=13
			font.bold=1
			font.color=clb
			alignment=taCenter 
			color=&HCFFFCB
		END CREATE
		
		
		CREATE PanelDLHLP AS QPANEL
			Left = 1
			Top = 51
			'Caption = "Boot CD - Boot CDs are designed to install ReactOS onto an HDD"
			TabOrder = 8
			'color=&HCFFFCB
			height=20
			visible=1
			align=altop
			
			
			CREATE btnDelFile AS QButtonXP
				XP=1
				Caption = "Del."
				Left = 100
				height=30
				Top = 70
				width=40
				align=alright
				color=clr
				font.size=10
				font.name="Arial"
				font.color=clr
				OnClick=MoveFiles 'DelFiles
				ShowHint=1
				hint="Delete selected files"
				visible=1
				tag=455
				
			END CREATE
			
			CREATE btnMoveFile AS QButtonXP
				XP=1
				Caption = "Move"
				width=40
				Left = 100
				height=30
				align=alright
				color=clr
				font.size=10
				font.name="Arial"
				font.color=clr
				OnClick=MoveFiles 
				ShowHint=1
				hint="Move selected files to archive"
				visible=1
				tag=473
				
			END CREATE
			
			CREATE CheckBox7z AS QCHECKBOX
				Caption = "7z"
				'width=250
				Left = 5
				Top = 2
				TabOrder = 8
				'font.Underline=1 'fsUnderline
				font.color=clr
				checked=1
				width=40
				OnClick=CheckBox7zOnClick
				align=alright 'top 'albottom 
				visible=1
				
			END CREATE
			CREATE CheckBoxISO AS QCHECKBOX
				Caption = "ISO"
				'width=250
				Left = 40
				Top = 2
				TabOrder = 8
				'font.Underline=1 'fsUnderline
				font.color=clr
				checked=1
				width=40
				OnClick=CheckBox7zOnClick
				align=alright 'top 'albottom 
				visible=1
			END CREATE
			
			CREATE LabelFileInfo AS QLABEL
				Caption = "File"
				align=alright 'top 'albottom 
				alignment=taCenter
				font.size=10
				font.name="Arial"
				color=clw
				width=40
				visible=0
			END CREATE
			
			CREATE MainTab AS QTabControl
				AddTabs "Latest", "Archive"
				align=alclient
				OnChange = MainTabChange
				HotTrack = True
				visible=1
				
			END CREATE
			
			
			
			
		END CREATE
		
		CREATE btnCopyISOToFlash AS QButtonXP
			XP=1
			Caption = "Copy selected ISO to flash"
			Left = 100
			height=30
			Top = 70
			align=albottom 'top 'albottom
			color=clr
			font.size=13
			font.bold=1
			font.name="Arial"
			font.color=clr
			OnClick=CopyISOToFlash 
			ShowHint=1
			hint="Boot CD - Boot CDs are designed to install ReactOS onto an HDD"
			visible=1
			
		END CREATE
		
		
		CREATE ISOList AS QFileListBox
			ShowIcons = True
			Mask = "*.7z;*.iso"
			'OnDblClick = ExecuteApplication
			onclick=ShowFileInfo
			top=111
			Height = 152 'Form.ClientHeight
			align=alclient 'altop'
			'color=cls
			MultiSelect=1
			ExtendedSelect=1
			font.name="Courier"
			visible=1
			color=&HCFFFCB
			
		END CREATE
		
		
		
	END CREATE
	
	CREATE LogEdit AS QRICHEDIT
		Left = 1
		Top = 11
		TabOrder = 8
		HideScrollBars=0
		HideSelection=0
		ScrollBars=3 'ssBoth = 3  
		Wordwrap=0
		align=albottom
		height=100
	END CREATE
	
	CREATE FrLdrEdit AS QRICHEDIT
		Left = 1
		Top = 11
		TabOrder = 8
		HideScrollBars=0
		HideSelection=0
		ScrollBars=3 'ssBoth = 3  
		Wordwrap=0
		align=alleft
		width=300
		visible=0
	END CREATE
	
	
	
END CREATE

flRuf=0

for i=0 to 5
	DownGauge(i).parent=PanelCheckBox 'PanelDownISO
	DownGauge(i).top=20*i+CheckBox1.top+3
	DownGauge(i).left=130
	DownGauge(i).width=90
	DownGauge(i).ForeColor=clb
	DownGauge(i).Kind=gkHorizontalBar
	DownGauge(i).height=15
	
	DownLAbel(i).parent=PanelCheckBox
	DownLAbel(i).top=20*i+CheckBox1.top+3
	DownLAbel(i).left=225
	DownLAbel(i).height=15
	
	
next i

ISOPath$=StartPath$+"ISO\"  ' archive
ISOList.Directory=ISOPath$

print"324:ISOList.Directory=";ISOList.Directory


call GetFlashLetter

Timer1.Enabled = True 


Form.ShowModal  

brem 0
============
"Boot CD" - Boot CDs are designed to install ReactOS onto an HDD and enjoy 
the new features since last release. You will need the ISO only for the installation. 
This is the recommended variant to install into a VM (VirtualBox, VMWare, QEMU).

"Live CD" - Live CDs allow you to use ReactOS without installing it. 
It can be used to test ReactOS in case your HDD is not detected during installation, 
or if you have no alternative system/VMs to install it on.

"Debug" - Debug versions include debugging information, use these versions to test, 
produce logs and report bugs. This is the recommended variant to install to report bugs.

"Release" - Release versions include no debugging information. These versions are smaller, 
but cannot be used to produce logs.
===========
erem


'!******************************************
Sub RunRufus

'print"72:flRuf=";flRuf
'call  AddClrString ("501:flRuf="+str$(flRuf), clred, LogEdit)



if flRuf=0 then
	flRuf=1
	BtnRufus.color=clr
	BtnRufus.XP=0
	BtnRufus.font.color=clgrey
	BtnRufus.font.bold=0
	'BtnRufus.visible=0
	BtnRufus.caption="Back to run Rufus"
	BtnRufus.visible=1
	Form.Repaint
	DoEvents
	'call  AddClrString ("613:StartPath$+rufus.exe="+(StartPath$+"rufus.exe"), clred, LogEdit)
	
	'PID=
	'shell (StartPath$+"rufus.exe")
	'run (StartPath$+"rufus.exe")
	
	'PID=ShellExecute (0,"open",StartPath$+"rufus.exe","","",SW_SHOWNORMAL)
	PID=ShellExecute (0,"open",RufusPath$+"\rufus.exe","","",SW_SHOWNORMAL)
	
	'DoEvents
	WaitForSingleObject (PID, 200)
	'sleep 5
	
	
else
	flRuf=0
	BtnRufus.color=clg
	BtnRufus.caption="Run Rufus and format bootable flash drive""
	BtnRufus.font.color=clblack
	BtnRufus.XP=1
	BtnRufus.font.bold=1
	'PanelDownISO.visible=0
	'PanelDLHLP.visible=0
	'ISOList.visible=0
	
	ImagePanel.visible=1
	ImageRufus.Repaint
	Form.Repaint
	DoEvents
	
end if



End Sub 
'!******************************************
Sub RunWizard
'if FlashDrive$="" then 
'ShowMessage ("437: FlashDrive not found")
'exit sub
'end if


RunRufus
ImagePanel.visible=0

FlashFileList.update

' ImageRufus.visible=0

'BtnRufus.color=clsilver
BtnRufus.XP=1

BtnDiowloadISO.XP=1 
PanelDownISO.visible=1
PanelDLHLP.visible=1
ISOList.visible=1

End Sub 

'!******************************************
Sub DownLoadISO

' https://iso.reactos.org/bootcd/latest-x64-msvc-win-dbg
' https://iso.reactos.org/bootcd/latest-x86-gcc-lin-rel
' https://iso.reactos.org/bootcd/latest-x86-msvc-win-dbg
' https://iso.reactos.org/bootcd/latest-x86-gcc-lin-dbg

'https://iso.reactos.org/livecd/latest-x64-msvc-win-dbg
'https://iso.reactos.org/livecd/latest-x86-gcc-lin-rel
'https://iso.reactos.org/livecd/latest-x86-msvc-win-dbg

'https://iso.reactos.org/livecd/latest-x86-gcc-lin-dbg
'https://iso.reactos.org/livecd/latest-x86-gcc8.3-lin-dbg

BtnDiowloadISO.enabled=0


if CheckBox1.checked=1 then 
	url$="https://iso.reactos.org/livecd/latest-x64-msvc-win-dbg"
	
	dFileName$=StartPath$+"ISO\reactos-livecd-0.4.15-x64-dbg.7z"
	'call  AddClrString ("541:dFileName$="+(dFileName$), clred, LogEdit)
	call DownLoadURL2 (url$, dFileName$, 0,27000000)
	ISOList.Update
	
end if

if CheckBox2.checked=1 then 
	
	url$="https://iso.reactos.org/livecd/latest-x86-msvc-win-dbg"
	
	dFileName$=StartPath$+"ISO\reactos-livecd-0.4.15-x86-dbg.7z"
	'call  AddClrString ("303:dFileName$="+(dFileName$), clred, LogEdit)
	call DownLoadURL2 (url$, dFileName$,1, 25000000)
	ISOList.Update
	
end if

if CheckBox3.checked=1 then 
	
	url$="https://iso.reactos.org/livecd/latest-x86-gcc-lin-rel"
	
	dFileName$=StartPath$+"ISO\reactos-livecd-0.4.15-x86-rel.7z"
	'call  AddClrString ("303:dFileName$="+(dFileName$), clred, LogEdit)
	call DownLoadURL2 (url$, dFileName$, 2,33000000)
	ISOList.Update
	
end if

if CheckBox4.checked=1 then 
	
	url$="https://iso.reactos.org/bootcd/latest-x64-msvc-win-dbg"
	
	dFileName$=StartPath$+"ISO\reactos-bootcd-0.4.15-x64-dbg.7z"
	'call  AddClrString ("303:dFileName$="+(dFileName$), clred, LogEdit)
	call DownLoadURL2 (url$, dFileName$, 3,53288512)
	ISOList.Update
	
end if

if CheckBox5.checked=1 then 
	
	url$="https://iso.reactos.org/bootcd/latest-x86-msvc-win-dbg"
	
	dFileName$=StartPath$+"ISO\reactos-bootcd-0.4.15-x86-dbg.7z"
	'call  AddClrString ("303:dFileName$="+(dFileName$), clred, LogEdit)
	call DownLoadURL2 (url$, dFileName$, 4,48288512)
	ISOList.Update
	
end if
if CheckBox6.checked=1 then 
	
	url$="https://iso.reactos.org/bootcd/latest-x86-gcc-lin-rel"
	
	dFileName$=StartPath$+"ISO\reactos-bootcd-0.4.15-x86-rel.7z"
	'call  AddClrString ("303:dFileName$="+(dFileName$), clred, LogEdit)
	call DownLoadURL2 (url$, dFileName$, 5,68288512)
	ISOList.Update
	
end if

ISOList.Update

'! unpack 7z files --------------------
CHDIR ISOList.Directory
call  AddClrString ("717:ISOList.Directory="+(ISOList.Directory), clb, LogEdit)


for i617=0 to ISOList.ItemCount-1
	zzz=1 '(ISOList.Selected(i617))
	if zzz>0 then
		fext$=StripFileExt (ISOList.Item(i617))
		'call  AddClrString ("678:fext$="+(fext$), clred, LogEdit)
		if fext$=".7z" then
			ff1$=ISOList.Directory+"\"+ISOList.Item(i617)
			'call  AddClrString ("626:ff1$="+(ff1$), clred, LogEdit)
			bat7z$=ISOList.Directory+"\7z.exe x -y "+ ISOList.Directory+"\"+ISOList.Item(i617)
			'call  AddClrString ("629:7zbat$="+(bat7z$), clred, LogEdit)
			savestring(bat7z$,ISOList.Directory+"\7z.bat")
			'shell bat7z$ >ISOList.Directory+"\_zzzzz"
			pid=shell (bat7z$) 'э+" > "+ISOList.Directory+"\_zzzzz"
			WaitForSingleObject (PID, 5000)
		end if
	end if
next i617
ISOList.Update


for i617=0 to ISOList.ItemCount-1
	fext$=StripFileExt (ISOList.Item(i617))
	if fext$=".7z" then
		ff1$=ISOList.Directory+"\"+ISOList.Item(i617)
		kill ff1$ 
	end if
next i617
ISOList.Update


CHDIR StartPath$

BtnDiowloadISO.enabled=1


End Sub 


'!******************************************
Sub DloadISOGBoxRepaint

'DloadISOGBox.RoundRect (1,1,225,103,19,19,clb)
'DloadISOGBox.RoundRect (3,3,223,101,16,16,clWhite)

DloadISOGBox.RoundRect (1,5,DloadISOGBox.width-2,DloadISOGBox.height-2,19,19,clgrey)
DloadISOGBox.RoundRect (3,7,DloadISOGBox.width-4,DloadISOGBox.height-4,16,16,clBtnFace)

End Sub 
'!******************************************
Sub formOnShow
CheckBox4.visible=0 
doevents
CheckBox4.font.color=clr
CheckBox4.visible=1 
doevents

End Sub 


'!*****************************************
function DownLoadURL2 (sUrl as string, dFileName as string, ChbIdx as long, TotalByt as long) as long
call  AddClrString ("322:dFileName="+(dFileName), clred, LogEdit)
' скачивает файл с заданного sURL и сохраняет под именем dFileName
'возвращает число скачанных байт TotalReadBytes или -1 в случае ошибки

dim FileDst as QFileStream
FileDst.Open(dFileName, fmCreate)


Dim hOpen As Long, hFile As Long, sBuffer As String, ReadBytes As Long, TotalReadBytes As Long
defint sBufferSize=128*64 '2048*64
filetxt$=""

ReadBytes=0
TotalReadBytes=0

'Create a buffer for the file we're going to download
sBuffer = Space(sBufferSize)'+"<<<"


'Create an internet connection
hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
''call  AddClrString ("350:hOpen="+str$(hOpen), clred, LogEdit)

hFile = InternetOpenUrl(hOpen, sURL, vbNullString, 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
''call  AddClrString ("353:hFile="+str$(hFile), clred, LogEdit)
'PanelBytes.caption=sURL
doevents
'Read the first 1000 bytes of the file
zzz=InternetReadFile (hFile, sBuffer, sBufferSize, ReadBytes)
''call  AddClrString ("358:zzz="+str$(zzz), clred, LogEdit)
filetxt$=left$(sBuffer,ReadBytes)
TotalReadBytes=ReadBytes
OldTB=ReadBytes

FileDst.WriteBinStr(filetxt$,ReadBytes)


''call  AddClrString ("369:TotalByt="+str$(TotalByt), clred, LogEdit)
'DownGauge(ChbIdx).max=TotalByt
while ReadBytes>0
	zzz=InternetReadFile (hFile, sBuffer, sBufferSize, ReadBytes)
	'filetxt$=filetxt$+left$(sBuffer,ReadBytes)
	TotalReadBytes=TotalReadBytes+ReadBytes
	''call  AddClrString ("367:TotalReadBytes="+str$(TotalReadBytes), clred, LogEdit)
	filetxt$=left$(sBuffer,ReadBytes)
	
	FileDst.WriteBinStr(filetxt$,ReadBytes)
	
	if TotalReadBytes>(OldTB+400000) then
		OldTB=TotalReadBytes
		''call  AddClrString ("386:OldTB="+str$(OldTB), clred, LogEdit)
		DownGauge(ChbIdx).Position=100*TotalReadBytes/(TotalByt)  ' 28 052 317
		DownLAbel(ChbIdx).caption="Read "+str$(TotalReadBytes)+" b"
		'PanelBytes.caption=str$(ReadBytes)+"... Total read "+str$(TotalReadBytes)
		doevents
	end if
	
	
wend

DownGauge(ChbIdx).Position=100*TotalReadBytes/(TotalByt)  ' 28 052 317
DownLAbel(ChbIdx).caption="Read "+str$(TotalReadBytes)+" b"


InternetCloseHandle (hFile)
InternetCloseHandle (hOpen)

'zzj=savestring(filetxt$,dFileName)
FileDst.close

DownLoadURL1=TotalReadBytes

End function


'!******************************************
Sub OntimerSh

End Sub 

'!******************************************
Sub CheckBoxOnClick

BtnDiowloadISO.enabled=1

End Sub 

'!******************************************
Sub CopyISOToFlash

if FlashDrive$="" then
	ShowMessage ("721:Select flash drive to copy")
	call GetFlashLetter 
	exit sub
end if

'! copy selected ISO to flash ---------------
'call  AddClrString ("721:ISOList.SelCount="+str$(ISOList.SelCount), clred, LogEdit)

if ISOList.SelCount=0 then
	ShowMessage ("729:Select ISO files to copy")
	exit sub
end if


for i618=0 to ISOList.ItemCount-1
	zzz=(ISOList.Selected(i618))
	ISOFN$=ISOList.Item(i618)
	
	if zzz>0 then
		fext$=StripFileExt (ISOList.Item(i618))
		if fext$=".iso" then
			ISOFN$=ISOList.Item(i618)
			'call  AddClrString ("736:ISOFN$="+(ISOFN$), clred, LogEdit)
			'secb=SecTime(Time$)
			''call  AddClrString ("750:secb="+str$(secb), clred, LogEdit)
			
			SHFile.Copy (ISOList.Directory+"\"+ISOFN$,FlashDrive$+"\"+ISOFN$)
			'sece=SecTime(Time$)
			'tts=sece-secb
			''call  AddClrString ("759:tts="+str$(tts), clred, LogEdit)
			
		end if
	end if
next i618


SHFile.Copy (StartPath$+"freeldr.sys",FlashDrive$+"\freeldr.sys")


FlashFileList.update


'!!ISOList.Update

End Sub 


'!******************************************
Sub FlashDriveListBoxonclick

fldr$=FlashDriveListBox.Item(FlashDriveListBox.ItemIndex)
fldr$=left$(fldr$,instr(fldr$," ")-1)

FlashDrive$=fldr$
call  AddClrString ("736:FlashDrive$="+(FlashDrive$), clred, LogEdit)

if direxists (fldr$) >0 then
	FlashFileList.Mask = "*.*"
	FlashFileList.Directory=fldr$
else
	FlashFileList.Mask = "*.%%%"
end if

'FlashDriveListBox.update
FlashFileList.update

End Sub 


brem 0
================================================================
7-Zip 23.01 (x86) : Copyright (c) 1999-2023 Igor Pavlov : 2023-06-20

Usage: 7z <command> [<switches>...] <archive_name> [<file_names>...] [@listfile]

<Commands>
a : Add files to archive
b : Benchmark
d : Delete files from archive
e : Extract files from archive (without using directory names)
h : Calculate hash values for files
i : Show information about supported formats
l : List contents of archive
rn : Rename files in archive
t : Test integrity of archive
u : Update files to archive
x : eXtract files with full paths

<Switches>
-- : Stop switches and @listfile parsing
-ai[r[-|0]]{@listfile|!wildcard} : Include archives
-ax[r[-|0]]{@listfile|!wildcard} : eXclude archives
-ao{a|s|t|u} : set Overwrite mode
-an : disable archive_name field
-bb[0-3] : set output log level
-bd : disable progress indicator
-bs{o|e|p}{0|1|2} : set output stream for output/error/progress line
-bt : show execution time statistics
-i[r[-|0]]{@listfile|!wildcard} : Include filenames
-m{Parameters} : set compression Method
-mmt[N] : set number of CPU threads
-mx[N] : set compression level: -mx1 (fastest) ... -mx9 (ultra)
-o{Directory} : set Output directory
-p{Password} : set Password
-r[-|0] : Recurse subdirectories for name search
-sa{a|e|s} : set Archive name mode
-scc{UTF-8|WIN|DOS} : set charset for console input/output
-scs{UTF-8|UTF-16LE|UTF-16BE|WIN|DOS|{id}} : set charset for list files
-scrc[CRC32|CRC64|SHA1|SHA256|*] : set hash function for x, e, h commands
-sdel : delete files after compression
-seml[.] : send archive by email
-sfx[{name}] : Create SFX archive
-si[{name}] : read data from stdin
-slp : set Large Pages mode
-slt : show technical information for l (List) command
-snh : store hard links as links
-snl : store symbolic links as links
-sni : store NT security information
-sns[-] : store NTFS alternate streams
-so : write data to stdout
-spd : disable wildcard matching for file names
-spe : eliminate duplication of root folder for extract command
-spf[2] : use fully qualified file paths
-ssc[-] : set sensitive case mode
-sse : stop archive creating, if it can't open some input file
-ssp : do not change Last Access Time of source files while archiving
-ssw : compress shared files
-stl : set archive timestamp from the most recently modified file
-stm{HexMask} : set CPU thread affinity mask (hexadecimal number)
-stx{Type} : exclude archive type
-t{Type} : Set type of archive
-u[-][p#][q#][r#][x#][y#][z#][!newArchiveName] : Update options
-v{Size}[b|k|m|g] : Create volumes
-w[{path}] : assign Work directory. Empty path means a temporary directory
-x[r[-|0]]{@listfile|!wildcard} : eXclude filenames
-y : assume Yes on all queries
======================================================================
erem
'!******************************************
Sub ShowFileInfo
CHDIR ISOList.Directory

'ISOList.FileName
call  AddClrString ("782:ISOList.FileName="+(ISOList.FileName), cldg, LogEdit)
FileName$ = DIR$(ISOList.FileName, 0)
call  AddClrString ("784:FileName$="+(FileName$), clo, LogEdit)

CHDIR StartPath$

End Sub 

'!******************************************
Sub CheckBox7zOnClick

IsoList.Mask = "*.7z;*.iso"

if CheckBox7z.checked=1 then
	
else
	IsoList.Mask =IsoList.Mask- "*.7z"
end if

if CheckBoxISO.checked=1 then
else
	IsoList.Mask =IsoList.Mask- "*.iso"
	
end if

ISOList.Update 
End Sub 


'!******************************************
Sub GetFlashLetter 

'g_driveCount=0
drive.GetDrives

if g_driveCount<>drive.Count then
	g_driveCount=drive.Count
	FlashDriveListBox.clear
	FlashFileList.update
else
	goto  GFSubExit
end if


For i=1 to drive.Count
	'print"795:drive.Count=";drive.Count
	''call  AddClrString ("852:drive.Count="+str$(drive.Count), clred, LogEdit)
	'if drive.GetType(i)=DRIVE_FIXED then number=i
	z=drive.GetType(i)
	
	Select case drive.GetType(i)
	case DRIVE_REMOVABLE
		DriveType$="REMOVABLE"
		
	case DRIVE_FIXED
		DriveType$="FIXED"
		
	case DRIVE_REMOTE
		DriveType$="REMOTE"
		
	case DRIVE_CDROM
		DriveType$="CDROM"
		
	case DRIVE_RAMDISK
		DriveType$="RAMDISK"
		
	case else
		
	end select
	
	
	'DRIVE_REMOVABLE (2)
	'DRIVE_FIXED(3)
	'DRIVE_REMOTE(4)
	'DRIVE_CDROM(5)
	'DRIVE_RAMDISK(6)
	
	driveSize$=str$(ROUND(drive.GetSize(i)/1024/1024))
	driveFreeSpace$=str$(ROUND(drive.GetFreeSpace(i)/1024/1024))
	
	'print"362---- :DriveType$=";DriveType$
	
	DriveData$=drive.name(i)+" "+ _
	"Size:"+driveSize$+"Mb"+" "+ _
	"Free:"+driveFreeSpace$+"Mb"+ _
	" "+DriveType$+" "
	
	'print"833:DriveData$=";DriveData$
	
	
	'print"841:FlashDrive$=";FlashDrive$
	
	if DriveType$="REMOVABLE" then ' только флешки
		'FlashDriveListBox.AddItems DriveData$
		'FlashDrive$=drive.name(i)
		'print"846:FlashDrive$=";FlashDrive$
		
		if direxists (drive.name(i))>0 then
			FlashDriveListBox.AddItems DriveData$
			FlashFileList.Directory=drive.name(i) 'FlashDrive$
		else
			'ShowMessage ("850: Path not found "+FlashDrive$)
			''call  AddClrString ("850:Path not found "+(FlashDrive$), clred, LogEdit)
			'Timer1.enabled=0
		end if
	end if
	
next i


GFSubExit:

End Sub 

'!******************************************
Sub Timer1Over
''call  AddClrString ("874:drive.Count="+str$(drive.Count), clred, LogEdit)
''call  AddClrString ("876:g_driveCount="+str$(g_driveCount), clred, LogEdit)
GetFlashLetter


End Sub 

'!******************************************
Function OsName$ (OsFN$ as string) as string

rn$=field$(OsFN$,"-",6)
'call  AddClrString ("1022:rn$="+(rn$), cldg, LogEdit)

OsFN$=OsFN$-"reactos-"-"gcc-lin-rel"-rn$
OsFN$=OsFN$-"-"-".iso"
'call  AddClrString ("1039:OsFN$="+(OsFN$), 44444, LogEdit)

OsName$=OsFN$

End Function  

'!******************************************
Sub CreateFreeldrIni
'call  AddClrString ("1031:FreeldrFN$="+(FreeldrFN$), clred, LogEdit)

dim FreeldrIniLst as QStringList
FreeldrIniLst.LoadFromFile (FreeldrFN$)

dim OSNm$(128) as string

kill FlashDrive$+"\"+"freeldr.ini"

'FreeldrIniLstTxt$=FreeldrIniLst.Text

'!FlashFileList.Mask = "*.iso"

OS$=FlashFileList.Item(0)

DefaultOS$=OsName$(OS$)
'call  AddClrString ("1032:DefaultOS$="+(DefaultOS$), clo, LogEdit)

FreeldrIniLst.Text=replacesubstr$(FreeldrIniLst.Text,"DefaultOS$",DefaultOS$)

TimeOut$="5"
FreeldrIniLst.Text=replacesubstr$(FreeldrIniLst.Text,"TimeOut$",TimeOut$)
''call  AddClrString ("1046:FreeldrIniLstTxt$="+(FreeldrIniLstTxt$), 52222, LogEdit)


FreeldrIniLst.AddItems "[Operating Systems]"



'!перебираем по списку iso на флешке и формируем имена OS

for i=0 to FlashFileList.ItemCount-1
	
	OSNm$(i)=FlashFileList.Item(i)
	os$=OSNm$(i)
	OSNm$(i)=OsName$(os$)
	
	FreeldrIniLst.AddItems OSNm$(i)+"="+qt+OSNm$(i)+qt
	
	
next i

FreeldrIniLst.AddItems ";===="

for i=0 to FlashFileList.ItemCount-1
	
	FreeldrIniLst.AddItems "["+OSNm$(i)+"]" 
	
	FreeldrIniLst.AddItems "BootType=Windows2003"
	FreeldrIniLst.AddItems "SystemPath=ramdisk(0)\reactos"
	FreeldrIniLst.AddItems "Options=/MININT /RDPATH="+FlashFileList.Item(i)+" /RDEXPORTASCD"
	FreeldrIniLst.AddItems "===="
	
next i

'call  AddClrString ("1048:FreeldrIniLst.Text="+(FreeldrIniLst.Text), 111, LogEdit)

FreeldrIniLst.SaveToFile FlashDrive$+"\"+"freeldr.ini"



FlashFileList.update


brem 0
======================

[LiveCD0414VG]
BootType=Windows2003
SystemPath=ramdisk(0)\reactos
Options=/MININT /RDPATH=livecd0414VG.iso /RDEXPORTASCD

======================
erem


End Sub 

'!******************************************
Sub MainTabChange

SELECT CASE MainTab.TabIndex
CASE 0
	'ISOPath$=StartPath$+"ISO\"  ' archive
	ISOList.Directory=ISOPath$
	
CASE 1
	ISOList.Directory=StartPath$+"archive\"  ' 
end select

ISOList.Update
End Sub 

'!******************************************
'Sub TimerCopyOver

'LabelFileInfo.Caption=Time$

'LabelFileInfo.Caption=str$(FileSize(FlashDrive$+"\"+ISOFN$))
'End Sub 

'!******************************************
Sub DelFiles

End Sub 

'!******************************************
Sub MoveFiles (Sender as QButtonXP)

'! copy selected ISO to flash ---------------
'call  AddClrString ("1347:ISOList.SelCount="+str$(ISOList.SelCount), 555667, LogEdit)

if ISOList.SelCount=0 then
	ShowMessage ("1350: Select ISO files to copy.")
	exit sub
end if


for i618=0 to ISOList.ItemCount-1
	zzz=(ISOList.Selected(i618))
	ISOFN$=ISOList.Item(i618)
	
	if zzz>0 then
		fext$=StripFileExt (ISOList.Item(i618))
		ISOFN$=ISOList.Item(i618)
		'call  AddClrString ("736:ISOFN$="+(ISOFN$), clred, LogEdit)
		
		if Sender.tag=473 then
			SHFile.Move (ISOList.Directory+"\"+ISOFN$,arcPAth$+"\"+ISOFN$)
		elseif Sender.tag=455 then
			kill (ISOList.Directory+"\"+ISOFN$)
		end if
		
	end if
next i618

ISOList.Update


End Sub 
