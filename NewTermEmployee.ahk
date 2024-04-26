#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory
#SingleInstance Force

Start:
Gui, +AlwaysOnTop -SysMenu +Owner  ; +Owner avoids a taskbar button.
Gui, Add, Button, gNewEmployee default, &New Employee
Gui, Add, Button, gTerminatedEmployee, &Terminated Employee
Gui, Add, Button, gCancel, &Cancel
Gui, Show,, Employee Email Macro
Hotkey, esc, Cancel
return  ; End of auto-execute section. The script is idle until the user does something.



NewEmployee:
Gui, Destroy
Gui, +AlwaysOnTop -SysMenu +Owner  ; +Owner avoids a taskbar button.
Gui, Add, Text,, Employee Name:
Gui, Add, Text,, Employee Username:
Gui, Add, Text,, Employee ID Number:
Gui, Add, Text,, Employee Password:
Gui, Add, Text,, Employee Supervisor:
Gui, Add, Edit, vNewEmployeeName ym  ; The ym option starts a new column of controls.
Gui, Add, Edit, vNewEmployeeUsername
Gui, Add, Edit, vNewEmployeeID
Gui, Add, Edit, vNewEmployeePass

Gui, Add, Button, gbcopy yp xp+85, &Copy

Gui, Add, Edit, vNewEmployeeSV yp+25 xp-85
Gui, Add, Button, gRestart yp+30 xp+5, &Back
Gui, Add, Button, default xp+40, O&K  ; The label ButtonOK (if it exists) will be run when the button is pressed.
Gui, Add, Button, gCancel xp+30, Cance&l
Gui, Add, Button, gOpenAD xp-190, Open &AD
Gui, Show,  x250, New Employee
return  ; End of auto-execute section. The script is idle until the user does something.

ButtonOK:
Gui, Submit  ; Save the input from the user to each control's associated variable.
PACNumberBody := "Please add the following to our list of providers who are authorized to place long distance phone calls.%0D%0A%0D%0AName: " . NewEmployeeName . "%0D%0APAC/Provider#: " . NewEmployeeID . ""

CraigEmailBody := "I've created an account in AD for the following employee. Sending you this notification email to manually sync the email with Outlook 365:%0D%0A%0D%0AUsername: " . NewEmployeeUsername . "%0D%0AEmail Address: " . NewEmployeeUsername . "@adanta.org%0D%0APassword: " . NewEmployeePass . "%0D%0APlease let me know when I can notify the HR and the Supervisor."

SupervisorBody := "The Account for the new employee to login to Adanta Computer’s and Email has been setup:%0D%0A%0D%0AUsername: " . NewEmployeeUsername . "%0D%0AEmail Address: " . NewEmployeeUsername . "@adanta.org%0D%0APassword: " . NewEmployeePass . "%0D%0A%0D%0ALong distance phone call/fax%0D%0AProvider #: " . NewEmployeeID . "%0D%0APlease allow up to 72 business hours for long distance access."

EHRBody := "This employee’s account to access Adanta computers and email have been created%0D%0APlease allow up to 72 business hours for long distance access.%0D%0A%0D%0AUsername: " . NewEmployeeUsername . "%0D%0AEmail Address: " . NewEmployeeUsername . "@adanta.org%0D%0AProvider #: " . NewEmployeeID . "%0D%0AEmployee Name: " . NewEmployeeName . ""

Run mailto:care.inquiry@centurylink.com?Subject=Account number:  53691349&Body=%PACNumberBody%
Run mailto:chines@adanta.org?Subject=New Employee to Create O365 Account - %NewEmployeeName%&Body=%CraigEmailBody%
svmailto:= "mailto:" . NewEmployeeSV . "?Subject=New Employee - " . NewEmployeeName . "&Cc=hrdepartment@adanta.org&Body=" . SupervisorBody
clipboard:= svmailto
Run mailto:EHR Support?Subject=New Employee Account - %NewEmployeeName%&Body=%EHRBody%
ExitApp

TerminatedEmployee:
Gui, Destroy
Gui, +AlwaysOnTop -SysMenu +Owner  ; +Owner avoids a taskbar button.
Gui, Add, Text,, Employee Name:
Gui, Add, Text,, Employee ID Number:
Gui, Add, Button, gOpenAD, Open AD
Gui, Add, Edit, vTermdEmployeeName ym  ; The ym option starts a new column of controls.
Gui, Add, Edit, vTermdEmployeeID
Gui, Add, Button, gRestart yp+30 xp+15, &Back
Gui, Add, Button, gTermEmployeeEmail xp+40, O&K  ; The label ButtonOK (if it exists) will be run when the button is pressed.
Gui, Add, Button, gCancel xp+30, & Cancel
Gui, Show,, Terminated Employee
return  ; End of auto-execute section. The script is idle until the user does something.

TermEmployeeEmail:
Gui, Submit
PACNumberBody := "Please remove the following from our list of providers who are authorized to place long distance phone calls.%0D%0A%0D%0AName: " . TermdEmployeeName . "%0D%0APAC/Provider#: " . TermdEmployeeID . ""
Run mailto:care.inquiry@centurylink.com?Subject=Account number:  53691349&Body=%PACNumberBody%
ExitApp

Restart:
Gui, Destroy
Gosub, Start
return

Cancel:
GuiClose:
Gui, Destroy
ExitApp

OpenAD:
IfWinExist, Active Directory Users and Computers
{
	WinActivate ; Use the window found by IfWinNotExist.
	return
}
else
{
	RunAs, tagadmin, Jamaica10$130, adanta.org
	Run, cmd /c dsa.msc, %A_WinDir%\system32\,Hide
	WinWaitActive, Active Directory Users and Computers
	return
}
