Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

'user input
dTimer=InputBox("Enter timer interval in minutes","Set Timer") 

'error handling
do until IsNumeric(dTimer)=True
  dTimer=InputBox("Invalid Entry" & vbnewline & vbnewline & "Enter timer interval in minutes","Set Timer") 
loop

'sleep, then notify user if they choose yes
WScript.Sleep dTimer*60*1000 'convert from minutes to milliseconds
t=MsgBox("Recharge."& vbnewline & vbnewline & vbnewline & vbnewline & vbnewline & "Info about the timer?", vbYesNo)
if t=6 then 'if yes
   MsgBox("The timer was " & dTimer & " minutes long")
end if

