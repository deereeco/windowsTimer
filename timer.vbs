
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

dTimer=InputBox("Enter timer interval in minutes","Set Timer") 'minutes

'error handling
do until IsNumeric(dTimer)=True
  dTimer=InputBox("Invalid Entry" & vbnewline & vbnewline & "Enter timer interval in minutes","Set Timer") 'minutes
end do


'mins = dTimer*60*1000 'convert from minutes to milliseconds
WScript.Sleep dTimer

t=MsgBox("Recharge." & vbnewline & vbnewline & vbnewline "Info about the timer?" , vbYesNo)

if t=6 then 'if yes, show how long the timer was
   MsgBox("The timer was " & dTimer & " long")
end if
   