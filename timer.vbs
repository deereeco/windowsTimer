
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

dTimer=InputBox("Enter timer interval in minutes","Set Timer") 'minutes

do until IsNumeric(dTimer)=True
  dTimer=InputBox("Invalid Entry" & vbnewline & vbnewline & _ 
         "Enter timer interval in minutes","Set Timer") 'minutes
loop

if dTimer<>"" Then 'change () to brackets before run program
do
  'objShell.Run "walkNotification.vbs" 
  WScript.Sleep dTimer*60*1000 'convert from minutes to milliseconds
  t=MsgBox("Your "& dTimer & " Minute timer is done" & vbnewline & vbnewline &"" & vbnewline & vbnewline & "Restart Timer?", _
    vbYesNo, "Levantate el orto")
  if t=6 then 'if yes
       'continue loop
  else 'exit loop
       exit do
  end if
loop
end if