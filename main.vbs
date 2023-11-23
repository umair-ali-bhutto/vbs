dim count
dim userin
userin=Int(InputBox("number"))
set object = wscript.CreateObject("wscript.shell")

do
object.run "b.vbs"
count = count + 1
loop until count = userin