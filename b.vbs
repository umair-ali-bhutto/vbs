dim NIC1, Nic, StrIP, CompName

Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1

   if Nic.IPEnabled then
        StrIP = Nic.IPAddress(0)

        Set WshNetwork = WScript.CreateObject("WScript.Network")
        CompName= WshNetwork.Computername
	do

        MsgBox "You Are HACKED ;) ;)"& vbNewLine _
            & "IP Address:  "&StrIP & vbNewLine _
            & "Computer Name:  "&CompName,4160,"Bro ERROR"
	loop
	wscript.quit
   End if
        
Next

