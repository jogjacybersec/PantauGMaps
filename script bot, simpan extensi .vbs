dim vhotel, hotelname, clipboardText, nameofscreen, address1, address2, address3, town
stataddress1=0
stataddress2=0
stataddress3=0
stattown=0

vhotel = "https://www.google.com/search?q="  ' ambil dari address bar browser, masukkan nama hotel
hotelname = " "
nameofscreen=" "    ' ini nama tab firefox nya, contoh: Swissbell Banjar
address1=" "        ' sesuai yang di Google Business Profile
address2=" "         ' sesuai yang di Google Business Profile
address3=" "        ' sesuai yang di Google Business Profile
town= ""            ' sesuai yang di Google Business Profile

Set objShell = CreateObject("WScript.Shell")

' Buka Firefox dengan URL yang ditentukan
objShell.Run "firefox.exe " & vhotel
WScript.Sleep 10000 ' Tunggu 10 detik agar halaman sepenuhnya dimuat

' Pindahkan fokus ke jendela Firefox
If Not objShell.AppActivate("Hotel Tentrem Yogyakarta") Then
    WScript.Echo "Jendela Firefox tidak ditemukan!"
    WScript.Quit
Else
	objShell.AppActivate(nameofscreen)
End If

' Kirim perintah 24 TAB untuk klik profile 
For i = 1 To 24
    objShell.SendKeys "{TAB}"
    WScript.Sleep 200 
Next
objShell.SendKeys "{ENTER}"
WScript.Sleep 3000 

' 6 TAB untuk edit nama property
For i = 1 To 7
    objShell.SendKeys "{TAB}"
    WScript.Sleep 100 
Next
objShell.SendKeys "{ENTER}"
WScript.Sleep 3000  

' 2 TAB untuk check ABOUT
For i = 1 To 2
    objShell.SendKeys "{TAB}"
    WScript.Sleep 100 
Next
objShell.SendKeys "{ENTER}"
WScript.Sleep 500 

' Salin teks yang dipilih ke clipboard
objShell.SendKeys "^c"
WScript.Sleep 500 

' Jalankan PowerShell untuk membaca teks dari clipboard
Set objExec = objShell.Exec("powershell.exe -Command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()""")
clipboardText = Trim(objExec.StdOut.ReadAll())

clipboardText = CleanText(Trim(clipboardText))
hotelname = CleanText(Trim(hotelname))
 

If cstr(clipboardText) <> cstr(hotelname) Then
    objShell.SendKeys "{DEL}" ' Hapus teks yang ada
    WScript.Sleep 1000
    objShell.SendKeys hotelname
	 WScript.Sleep 1000
	objShell.SendKeys "{TAB}"
	WScript.Sleep 1000 
	objShell.SendKeys "{ENTER}"
	WScript.Sleep 2000 

End If
WScript.Sleep 2000
objShell.SendKeys "{ESC}"
WScript.Sleep 2000

'CEK ADDRESS1
objShell.SendKeys "{ENTER}"
WScript.Sleep 3000 
For i = 1 To 14
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500 
Next
objShell.SendKeys "{ENTER}"
For i = 1 To 5
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500 
Next
objShell.SendKeys "^c"
WScript.Sleep 800 
Set objExec = objShell.Exec("powershell.exe -Command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()""")
clipaddress1 = CleanText(Trim(objExec.StdOut.ReadAll()))
WScript.Sleep 800 
if clipaddress1<>address1 Then
  objShell.SendKeys address1
  WScript.Sleep 600
  stataddress1=1
end if

'CHECK ADDRESS2
objShell.SendKeys "{TAB}"
WScript.Sleep 800
objShell.SendKeys "^c"
WScript.Sleep 800 
Set objExec = objShell.Exec("powershell.exe -Command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()""")
clipaddress2 = CleanText(Trim(objExec.StdOut.ReadAll()))
WScript.Sleep 500 
if clipaddress2<>address2 Then
  objShell.SendKeys address2
  WScript.Sleep 600
  stataddress2=1
end if

'CHECK ADDRESS3
objShell.SendKeys "{TAB}"
WScript.Sleep 800
objShell.SendKeys "^c"
WScript.Sleep 800 
Set objExec = objShell.Exec("powershell.exe -Command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()""")
clipaddress3 = CleanText(Trim(objExec.StdOut.ReadAll()))
WScript.Sleep 500 
if clipaddress3<>address3 Then
  objShell.SendKeys address3
  WScript.Sleep 600
  stataddress3=1
end if


'CHECK TOWN
objShell.SendKeys "{TAB}"
WScript.Sleep 600
objShell.SendKeys "{TAB}"
WScript.Sleep 800
objShell.SendKeys "^c"
WScript.Sleep 800 
Set objExec = objShell.Exec("powershell.exe -Command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()""")
cliptown = CleanText(Trim(objExec.StdOut.ReadAll()))
WScript.Sleep 500 
if cliptown<>town Then
  objShell.SendKeys town
  WScript.Sleep 600
  stattown=1
end if
WScript.Sleep 500 

if stataddress1=1 or stataddress2=1 or stataddress3=1 or stattown=1 then 
	For i = 1 To 8
		objShell.SendKeys "{TAB}"
		WScript.Sleep 500 
	Next
	WScript.Sleep 800 
	objShell.SendKeys "{ENTER}"
end if
if stataddress1=0 and stataddress2=0 and stataddress3=0 and stattown=1 then 
	For i = 1 To 9
		objShell.SendKeys "{TAB}"
		WScript.Sleep 500 
	Next
	objShell.SendKeys "{ENTER}"
	WScript.Sleep 500 
	
end if
 
'CLOSE FIREFOX
objShell.AppActivate(nameofscreen)  
WScript.Sleep 600 
objShell.SendKeys "%{F4}" ' 

Function CleanText(text)
    ' Menghapus karakter CR (Carriage Return), LF (Line Feed), dan karakter spesial lainnya
    CleanText = Replace(text, vbCr, "")
    CleanText = Replace(CleanText, vbLf, "")
    CleanText = Replace(CleanText, vbTab, "")
    CleanText = Replace(CleanText, Chr(160), " ") ' Menghapus non-breaking space (spasi non-pemisah)
    CleanText = Trim(CleanText)
End Function
