<div align="center">

## Solution for Windows Vista


</div>

### Description

Use of the API SendInput instead of SendKeys!

This will prevent the error of access denied in the Windows Vista.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rodrigo da Silva Brito](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rodrigo-da-silva-brito.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rodrigo-da-silva-brito-solution-for-windows-vista__1-65665/archive/master.zip)





### Source Code


<p><font face="Tahoma" style="font-size: 8pt"><u><b>Use of the API SendInput instead
of SendKeys<br>
This will prevent the error of access denied in the Windows Vista.</b></u><br>
&nbsp;</font></p>
<p><font face="Tahoma" style="font-size: 8pt">They forgive me but my English is
not very good! I am Brazilian!<br>
<br>
Copy and paste this code in a module of vb! It will go to substitute the
SendKeys standard of the VB! <br>
&nbsp;</font></p>
<p><font face="Tahoma" style="font-size: 8pt"><font color="#000080">Option
Explicit</font><br>
<br>
<font color="#000080">Private Const</font> KEYEVENTF_KEYUP = &amp;H2<br>
<font color="#000080">Private Const</font> INPUT_KEYBOARD = 1<br>
<br>
<font color="#000080">Private Type</font> KEYBDINPUT<br>
wVk <font color="#000080">As Integer</font><br>
wScan <font color="#000080">As Integer</font><br>
dwFlags <font color="#000080">As Long</font><br>
time <font color="#000080">As Long</font><br>
dwExtraInfo <font color="#000080">As Long</font><br>
<font color="#000080">End Type</font><br>
<br>
<font color="#000080">Private Type</font> GENERALINPUT<br>
dwType As Long<br>
xi(0 To 23) As Byte<br>
<font color="#000080">End Type</font><br>
<br>
<font color="#000080">Private Declare Function</font> SendInput
<font color="#000080">Lib</font> &quot;user32.dll&quot; (<font color="#000080">ByVal
</font>nInputs <font color="#000080">As Long</font>, pInputs
<font color="#000080">As GENERALINPUT</font>, <font color="#000080">ByVal</font>
cbSize <font color="#000080">As Long</font>) <font color="#000080">As Long</font><br>
<font color="#000080">Private Declare Sub</font> CopyMemory
<font color="#000080">Lib</font> &quot;kernel32&quot; <font color="#000080">Alias</font> &quot;RtlMoveMemory&quot;
(pDst <font color="#000080">As Any</font>, pSrc <font color="#000080">As Any</font>,
<font color="#000080">ByVa</font>l ByteLen <font color="#000080">As Long</font>)<br>
<br>
<font color="#000080">Public Function</font> SendKeysA(<font color="#000080">ByVal
</font>vKey <font color="#000080">As Integer</font>, <font color="#000080">
Optional</font> booDown <font color="#000080">As Boolean</font> =
<font color="#000080">False</font>)<br>
<font color="#000080">Dim </font>GInput(0) <font color="#000080">As GENERALINPUT</font><br>
<font color="#000080">Dim</font> KInput <font color="#000080">As KEYBDINPUT</font><br>
KInput.wVk = vKey<br>
<font color="#000080">If Not</font> booDown <font color="#000080">Then</font><br>
&nbsp;&nbsp;&nbsp; KInput.dwFlags = KEYEVENTF_KEYUP<br>
<font color="#000080">End If</font><br>
GInput(0).dwType = INPUT_KEYBOARD<br>
CopyMemory GInput(0).xi(0), KInput, Len(KInput)<br>
<font color="#000080">Call</font> SendInput(1, GInput(0), Len(GInput(0)))<br>
<font color="#000080">End Function</font><br>
<br>
Using in form! <br>
Example: Instead of SendKeys (“{TAB}”) you it will use SendKeys vbKeyTab, True<br>
&nbsp;</font></p>
<p><font face="Tahoma" style="font-size: 8pt"><b><u>Simulation:<br>
Before!</u><br>
</b><font color="#000080">Private Sub</font> Form_KeyPress(KeyAscii
<font color="#000080">As</font> <font color="#000080">Integer</font>)<br>
<font color="#000080">If</font> KeyAscii = vbKeyReturn <font color="#000080">
Then</font><br>
&nbsp;&nbsp; SendKeys (&quot;{TAB}&quot;)<br>
&nbsp;&nbsp; KeyAscii = 0<br>
<font color="#000080">End If</font><br>
<font color="#000080">End Sub</font></font></p>
<p><font face="Tahoma" style="font-size: 8pt"><br>
<b><u>Later!</u><br>
</b><font color="#000080">Private Sub</font> Form_KeyPress(KeyAscii
<font color="#000080">As Integer</font>)<br>
<font color="#000080">If</font> KeyAscii = vbKeyReturn <font color="#000080">
Then</font><br>
&nbsp;&nbsp; SendKeys vbKeyTab, <font color="#000080">True</font><br>
&nbsp;&nbsp; KeyAscii = 0<br>
<font color="#000080">End If<br>
End Sub</font><br>
&nbsp;</font></p>
<p><font face="Tahoma" style="font-size: 8pt">Abraços.<br>
Solução para o Windows Vista! Envio de teclas através da API SendInput!</font></p>

