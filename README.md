<div align="center">

## \_Disable X button on form


</div>

### Description

_Disable X button on form
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Petko Petkov](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/petko-petkov.md)
**Level**          |Intermediate
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/petko-petkov-disable-x-button-on-form__1-48862/archive/master.zip)





### Source Code

<font color="#000099">Private Declare Function</font><font color="#FFFFFF"> </font>GetSystemMenu<font color="#000099">
Lib</font> &quot;user32&quot; (<font color="#000099">ByVal</font> hwnd <font color="#000099">As
Long</font>, <font color="#000099">ByVal </font>bRevert<font color="#000099">
As Long</font>) <font color="#000099">As Long</font><br>
<font color="#000099">Private Declare Function </font>GetMenuItemCount<font color="#000099">
Lib </font>&quot;user32&quot; (ByVal hMenu As Long) <font color="#000099">As Long</font><br>
<font color="#000099">Private Declare Function </font>RemoveMenu <font color="#000099">Lib
</font>&quot;user32&quot; (<font color="#000099">ByVal</font> hMenu <font color="#000099">As
Long</font>, <font color="#000099">ByVal</font> nPosition <font color="#000099">As
Long</font>, <font color="#000099">ByVal</font> wFlags <font color="#000099">As
Long</font>) <font color="#000099">As Long</font><br>
<font color="#000099">Private Declare Function </font>DrawMenuBar <font color="#000099">Lib</font>
&quot;user32&quot; (<font color="#000099">ByVal </font>hwnd <font color="#000099">As
Long</font>) <font color="#000099">As Long</font><br>
<font color="#000099">Private Const </font>MF_BYPOSITION = &amp;H400&amp;<br>
<font color="#000099">Private Const</font> MF_DISABLED = &amp;H2&amp;
<p><font color="#000099">Public Sub</font> DisableX(Frm <font color="#000099">As</font>
 Form)<br>
 <font color="#000099">Dim</font> hMenu <font color="#000099">As Long</font><br>
 <font color="#000099">Dim</font> nCount <font color="#000099">As Long</font><br>
 &nbsp;&nbsp;hMenu = GetSystemMenu(Frm.hwnd, 0)<br>
 &nbsp;&nbsp;nCount = GetMenuItemCount(hMenu)<br>
 <font color="#000099">&nbsp;&nbsp;Call </font>RemoveMenu(hMenu, nCount - 1,
 MF_DISABLED Or MF_BYPOSITION)<br>
 &nbsp;&nbsp;DrawMenuBar Frm.hwnd<br>
 <font color="#000099">End Sub</font></p>
<p><font color="#000099">Private Sub</font><font color="#006699"> </font>Command1_Click()<br>
 &nbsp;&nbsp;DisableX Me<br>
 <font color="#000099">End Sub</font></p>

