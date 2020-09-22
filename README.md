<div align="center">

## Is Form Loaded check without causing form to load \(EASY\!\)


</div>

### Description

Check to see if a form is loaded without actually loading the form, or access public form variables without causing the form to load
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Gallant](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-gallant.md)
**Level**          |Intermediate
**User Rating**    |4.7 (61 globes from 13 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-gallant-is-form-loaded-check-without-causing-form-to-load-easy__1-52415/archive/master.zip)





### Source Code

<p><BR>
 <BR>
 Ever want to check to see if a form is loaded before you try to access it?<br>
 The only way I know of (other than this way) is to loop through the form collection...
 a rather large pain in the rear.<br>
 The trick is to create a new form property. <BR>
 <BR>
 Add the following code to any form:</p>
<blockquote>
 <p> <BR>
  <font color="#006699" size="-1">Option Explicit</font><font size="-1"><br>
  <br>
  <font color="#009900">' Create a new property variable</font><br>
  <font color="#006699">Dim</font></font> <font size="-1"><strong>m_bLoaded</strong>
  <font color="#006699">As Boolean</font><br>
  <font color="#009900">' get the value of the new property</font><br>
  <font color="#006699">Public Property Get</font></font> <font size="-1"><strong>Loaded()</strong>
  <font color="#006699">As Boolean</font></font></p>
 <blockquote>
  <p><font size="-1"> Loaded = m_bLoaded</font></p>
 </blockquote>
 <p> <font color="#006699" size="-1">End Property</font></p>
 <p><font color="#009900" size="-1"> ' set the value of the new property</font><font size="-1"><br>
  <font color="#006699">Public Property Let</font></font> <font size="-1"><strong>Loaded</strong>(<font color="#006699">ByVal</font>
  bLoaded <font color="#006699">As Boolean</font>)</font></p>
 <blockquote>
  <p><font size="-1"> m_bLoaded = bLoaded</font></p>
 </blockquote>
 <p> <font color="#006699" size="-1">End Property</font></p>
 <p><font size="-1"><br>
  <font color="#006699">Private Sub</font></font> <font size="-1"><strong>Form_Load()</strong></font></p>
 <blockquote>
  <p> <font color="#009900" size="-1">' set the loaded property to true</font><font size="-1"><br>
   Me.Loaded = <font color="#006699">True</font></font></p>
 </blockquote>
 <p> <font color="#006699" size="-1">End Sub</font></p>
 <p><font color="#006699" size="-1">Private Sub</font><font size="-1"> <strong>Form_Unload</strong>(Cancel
  <font color="#006699">As Integer</font>)</font></p>
 <blockquote>
  <p> <font color="#009900" size="-1">' set the loaded property to false</font><font size="-1"><br>
   Me.Loaded = <font color="#006699">False</font></font></p>
 </blockquote>
 <p> <font color="#006699" size="-1">End Sub</font> </p>
</blockquote>
<p>&nbsp;</p>
<p>Now, form any other form or module, you can do this (assume you are using the
 default form name)</p>
<p><font color="#006699" size="-1">If</font><font size="-1"> Form1.Loaded = <font color="#006699">True Then</font></font></p>
<blockquote>
 <p> <font color="#006699" size="-1">MsgBox</font> <font size="-1"><strong>&quot;Form is loaded&quot;</strong></font></p>
</blockquote>
<p> <font color="#006699" size="-1">Else</font></p>
<blockquote>
 <p> <font color="#006699" size="-1">MsgBox</font> <font size="-1"><strong>&quot;Form is not loaded&quot;</strong></font></p>
</blockquote>
<p> <font color="#006699" size="-1">End If</font></p>
<p><br>
 Accessing this property will not cause the form to load in the event that loaded
 is false.<br>
 However, if you make a single variable and make it public on the form, and try
 to access it, the form will load.<br>
 You can actually use this property method to retain any data and access it without
 reloading hte form.<br>
 I created a custom input box field in which the &quot;Return String&quot; is
 a custom property, like the loaded property.<br>
 then I just do this:</p>
<blockquote>
 <p> <font size="-1">Form1.show 1, me <font color="#009900">' (show my new form
  modal)</font><br>
  strInput <font color="#006699">=</font> Form1.strInput <font color="#009900">'(this
  will not cause the form to reload provided your property is called strInput!)</font></font></p>
</blockquote>
<p>give it a shot, let me know if you have any problems.<br>
</p>

