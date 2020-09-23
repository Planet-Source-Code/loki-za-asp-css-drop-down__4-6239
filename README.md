<div align="center">

## ASP/CSS Drop\-down


</div>

### Description

This code creates a CSS drop-down (small text/colours - whatever) that pulls values from a DB. It also validates the browser to check for CSS compatability. Can be nicely adapted for any site! *added (27/06/2K) - This assumes that you have a connection already and you'll have to customise a bit of the code for yourself. Any probs or if you want more of the same code, mail me (darrynb@digitalmall.com) - if you like the code then vote for me!*
 
### More Info
 
value from the drop down through a form

This is a very basic version, it can easily be made better. With a little work that is!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[LoKi\-ZA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/loki-za.md)
**Level**          |Beginner
**User Rating**    |4.3 (34 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/loki-za-asp-css-drop-down__4-6239/archive/master.zip)





### Source Code

```
<%'To check the browser capabilities this goes above everything
function checkbrowser
		set check = server.createObject("MSWC.BrowserType")
			if check.browser="IE" and check.version >= "4.0" then
				checkbrowser = "DropStyle"
			else
				checkbrowser = ""
			end if
	end function
'this goes in the head. very basic CSS
%>
<style>
<!--
.DropStyle {font-family: verdana;
    font-size: 8pt;
    font-weight: normal;
    color: black;
				width:144;
    }
-->
</style>
<%'This drop-down assumes that you have the conn 2 the database set up and you've build a RS.
%>
<table>
<tr>
 <td colspan="1" bgcolor=#66C299 align="center">
 <Font face="verdana" size=1>
 <select name="theme" class="<%=CheckBrowser%>">
			<option value="">Select a theme</option>
			<option value="">-----------------</option>
 				<%If not rs.eof Then
 					Do until rs.eof
 						Response.write("<option value='" & trim(rs("name")) & "'>" & trim(rs("name")) & "</option>")
 						rs.Movenext
 					Loop
 				End if%>
			</font></select>
	</td>
 </tr>
</table>
<%'any queries (or stuff-ups) pls mail me!
```

