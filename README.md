<div align="center">

## ADO Search Engine Function


</div>

### Description

Hi Friends,This is my 2nd submission in PSC. I am sure you will like this ADO Search Engine Function. Vote for me if you like it. Feel free to give his to your friends. You Commends and suggestion are welcomed.Please vote, thank you! johnmanavalan@hotmail.com or Call Me 00-91-487-422353
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Manavalan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-manavalan.md)
**Level**          |Intermediate
**User Rating**    |3.2 (16 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-manavalan-ado-search-engine-function__4-7655/archive/master.zip)





### Source Code

<br>
<br>
<br>
<p>Public Function search(a)<br>
' Name: ADO Search Engine Function <br>
' By: JohnManavalan@hotmail.com<br>
' PLEASE take the time and vote for Me <br>
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=7655&lngWId=4<br>
search = "select * from search where"<br>
Dim words<br>
For Each words In Split(a, " ")<br>
If Len(Trim(words)) > 0 Then search = search & " keyword LIKE '%" & words & "%'
and "<br>
Next<br>
search = Mid(search, 1, Len(search) - 5)<br>
End Function</p>

