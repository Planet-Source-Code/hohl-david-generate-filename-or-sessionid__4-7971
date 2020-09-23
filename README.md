<div align="center">

## Generate Filename or SessionID


</div>

### Description

Generate Filename with Randomstring, Date and time.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hohl David](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hohl-david.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__4-35.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hohl-david-generate-filename-or-sessionid__4-7971/archive/master.zip)





### Source Code

```
function generateFileName()
	dim counter
	Dim TempLongID
	For Counter = 1 to 10
	 Randomize Timer / RND
	 TempLongID = TempLongID & Chr((ASC("z") - ASC("a")) * RND + ASC("a"))
	Next
	dim temp
	dim datetime
	datetime = now()
	temp = day(datetime) & month(datetime) & year(datetime) & hour(datetime) & minute(datetime) & second(datetime)
	generateFileName = temp & TempLongID
end function
```

