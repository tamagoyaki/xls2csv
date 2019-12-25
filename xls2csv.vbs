'
' Convert XLSX to CSV
'
' USAGE
'
'   $ wscript.exe xls2csv.vbs hoge.xlsx hoge.csv
'
'
' REFERENCE
'
'   stackoverflow.com/questions/1858195/convert-xls-to-csv-on-command-Line
'



'
' check arguments
'
argcount = wscript.arguments.count

if argcount < 2 Or argcount > 3 Then
   WScript.Echo "Usage: xls2csv hoge.xls hoge.csv [sheet number]" & vbcrlf & "sheet number = 1...n"
   
    Wscript.Quit
End If

argsrc = Wscript.Arguments.Item(0)
argdst = Wscript.Arguments.Item(1)

If argcount = 2 Then
   sheetnum = 1
Elseif argcount = 3 then
   sheetnum = Int(Wscript.Arguments.Item(2))
End If


'
' to accept filename with relative path 
'
Set objFSO = CreateObject("Scripting.FileSystemObject")
srcfile = objFSO.GetAbsolutePathName(argsrc)
destfile = objFSO.GetAbsolutePathName(argdst)


'
' https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat
'
XLCSV = 6


'
' open xls
'
set oExcel = CreateObject("Excel.Application")
set oBook = oExcel.Workbooks.Open(srcfile)

'
' check if it has multiple sheets.
'
sheetcnt = obook.sheets.count

if sheetcnt <> 1 And sheetnum = Empty then
    wscript.echo  argsrc & " has " & sheetcnt & " sheets" & vbcrlf & "please speciry sheet number"
    wscript.quit
end If

Set xsheet = oexcel.sheets(sheetnum)
xsheet.saveas destfile, XLCSV
'oBook.SaveAs destfile, XLCSV

oBook.Close False
oExcel.Quit
