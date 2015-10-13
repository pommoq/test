<%


set p = server.createObject("DypsRTF.Document")
p.Header = "hello"

p.WriteBlankLine
p.WriteBlankLine

'Set the font and alignment
p.setFont("arial")
p.alignjustify

p.writebold("Rich Text Format (RTF) Specification, version 1.6")
p.WriteBlankLine
p.WriteBlankLine

p.WriteText("Microsoft Corporation")
p.writeTab
p.WriteText("May 1999")

p.WriteText("Summary: ")
  
  
p.WriteText("The Rich Text Format (RTF) Specification "&_
"provides a format for text and graphics interchange that"&_
" can be used with different output devices, operating "&_
"environments, and operating systems. RTF uses the American "&_
"National Standards Institute (ANSI), PC-8, Macintosh, or "&_
"IBM PC character set to control the representation and "&_
"formatting of a document, both on the screen and in print."&_
" With the RTF Specification, documents created "&_
"under different operating systems and "&_
"with different software applications "&_
"can be transferred between those operating systems "&_
"and applications. (248 printed pages)")
 
p.AlignRight
p.WriteText("The Rich Text Format (RTF) Specification"&_
" is a method of encoding formatted text and graphics for "&_
"easy transfer between applications. Currently, users depend "&_
"on special translation software to move word-processing "&_
"documents between different MS-DOS?, Microsoft? Windows?,"&_
" OS/2, Macintosh?, and Power Macintosh? applications.")
 
  
p.AlignLeft 
p.WriteText("The RTF Specification provides a format "&_
"for text and graphics interchange that can be used with "&_
"different output devices, operating environments, and "&_
"operating systems. RTF uses the American National "&_
"Standards Institute (ANSI), PC-8, Macintosh, or IBM PC "&_
"character set to control the representation and formatting"&_
" of a document, both on the screen and in print. With the"&_
" RTF Specification, documents created under different "&_
"operating systems and with different software applications"&_
" can be transferred between those operating systems and "&_
"applications. RTF files created in Word 6.0 (and later) "&_
"for the Macintosh and Power Macintosh have a file"&_
" type of RTF.")
p.WriteBlankLine

p.writeItalic("This is italic")

p.WriteBlankLine


'Others fonts
p.setFont("times new roman")
p.writetext("words in times")
p.WriteBlankLine


'p.setFont("courier")
'p.writetext("words in courier")
'p.WriteBlankLine
	
	
'p.setFont("arial")
'p.writetext("words in arial")
'p.WriteBlankLine

'Save the file on disk
'be sure to have write permission
'before!  
p.SaveDoc(server.mapPath("myDoc.rtf"))
set p = nothing
  

  %>


