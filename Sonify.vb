Sub Sonify()


Dim fsT, mxlFile, fname, folName, DataDesc, curDate As String
Dim i As Integer
Dim rCells As Range

fname = InputBox(Prompt:="Name of file to send output to ", _
          Title:="SONIFY - Output File Name", Default:=Range("B4").Value)

If fname = "Enter file name" Or fname = vbNullString Then
    MsgBox ("No file name provided - exiting Sonify " & fname)
    Exit Sub
End If
curDate = Range("B2").Value
folName = Range("B1").Value
DataDesc = fname
Do While DataDesc = vbNullString
    DataDesc = InputBox(Prompt:="Enter data set description", _
          Title:="SONIFY - Data description", Default:="Enter data set description")
Loop
mxlFile = folName + fname + ".xml"
' Create Stream object
Set fsT = CreateObject("ADODB.Stream")
' Specify stream type - we want To save text/string data.
fsT.Type = 2
' Specify charset For the source text data.
fsT.Charset = "utf-8"
' Open the stream And write binary data To the object
fsT.Open
    fsT.writetext "<?xml version=" & """" & "1.0" & """" & " standalone=" & """" & "no" & """" & "?>" & vbCrLf
    fsT.writetext "<score-partwise>" & vbCrLf
    fsT.writetext "    <movement-title> " & DataDesc & " </movement-title>" & vbCrLf
    fsT.writetext "    <identification>" & vbCrLf
    fsT.writetext "        <rights>Copyright © 2009 Neel Sanjay Patel</rights>" & vbCrLf
    fsT.writetext "        <encoding>" & vbCrLf
    fsT.writetext "            <software>Sonification Guru 0.1</software>" & vbCrLf
    fsT.writetext "            <encoding-date>" & curDate & "</encoding-date>" & vbCrLf
    fsT.writetext "        </encoding>" & vbCrLf
    fsT.writetext "    </identification>" & vbCrLf
    fsT.writetext "    <part-list>" & vbCrLf
    fsT.writetext "        <score-part id=" & """" & "P1" & """" & ">" & vbCrLf
    fsT.writetext "            <part-name>Voice</part-name>" & vbCrLf
    fsT.writetext "        </score-part>" & vbCrLf
    fsT.writetext "    </part-list>" & vbCrLf
    fsT.writetext "    <part id=" & """" & "P1" & """" & ">" & vbCrLf
    fsT.writetext "        <measure number=" & """" & "1" & """" & ">" & vbCrLf
    fsT.writetext "            <attributes>" & vbCrLf
    fsT.writetext "                <divisions>2</divisions>" & vbCrLf
    fsT.writetext "                <clef>" & vbCrLf
    fsT.writetext "                    <sign>G</sign>" & vbCrLf
    fsT.writetext "                    <line>2</line>" & vbCrLf
    fsT.writetext "                </clef>" & vbCrLf
    fsT.writetext "            </attributes>" & vbCrLf
    fsT.writetext "            <direction placement=" & """" & "above" & """" & ">" & vbCrLf
    fsT.writetext "                <direction-type>" & vbCrLf
    fsT.writetext "                    <words xml:lang=" & """" & "la" & """" & " relative-y=" & """" & "5" & """" & " relative-x=" & """" & "- 5" & """" & ">Angelus dicit:</words>" & vbCrLf
    fsT.writetext "                </direction-type>" & vbCrLf
    fsT.writetext "            </direction>" & vbCrLf
    fsT.writetext vbCrLf

    
    Set rCells = Range("N10:U209")
    
    
    For i = 1 To Range("N1").Value
        fsT.writetext "                     <note>" & vbCrLf
        fsT.writetext "                         <pitch>" & vbCrLf
        fsT.writetext "                             <step>" & rCells(i, 1) & "</step>" & vbCrLf
        fsT.writetext "                             <octave>" & rCells(i, 2) & "</octave>" & vbCrLf
        fsT.writetext "                         </pitch>" & vbCrLf
        fsT.writetext "                         <duration>" & rCells(i, 3) & "</duration>" & vbCrLf
        fsT.writetext "                         <voice>" & rCells(i, 5) & "</voice>" & vbCrLf
        fsT.writetext "                         <volume>" & rCells(i, 4) & "</volume>" & vbCrLf
        fsT.writetext "                         <type>half</type>" & vbCrLf
        fsT.writetext "                         <stem>up</stem>" & vbCrLf
        fsT.writetext "                         <notations>" & vbCrLf
        fsT.writetext "                             <slur type=" & """" & "start" & """" & " number=" & """" & "1" & """" & "/>" & vbCrLf
        fsT.writetext "                         </notations>" & vbCrLf
        fsT.writetext "                         <lyric number=" & """" & "1" & """" & ">" & vbCrLf
        fsT.writetext "                             <syllabic>single</syllabic>" & vbCrLf
        fsT.writetext "                             <text>Quem</text>" & vbCrLf
        fsT.writetext "                         </lyric>" & vbCrLf
        fsT.writetext "                     </note>    " & vbCrLf
    
        fsT.writetext "                     <note>" & vbCrLf
        fsT.writetext "                         <chord/>" & vbCrLf
        fsT.writetext "                         <pitch>" & vbCrLf
        fsT.writetext "                             <step>" & rCells(i, 7) & "</step>" & vbCrLf
        fsT.writetext "                             <octave>" & rCells(i, 2) & "</octave>" & vbCrLf
        fsT.writetext "                         </pitch>" & vbCrLf
        fsT.writetext "                         <duration>" & rCells(i, 3) - 1 & "</duration>" & vbCrLf
        fsT.writetext "                     </note>    "
    
        fsT.writetext "                     <note>" & vbCrLf
        fsT.writetext "                         <chord/>" & vbCrLf
        fsT.writetext "                         <pitch>" & vbCrLf
        fsT.writetext "                             <step>" & rCells(i, 8) & "</step>" & vbCrLf
        fsT.writetext "                             <octave>" & rCells(i, 2) & "</octave>" & vbCrLf
        fsT.writetext "                         </pitch>" & vbCrLf
        fsT.writetext "                         <duration>" & rCells(i, 3) - 2 & "</duration>" & vbCrLf
        fsT.writetext "                     </note>    " & vbCrLf
    
    Next i
    
    
fsT.writetext "         </measure>" & vbCrLf
fsT.writetext "     </part>" & vbCrLf
fsT.writetext "</score-partwise>" & vbCrLf

fsT.SaveToFile mxlFile, 2

End Sub
