Attribute VB_Name = "CreateEmail"
Option Explicit

Public outApp                   As Outlook.Application
Public outMail                  As Outlook.MailItem

Public wdDoc                    As Word.Document
Public wdPar                    As Word.Paragraph
Public wdTab                    As Word.Table
Public wdCell                   As Word.Cell
Public wdShape                  As Word.InlineShape
Public wdFoot                   As Word.Footnote
Public wdChar                   As Variant

Sub CreateEmail()
        
'#######################################################################################################
'#  PREPARE SUBJECT AND CONTENT OF THE EMAIL
'#######################################################################################################

    Dim Month                   As Variant
    Dim IndexPosition           As Integer

    Set wdDoc = ActiveDocument
    Set outApp = New Outlook.Application
    Set outMail = outApp.CreateItem(olMailItem)
    
    ' EXTRACT THE CORRESPONDING MONTH FROM THE TEXT!
    For Each wdPar In wdDoc.Sections(2).Range.Paragraphs
        IndexPosition = InStr(1, wdPar, "PRELIMINARY ASSESSMENT OF MONETARY DATA", vbTextCompare)
        If IndexPosition <> 0 Then
            Month = Mid(wdPar.Range.Text, 43)
            Month = Left(Month, Len(Month) - 2)
            Exit For
        End If
    Next

    outMail.To = "Ueue"
    outMail.CC = "Ueue"
    outMail.Subject = "Preliminary assessment of monetary data - " & Month & " (ECB-confidential)"
    
'#######################################################################################################
'#  OPEN FILE FOR HEADER
'#######################################################################################################
    
    Dim strFileContent          As String
    Dim iFile                   As Integer: iFile = FreeFile
    
    Open "D:\header1.txt" For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    outMail.HTMLBody = strFileContent & Format(Date, "dd mmmm yyyy")
    
    outMail.HTMLBody = outMail.HTMLBody & "</span><span style='font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman"";color:white;mso-font-kerning:18.0pt'><br>" _
          & "<br></span><span style='font-size:12.0pt;mso-bidi-font-size:18.0pt; font-family:""Arial"",""sans-serif"";mso-fareast-font-family:""Times New Roman"";" _
          & "color:white;mso-font-kerning:18.0pt'>Preliminary assessment of monetary data - " & Month
    
    
    Open "D:\header2.txt" For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    outMail.HTMLBody = outMail.HTMLBody & strFileContent
    
'#######################################################################################################
'#  LOOP THROUGH THE MONO PARAGARPHS AND COPY THEM IN THE EMAIL
'#######################################################################################################
    
    Dim prevChar                As Variant
    Dim Flag                    As Integer
    Dim StrBold                 As String
    Dim StrNonBold              As String
    
    For Each wdPar In wdDoc.Sections(2).Range.Paragraphs
        If wdPar.Range.Words.Count > 50 And wdPar.Range.Tables.Count = 0 Then
            Flag = 1
            For Each wdChar In wdPar.Range.Characters
                If wdChar.Bold = True And Flag = 1 Then
                    StrBold = StrBold & wdChar
                    prevChar = wdChar
                Else
                    Flag = 0
                    StrNonBold = StrNonBold & wdChar
                End If
            Next

            outMail.HTMLBody = outMail.HTMLBody & "<p class=MsoNormal style='text-align:justify'><b><span style='font-size:10.0pt;font-family:""Arial"",""sans-serif""'>"
            outMail.HTMLBody = outMail.HTMLBody & StrBold
            outMail.HTMLBody = outMail.HTMLBody & "</span></b><span style='font-size:10.0pt;font-family:""Arial"",""sans-serif"";mso-bidi-font-weight:bold'>"
            outMail.HTMLBody = outMail.HTMLBody & HighlightText(StrNonBold) & "</p><br>"
            
            StrBold = ""
            StrNonBold = ""
        End If
    Next
    
'#######################################################################################################
'#  OPEN FILE FOR CHARTS
'#######################################################################################################
    
    Dim ShapeIndex              As String
    Dim ShapeHeight             As Variant
    Dim ShapeWidth              As Variant
    Dim Titles(6)               As Variant
    Dim Units(6)                As Variant
    Dim Footnotes(6)            As Variant
    Dim k                       As Integer
    Dim lhs                     As Integer
    Dim rhs                     As Integer
    

    For Each wdTab In wdDoc.Tables
    
        For Each wdCell In wdTab.Range.Cells
        
            If wdCell.Range.Find.Execute("Chart") = True Then
            
                ShapeIndex = ChartIndex(Mid(wdCell.Range.Text, 7, 3))
                
                If ShapeIndex = "1" Or ShapeIndex = "2" Or ShapeIndex = "3" Or ShapeIndex = "4" Or ShapeIndex = "5" Or ShapeIndex = "6" Then
                    
                    'PARSE PARAGRAPH INTO TITLE AND UNITS
                    For Each wdPar In wdCell.Range.Paragraphs
                    
                        For Each wdChar In wdPar.Range.Characters
                        
                            If wdChar.Bold = True Then
                            
                                StrBold = StrBold & wdChar
                                
                            Else
                            
                                StrNonBold = StrNonBold & wdChar
                                
                            End If
                            
                        Next
                        
                    Next
                    
                    Titles(ShapeIndex) = removeSpecChar(StrBold)
                    Units(ShapeIndex) = removeSpecChar(StrNonBold)
                    
                    StrBold = ""
                    StrNonBold = ""
                    
                    
                    'DOWNLOAD ALL CHARTS
                    Set wdShape = wdTab.Cell(wdCell.RowIndex + 1, wdCell.ColumnIndex).Range.InlineShapes(1)
                    
                    wdShape.LockAspectRatio = msoFalse
                    ShapeHeight = wdShape.Height
                    ShapeWidth = wdShape.Width
                    wdShape.Height = CentimetersToPoints(13.95)
                    wdShape.Width = CentimetersToPoints(19.9)
                    
                    wdShape.Range.Select
                    wdShape.Range.Copy
                    ExportEMF ("D:\MoNoCharts\Chart_" & ShapeIndex & ".emf")
                    
                    wdShape.Height = ShapeHeight
                    wdShape.Width = ShapeWidth
                    
                    
                    'PARSE PARAGRAPH INTO TITLE AND UNITS
                    For Each wdPar In wdTab.Cell(wdCell.RowIndex + 2, wdCell.ColumnIndex).Range.Paragraphs
                        StrNonBold = StrNonBold & removeSpecChar(RangeToString(wdPar.Range)) & "<br>"
                    Next
                    
                    Footnotes(ShapeIndex) = StrNonBold
                    StrNonBold = ""
                    
                End If
                
            End If
            
        Next
        
    Next

    Open "D:\charts1.txt" For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    outMail.HTMLBody = outMail.HTMLBody & strFileContent
    
    ' OPEN TABLE
    outMail.HTMLBody = outMail.HTMLBody & "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=""100%"" style='width:100.0%;border-collapse:collapse;mso-yfti-tbllook:" _
    & "1184;mso-padding-alt:0cm 0cm 0cm 0cm;border-spacing:0'>"
    
    
    ' INITIALIZE THE INDEXES FOR THE TABLE ITERATION
    lhs = 1
    rhs = 2
    
    For k = 1 To 3
    
        ' CHART TITLE AND UNITS ROW
        outMail.HTMLBody = outMail.HTMLBody & "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>"
    
        ' TITLE CELL # 1
            outMail.HTMLBody = outMail.HTMLBody & "<td width=202 valign=top style='width:202.1pt;background:#003299; padding:0.1cm 0.1cm 0.1cm 0.1cm'>" _
            & "<p class=MsoNormal><b><span style='font-size:10.0pt;font-family:""Arial"",""sans-serif"";color:white'>"
            outMail.HTMLBody = outMail.HTMLBody & Titles(lhs) & "</span></b></p>"
            
            'UNIT
            outMail.HTMLBody = outMail.HTMLBody & "<p class=MsoNormal style='text-align:justify'><i><span style='font-size:9.0pt;font-family:""Arial"",""sans-serif"";color:white'>"
            outMail.HTMLBody = outMail.HTMLBody & Units(lhs) & "</span></i></p></td>"
            
            ' TITLE CELL # 2
            outMail.HTMLBody = outMail.HTMLBody & "<td width=202 valign=top style='width:202.1pt;background:#003299; padding:0.1cm 0.1cm 0.1cm 0.1cm'>" _
            & "<p class=MsoNormal><b><span style='font-size:10.0pt;font-family:""Arial"",""sans-serif"";color:white'>"
            outMail.HTMLBody = outMail.HTMLBody & Titles(rhs) & "</span></b></p>"
            
            'UNIT
            outMail.HTMLBody = outMail.HTMLBody & "<p class=MsoNormal style='text-align:justify'><i><span style='font-size:9.0pt;font-family:""Arial"",""sans-serif"";color:white'>"
            outMail.HTMLBody = outMail.HTMLBody & Units(rhs) & "</span></i></p></td></tr>"
        
        
        'CAHRT ROW
        outMail.HTMLBody = outMail.HTMLBody & "<tr style='mso-yfti-irow:1'>"
                
            ' CHART CELL # 1
            outMail.HTMLBody = outMail.HTMLBody & "<td width=202 valign=top style='width:202.1pt;background:white; mso-background-themecolor:background1;padding:0cm 0cm 0cm 0cm'>"
            outMail.HTMLBody = outMail.HTMLBody & "<img width=290 height=200 src=""D:\MoNoCharts\Chart_" & lhs & ".emf""></td>"
            
            ' CHART CELL # 2
            outMail.HTMLBody = outMail.HTMLBody & "<td width=202 valign=top style='width:202.1pt;background:white; mso-background-themecolor:background1;padding:0cm 0cm 0cm 0cm'>"
            outMail.HTMLBody = outMail.HTMLBody & "<img width=290 height=200 src=""D:\MoNoCharts\Chart_" & rhs & ".emf""></td></tr>"
        
        
        'FOOTNOTES ROW
        outMail.HTMLBody = outMail.HTMLBody & "<tr style='mso-yfti-irow:1'>"
               
            ' FOOTNOTE CELL # 1
            outMail.HTMLBody = outMail.HTMLBody & "<td width=202 valign=top style='width:202.1pt;padding:0cm 0cm 0cm 0cm'><p class=MsoNormal><span style='font-size:7.0pt;font-family:""Arial"",""sans-serif""'>"
            outMail.HTMLBody = outMail.HTMLBody & Footnotes(lhs) & "</span></p></td>"
            
            ' FOOTNOTE CELL # 2
            outMail.HTMLBody = outMail.HTMLBody & "<td width=202 valign=top style='width:202.1pt;padding:0cm 0cm 0cm 0cm'><p class=MsoNormal><span style='font-size:7.0pt;font-family:""Arial"",""sans-serif""'>"
            outMail.HTMLBody = outMail.HTMLBody & Footnotes(rhs) & "</span></p></td></tr>"
        
        
        lhs = lhs + 2
        rhs = rhs + 2
    
    Next
       
      
    Open "D:\charts2.txt" For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    outMail.HTMLBody = outMail.HTMLBody & strFileContent
        
    
'#######################################################################################################
'#  OPEN FILE FOR FOOTER
'#######################################################################################################
    
    
    Open "D:\footer1.txt" For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    outMail.HTMLBody = outMail.HTMLBody & strFileContent

    
    outMail.HTMLBody = outMail.HTMLBody & wdDoc.Footnotes(1).Range.Text


    Open "D:\footer2.txt" For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    outMail.HTMLBody = outMail.HTMLBody & strFileContent


    outMail.Display

End Sub
