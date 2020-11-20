Attribute VB_Name = "Functions"
'CLIPBOARD AND ENHANCED METAFILE FUNCTIONS
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CopyEnhMetaFileA Lib "gdi32" (ByVal hENHSrc As Long, ByVal lpszFile As String) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long


'EXPORT CHART IN EMF VECTORIAL FORMAT
Public Function ExportEMF(strFileName As String) As Boolean

    Const CF_ENHMETAFILE    As Long = 14
    Dim ReturnValue         As Long
    
    OpenClipboard 0
    ReturnValue = CopyEnhMetaFileA(GetClipboardData(CF_ENHMETAFILE), strFileName)
    EmptyClipboard
    CloseClipboard
    DeleteEnhMetaFile ReturnValue 'Release resources to it eg You can now delete it if required or write over it. This is a MUST
    ExportEMF = (ReturnValue <> 0)

End Function


'HIGHLIGHT EVERY REFERENCE TO A CHART IN THE TEXT
Public Function HighlightText(Paragraph As String) As String

    Dim IndexPosition       As Integer
    Dim BracePosition       As Integer
    Dim StartChar           As Integer
    Dim Flag                As Integer
    Dim Chart               As Variant
    Dim ChartIndex          As Variant
    
    StartChar = 1
    Flag = 1
    
    Do While Flag = 1
        IndexPosition = InStr(StartChar, Paragraph, "Chart", vbTextCompare)
        
        If IndexPosition <> 0 Then
        
                BracePosition = InStr(IndexPosition, Paragraph, ")", vbTextCompare)
                Chart = Mid(Paragraph, IndexPosition, BracePosition - IndexPosition)
                ChartIndex = Mid(Paragraph, IndexPosition + 6, BracePosition - IndexPosition - 6)
                
                If ChartIndex <> "1" And ChartIndex <> "2" And ChartIndex <> "3" And ChartIndex <> "4" And ChartIndex <> "5" And ChartIndex <> "6" Then
                
                    Paragraph = Left(Paragraph, IndexPosition - 1) & "<span style=""background-color: #FFFF00""><b>" & Chart _
                    & "</b></span>" & Right(Paragraph, Len(Paragraph) - BracePosition + 1)
                    StartChar = IndexPosition + Len(Mid(Paragraph, IndexPosition, BracePosition - IndexPosition)) + 55
                    
                Else
                
                    Paragraph = Left(Paragraph, IndexPosition - 1) & "<b>" & Chart & "</b>" & Right(Paragraph, Len(Paragraph) - BracePosition + 1)
                    StartChar = IndexPosition + Len(Mid(Paragraph, IndexPosition, BracePosition - IndexPosition)) + 55
                    
                End If
            
        Else
        
            Flag = 0
            
        End If
        
    Loop
    'Table [A-Z][0-9]
    
    HighlightText = Paragraph
    
End Function


'IDENTIFY CHART/TABLE INDEX
Public Function ChartIndex(Str As String) As String
    
    Dim TrimString          As String
    Dim ReturnString        As String
    Dim i                   As Integer
    
    TrimString = Trim(Str)
    ReturnString = ""
    
    For i = 1 To Len(TrimString)
        If Mid(TrimString, i, 1) <> ":" Then
            ReturnString = ReturnString + Mid(TrimString, i, 1)
        End If
    Next
    
    ChartIndex = ReturnString
    
End Function


'CONVERT WORD RANGE TO STRING
Function RangeToString(ByVal wdrng As Word.Range) As String
    
    RangeToString = ""
    RangeToString = wdrng.Text
    
End Function


'REMOVE SPECIAL CHARACTERS FROM STRING
Function removeSpecChar(Str As String) As String

    Dim stringWithoutSpecChars, character, charPos
    
    charPos = 1
    
    Do
        character = Mid(Str, charPos, 1)
        ' If the character belongs to A-Z, a-z, 0-9 or space then store it in another string
        If character Like "[A-Za-z0-9 (),;.€-]" Then
            stringWithoutSpecChars = stringWithoutSpecChars & character
            Else
            '  Do nothing
            End If
            charPos = charPos + 1
    Loop Until charPos > Len(Str)
    
    removeSpecChar = stringWithoutSpecChars
    
End Function
