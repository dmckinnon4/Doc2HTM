VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateHTML()
' create HTML from current document
Dim myScriptResult As String
Dim filePath As String
Dim TheOS As String

currentOS = Application.System.OperatingSystem

With ActiveDocument
    filePath = .FullName
    ' test if it is a web address
    pos = InStr(filePath, "http")
    If (pos > 0) Then
        MsgBox "Save files downloaded from internet to local disk before converting to HTML.", , "Error"
    Else
        .Save  'save document to file so pandoc uses latest version
        If (currentOS = "Macintosh") Then
            ' Use applescript to run Python program
            ' Applescript file is located in: "/Users/david/Library/Application Scripts/com.microsoft.Word/pandocShellScrpt.scpt"
            ' mac locations
                ' python = "/Applications/anaconda/envs/py36/bin/python3.6"
                ' doc2HTM = " ~/Documents/Python/Doc2HTM/src/pandocHTML.py "
'comment out on Windows
                'myScriptResult = AppleScriptTask("pandocShellScrpt.scpt", "scriptHandler", filePath)
            Debug.Print myScriptResult
        Else ' default to Windows
            ' Windows locations
            python = "C:\Users\dmckinnon\AppData\Local\Continuum\anaconda3\python.exe "
            doc2HTM = " ""C:\Users\dmckinnon\Documents\9 Software\Python\Doc2HTM\src\pandocHTML.py"" "
            RetVal = Shell(python & doc2HTM & filePath)
            Debug.Print RetVal
        End If
    End If
        
    Debug.Print filePath
    Debug.Print pos

End With

End Sub  'CreateHTML

Sub AdjustImageHeights()
' This subroutine adjusts image heights within document based on codes in the Alternative Text field
' Default height is 2in = 144
' 2.5in = 180
' 1.5in = 108
' Debug.Print InchesToPoints(2.5)

Dim oShp As Shape
Dim oILShp As InlineShape
Dim aspectRatio As Double
Dim newWidth As Long

With ActiveDocument
    For Each oILShp In .InlineShapes
        With oILShp
            .LockAspectRatio = True
            
            If .AlternativeText = "2.5" Then
                .Height = 180
            ElseIf .AlternativeText = "1.5" Then
                .Height = 108
            Else
                .Height = 144
            End If
            
        End With
    Next
    .Save ' save document to file
End With
End Sub 'AdjustImageHeights


