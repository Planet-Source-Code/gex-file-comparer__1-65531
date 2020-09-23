Attribute VB_Name = "modComparer"
Option Explicit
Public Const BUFFERSIZE As Long = 1024 'For controlling the size of the buffer
'I havent actually tested changing this const but when comparing larger files I think a bigger buffer would be better

Public Function Compare(File1 As String, File2 As String, bGenReport As Boolean, ByRef outDiffArray() As Byte, ByRef bDiffInSize As Boolean) As Boolean
    'On Error Resume Next
    DoEvents
    
    Compare = True 'The function will return true by default unless code below sets it to false
    
    Dim a As Long 'Ignore this temp variable
    For a = 0 To 99
        outDiffArray(a) = 0 'Clear grafical difference array
    Next
    
    Dim buffer1(BUFFERSIZE - 1) As Byte 'buffers for reading files
    Dim buffer2(BUFFERSIZE - 1) As Byte

    Open File1 For Binary Access Read Lock Write As #1 'Open files
    Open File2 For Binary Access Read Lock Write As #2
    If bGenReport Then Open App.Path & "\Report.txt" For Append As #3 'If user wants report then open the file
    
    If Err.Number <> 0 Then 'if one file was not opened then close all and end
        Close #1
        Close #2
        If bGenReport Then Close #3
        Compare = False
        Exit Function
    End If
    
    If bGenReport Then Print #3, "Comparing " & Chr(34) & File1 & " to " & Chr(34) & File2 & vbCrLf
    
    Dim lBiggerFileSize As Long
    bDiffInSize = False
    
    If LOF(1) > LOF(2) Then
        lBiggerFileSize = LOF(1) 'If first file is bigger than second
        bDiffInSize = True 'used to notify the user files are different in size(returns to the calling funtion)
    ElseIf LOF(1) < LOF(2) Then
        lBiggerFileSize = LOF(2) 'If first file is smaller than second
        bDiffInSize = True 'used to notify the user files are different in size(returns to the calling funtion)
    Else
        lBiggerFileSize = LOF(1) 'if files are equal in sizes
    End If
    
    
    Dim lNumberOfLoops As Long
    lNumberOfLoops = lBiggerFileSize \ BUFFERSIZE 'calculate number of loops - how many times will buffer be filled untill the end of file
    
    Dim lFilePos As Long
    Dim lBuffPos As Long
    
    For lFilePos = 0 To lNumberOfLoops
        Get #1, lFilePos * BUFFERSIZE + 1, buffer1 'Read into the buffers
        Get #2, lFilePos * BUFFERSIZE + 1, buffer2
        
        For lBuffPos = 0 To BUFFERSIZE - 1
        
            If buffer1(lBuffPos) <> buffer2(lBuffPos) Then 'comparing the buffers
                If bGenReport Then Print #3, vbTab & "Byte #" & lFilePos * BUFFERSIZE + lBuffPos + 1 & " is different in files."
                outDiffArray(Round(1 / (lBiggerFileSize / 100 / (lFilePos * BUFFERSIZE + lBuffPos + 1)), 0)) = 1
                'The heavy formula inside arrays brackets calculates % using total file size and byte position
            End If
            
        Next
        
    Next
    
    If bGenReport Then 'SOme more reporting about percent differences
        Dim strTemp As String
        For a = 0 To 100
            If outDiffArray(a) = 1 Then strTemp = strTemp & ", " & a & "%"
        Next
        
        Print #3, vbTab & "Files are different approximately at positions: " & strTemp
        Print #3, vbCrLf & "Comparing files finished!" & vbCrLf
    End If
    
        Close #1 'Safe closing at end
        Close #2
        If bGenReport Then Close #3

End Function
