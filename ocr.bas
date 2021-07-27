Sub TestOCREngine()

    'Code By Kamal Bharakhda
    'kamal.9328093207@gmail.com
    
    'We are executing OCR through google's project tesseract.exe
    'first we have to download the package tesseract.exe from the github or from any resource
    'then we will use the Shell or command functionality to use tesseract
    'in following one line command, you will see the first part is of to locate where is the exe file
    'second part is of to locate the image file from which you want to extract information
    'and third and last part would be the file path of text file where you want to extract information with file name
    Call VBA.Shell(VBA.Chr(34) & "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" & VBA.Chr(34) & " " & _
    VBA.Chr(34) & "E:\DownloadedImages\2-5.png" & VBA.Chr(34) & " " & _
    VBA.Chr(34) & "E:\DownloadedImages\2-5" & VBA.Chr(34), vbMinimizedFocus)
    
    Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
    
    myFile = "E:\DownloadedImages\2-5.txt"
    'this processes is going to be asynchronous process,
    'so we have to explicitely wait till shell command gets executed
    Application.Wait Now + TimeValue("00:00:01")
    
    'in below you see, we are just reading the extracted text from the output text file and
    'putting it in cell A2.
    Open myFile For Input As #1
    Do Until EOF(1)
    Line Input #1, textline
    text = text & textline
    Loop
    Range("A2").Value = text
    Close #1
    
End Sub
