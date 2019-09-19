'CODE FOR RENAMING ALL THE IMAGES IN THE bbl_pic FOLDER LOCATED IN C:\Desktop
'IT IS USED IN PRODUCING THE 'UPDATE PPT PORTFOLIO PPT' SPECIFICALLY IT RENAMES THE IMAGES GRABBED FROM BLOOMBERG
'
'
'





Function GetFiles(path As String, Optional pattern As String = "") As Collection
    Dim rv As New Collection, f
    If Right(path, 1) <> "\" Then path = path & "\"
    f = Dir(path & pattern)
    Do While Len(f) > 0
        rv.Add path & f
        f = Dir() 'no parameter
    Loop
    Set GetFiles = rv
End Function

Sub SaveAttach()

    Dim fls, f
    C = 1
    Set fls = GetFiles("C:\Users\bloomberg03\Desktop\BBL_pic\", "")
    p = "C:\Users\bloomberg03\Desktop\BBL_pic\"
    For Each f In fls
        Name_file = f
        newName = Replace(Name_file, Name_file, p & C & ".jpg")
        Debug.Print f
        Name Name_file As newName
        
        C = C + 1
    Next f

End Sub
