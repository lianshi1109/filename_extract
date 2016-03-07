Sub image_inf()
    Dim myPath$, AK As Workbook, OAK As Workbook
    Dim filename As String
    Dim cellname As String
    Dim filepath
    Dim f
    Dim FD
    Dim fder
    
    Dim data_sheet, data_sheet1 As String

    
    Dim i As Integer
    i = 1

    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set OAK = ActiveWorkbook
    myPath = "\\whfile.csot.tcl.com\11制造中心\整合厂\工艺整合部\2.阵列工艺整合科\12 personal\李安石\Tools\IMAGE\list"    '把文件路径定义给变量
    Application.ScreenUpdating = False
    Set FD = fso.GetFolder(myPath)
    'Set fder = FD.subFolders
     
    For Each fder In FD.subFolders
            'Worksheets.Add().Name = j
    For Each f In fder.Files
            'data_sheet = Right(fder.Path, Len(f.Path) - InStrRev(fder.Path, "\"))
                filename = VBA.Right(f.Name, 3)
                If filename = "jpg" Then
                data_sheet1 = Right(f.Path, Len(f.Path) - InStrRev(f.Path, "\", 82))
                data_sheet = Left(data_sheet1, InStrRev(data_sheet1, "\") - 1)
                'data_sheet = Left(data_sheet1, 2)
            
                Sheets("data").Select
                Cells(i + 1, 1) = data_sheet
                'Cells(i + 1, 2).Select
                Cells(i + 1, 2) = f.Name
               ' Sheets("data").Select
               ' linshi = Range("B1", "K1")
               ' Sheets("data").Select
               ' Range("B1", "K1") = linshi
               Cells(i + 1, 2).Select
                data_select = Cells(i + 1, 2).Value
                data1 = Split(data_select, "_")(0)
data2 = Split(data_select, "_")(1)
data3 = Split(data_select, "_")(2)
data4 = Split(data_select, "_")(3)
data5 = Split(data_select, "_")(4)
data6 = Split(data_select, "_")(5)
data7 = Split(data_select, "_")(6)
data8 = Split(data_select, "_")(7)
data9 = Split(data_select, "_")(8)


Cells(i + 1, 3) = data1
Cells(i + 1, 4) = data2
Cells(i + 1, 5) = data3
Cells(i + 1, 6) = data4
Cells(i + 1, 7) = data5
Cells(i + 1, 8) = data6
Cells(i + 1, 9) = data7
Cells(i + 1, 10) = data8
Cells(i + 1, 11) = data9
                i = i + 1
            Else
            End If
            Next
            Next
     
    Application.ScreenUpdating = True '冻结屏幕,此类语句一般成对使用
End Sub



