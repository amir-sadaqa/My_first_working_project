Attribute VB_Name = "Module1"
Sub ������1()
Attribute ������1.VB_ProcData.VB_Invoke_Func = " \n14"
    Workbooks.Open Filename:= _
        "V:\Hyundai Mobility\Outgoing\����� ���������� ��\������\���� �������\������ � ���\���_����������.xlsx"
    Range("Q1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("���_������.xls").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").EntireColumn.AutoFit
    Windows("���_����������.xlsx").Activate
    Range("AK1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("���_������.xls").Activate
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").EntireColumn.AutoFit
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    ActiveWorkbook.Save
    Workbooks("���_����������.xlsx").Close SaveChanges:=False
    ChDir "V:\Hyundai Mobility\Outgoing\����� ���������� ��\������\���� �������\������ � ���\��������_�������"
    ActiveWorkbook.SaveAs Filename:= _
        "V:\Hyundai Mobility\Outgoing\����� ���������� ��\������\���� �������\������ � ���\��������_�������\all_damages.txt" _
        , FileFormat:=xlText, CreateBackup:=False
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    ChDir "V:\Hyundai Mobility\Outgoing\����� ���������� ��\������\���� �������\������ � ���\��������_�������"
    ActiveWorkbook.SaveAs Filename:="V:\Hyundai Mobility\Outgoing\����� ���������� ��\������\���� �������\������ � ���\��������_�������\���_������.xls", _
        FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
End Sub
