Attribute VB_Name = "Module1"
Option Explicit

'vb窗体中控件自动随窗体变化大小

Private FormOldWidth As Long
'保存窗体的原始宽度
Private FormOldHeight As Long
'保存窗体的原始高度

'在调用ResizeForm前先调用本函数
Public Sub ResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
        Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height
    Next
    On Error GoTo 0
End Sub

'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
    Dim Pos
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    
    ScaleX = FormName.ScaleWidth / FormOldWidth
    '保存窗体宽度缩放比例
    ScaleY = FormName.ScaleHeight / FormOldHeight
    '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        Pos = Split(Obj.Tag) '使用默认参数空格进行分割" "

        '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
        Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY

    Next
    On Error GoTo 0
End Sub
