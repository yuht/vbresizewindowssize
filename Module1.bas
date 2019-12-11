Attribute VB_Name = "Module1"
Option Explicit

'vb�����пؼ��Զ��洰��仯��С

Private FormOldWidth As Long
'���洰���ԭʼ���
Private FormOldHeight As Long
'���洰���ԭʼ�߶�

'�ڵ���ResizeFormǰ�ȵ��ñ�����
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

'�������ı���ڸ�Ԫ���Ĵ�С���ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)
    Dim Pos
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    
    ScaleX = FormName.ScaleWidth / FormOldWidth
    '���洰�������ű���
    ScaleY = FormName.ScaleHeight / FormOldHeight
    '���洰��߶����ű���
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        Pos = Split(Obj.Tag) 'ʹ��Ĭ�ϲ����ո���зָ�" "

        '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
        Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY

    Next
    On Error GoTo 0
End Sub
