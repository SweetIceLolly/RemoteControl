Attribute VB_Name = "modMouseControl"
'��갴������
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'���λ�ÿ���
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
'����������
Public Declare Function BlockInput Lib "user32" (ByVal fBlockIt As Long) As Long

'��갴����������
Public Const LeftDown = &H2             '����
Public Const LeftUp = &H4               '����
Public Const MiddleDown = &H20          '����
Public Const MiddleUp = &H40            '����
Public Const RightDown = &H8            '����
Public Const RightUp = &H10             '����
Public Const MOUSEEVENTF_WHEEL = &H800  '����
