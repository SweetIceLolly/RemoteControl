Attribute VB_Name = "modMouseControl"
'鼠标按键操作
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'鼠标位置控制
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
'禁用鼠标键盘
Public Declare Function BlockInput Lib "user32" (ByVal fBlockIt As Long) As Long

'鼠标按键操作常数
Public Const LeftDown = &H2             '左下
Public Const LeftUp = &H4               '左上
Public Const MiddleDown = &H20          '中下
Public Const MiddleUp = &H40            '中上
Public Const RightDown = &H8            '右下
Public Const RightUp = &H10             '右上
Public Const MOUSEEVENTF_WHEEL = &H800  '滚轮
