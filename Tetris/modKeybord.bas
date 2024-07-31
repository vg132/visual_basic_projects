Attribute VB_Name = "modKeybord"
Const VK_CANCEL = &H3
Const VK_BACK = &H8
Const VK_TAB = &H9
Const VK_CLEAR = &HC
Const VK_RETURN = &HD
Const VK_SHIFT = &H10
Const VK_CONTROL = &H11
Const VK_MENU = &H12
Const VK_PAUSE = &H13
Const VK_CAPITAL = &H14
Const VK_ESCAPE = &H1B
Const VK_SPACE = &H20
Const VK_PRIOR = &H21
Const VK_NEXT = &H22
Const VK_END = &H23
Const VK_HOME = &H24
Const VK_LEFT = &H25
Const VK_UP = &H26
Const VK_RIGHT = &H27
Const VK_DOWN = &H28
Const VK_SELECT = &H29
Const VK_PRINT = &H2A
Const VK_EXECUTE = &H2B
Const VK_SNAPSHOT = &H2C
Const VK_INSERT = &H2D
Const VK_DELETE = &H2E
Const VK_HELP = &H2F
Const VK_0 = &H30
Const VK_1 = &H31
Const VK_2 = &H32
Const VK_3 = &H33
Const VK_4 = &H34
Const VK_5 = &H35
Const VK_6 = &H36
Const VK_7 = &H37
Const VK_8 = &H38
Const VK_9 = &H39
Const VK_A = &H41
Const VK_B = &H42
Const VK_C = &H43
Const VK_D = &H44
Const VK_E = &H45
Const VK_F = &H46
Const VK_G = &H47
Const VK_H = &H48
Const VK_I = &H49
Const VK_J = &H4A
Const VK_K = &H4B
Const VK_L = &H4C
Const VK_M = &H4D
Const VK_N = &H4E
Const VK_O = &H4F
Const VK_P = &H50
Const VK_Q = &H51
Const VK_R = &H52
Const VK_S = &H53
Const VK_T = &H54
Const VK_U = &H55
Const VK_V = &H56
Const VK_W = &H57
Const VK_X = &H58
Const VK_Y = &H59
Const VK_Z = &H5A
Const VK_STARTKEY = &H5B
Const VK_CONTEXTKEY = &H5D
Const VK_NUMPAD0 = &H60
Const VK_NUMPAD1 = &H61
Const VK_NUMPAD2 = &H62
Const VK_NUMPAD3 = &H63
Const VK_NUMPAD4 = &H64
Const VK_NUMPAD5 = &H65
Const VK_NUMPAD6 = &H66
Const VK_NUMPAD7 = &H67
Const VK_NUMPAD8 = &H68
Const VK_NUMPAD9 = &H69
Const VK_MULTIPLY = &H6A
Const VK_ADD = &H6B
Const VK_SEPARATOR = &H6C
Const VK_SUBTRACT = &H6D
Const VK_DECIMAL = &H6E
Const VK_DIVIDE = &H6F
Const VK_F1 = &H70
Const VK_F2 = &H71
Const VK_F3 = &H72
Const VK_F4 = &H73
Const VK_F5 = &H74
Const VK_F6 = &H75
Const VK_F7 = &H76
Const VK_F8 = &H77
Const VK_F9 = &H78
Const VK_F10 = &H79
Const VK_F11 = &H7A
Const VK_F12 = &H7B
Const VK_F13 = &H7C
Const VK_F14 = &H7D
Const VK_F15 = &H7E
Const VK_F16 = &H7F
Const VK_F17 = &H80
Const VK_F18 = &H81
Const VK_F19 = &H82
Const VK_F20 = &H83
Const VK_F21 = &H84
Const VK_F22 = &H85
Const VK_F23 = &H86
Const VK_F24 = &H87
Const VK_NUMLOCK = &H90
Const VK_OEM_SCROLL = &H91
Const VK_OEM_1 = &HBA
Const VK_OEM_PLUS = &HBB
Const VK_OEM_COMMA = &HBC
Const VK_OEM_MINUS = &HBD
Const VK_OEM_PERIOD = &HBE
Const VK_OEM_2 = &HBF
Const VK_OEM_3 = &HC0
Const VK_OEM_4 = &HDB
Const VK_OEM_5 = &HDC
Const VK_OEM_6 = &HDD
Const VK_OEM_7 = &HDE
Const VK_OEM_8 = &HDF
Const VK_ICO_F17 = &HE0
Const VK_ICO_F18 = &HE1
Const VK_OEM102 = &HE2
Const VK_ICO_HELP = &HE3
Const VK_ICO_00 = &HE4
Const VK_ICO_CLEAR = &HE6
Const VK_OEM_RESET = &HE9
Const VK_OEM_JUMP = &HEA
Const VK_OEM_PA1 = &HEB
Const VK_OEM_PA2 = &HEC
Const VK_OEM_PA3 = &HED
Const VK_OEM_WSCTRL = &HEE
Const VK_OEM_CUSEL = &HEF
Const VK_OEM_ATTN = &HF0
Const VK_OEM_FINNISH = &HF1
Const VK_OEM_COPY = &HF2
Const VK_OEM_AUTO = &HF3
Const VK_OEM_ENLW = &HF4
Const VK_OEM_BACKTAB = &HF5
Const VK_ATTN = &HF6
Const VK_CRSEL = &HF7
Const VK_EXSEL = &HF8
Const VK_EREOF = &HF9
Const VK_PLAY = &HFA
Const VK_ZOOM = &HFB
Const VK_NONAME = &HFC
Const VK_PA1 = &HFD
Const VK_OEM_CLEAR = &HFE

Public Function getKeyName(KeyCode As Integer) As String
    Select Case KeyCode
    Case VK_BACK
        getKeyName = "Backspace"
    Case VK_TAB
        getKeyName = "Tab"
    Case VK_CLEAR
        getKeyName = "5 (keypad)"
    Case VK_RETURN
        getKeyName = "Enter"
    Case VK_SHIFT
        getKeyName = "Shift"
    Case VK_CONTROL
        getKeyName = "Ctrl"
    Case VK_MENU
        getKeyName = "Alt"
    Case VK_PAUSE
        getKeyName = "Pause"
    Case VK_CAPITAL
        getKeyName = "Caps Lock"
    Case VK_ESCAPE
        getKeyName = "Esc"
    Case VK_SPACE
        getKeyName = "Spacebar"
    Case VK_PRIOR
        getKeyName = "Page Up"
    Case VK_NEXT
        getKeyName = "Page Down"
    Case VK_END
        getKeyName = "End"
    Case VK_HOME
        getKeyName = "Home"
    Case VK_LEFT
        getKeyName = "Left Arrow"
    Case VK_UP
        getKeyName = "Up Arrow"
    Case VK_RIGHT
        getKeyName = "Right Arrow"
    Case VK_DOWN
        getKeyName = "Down Arrow"
    Case VK_SELECT
        getKeyName = "Select"
    Case VK_PRINT
        getKeyName = "Print"
    Case VK_SNAPSHOT
        getKeyName = "Print Screen"
    Case VK_INSERT
        getKeyName = "Insert"
    Case VK_DELETE
        getKeyName = "Delete"
    Case VK_HELP
        getKeyName = "Help"
    Case VK_0
        getKeyName = "0"
    Case VK_1
        getKeyName = "1"
    Case VK_2
        getKeyName = "2"
    Case VK_3
        getKeyName = "3"
    Case VK_4
        getKeyName = "4"
    Case VK_5
        getKeyName = "5"
    Case VK_6
        getKeyName = "6"
    Case VK_7
        getKeyName = "7"
    Case VK_8
        getKeyName = "8"
    Case VK_9
        getKeyName = "9"
    Case VK_A
        getKeyName = "A"
    Case VK_B
        getKeyName = "B"
    Case VK_C
        getKeyName = "C"
    Case VK_D
        getKeyName = "D"
    Case VK_E
        getKeyName = "E"
    Case VK_F
        getKeyName = "F"
    Case VK_G
        getKeyName = "G"
    Case VK_H
        getKeyName = "H"
    Case VK_I
        getKeyName = "I"
    Case VK_J
        getKeyName = "J"
    Case VK_K
        getKeyName = "K"
    Case VK_L
        getKeyName = "L"
    Case VK_M
        getKeyName = "M"
    Case VK_N
        getKeyName = "N"
    Case VK_O
        getKeyName = "O"
    Case VK_P
        getKeyName = "P"
    Case VK_Q
        getKeyName = "Q"
    Case VK_R
        getKeyName = "R"
    Case VK_S
        getKeyName = "S"
    Case VK_T
        getKeyName = "T"
    Case VK_U
        getKeyName = "U"
    Case VK_V
        getKeyName = "V"
    Case VK_W
        getKeyName = "W"
    Case VK_X
        getKeyName = "x"
    Case VK_Y
        getKeyName = "y"
    Case VK_Z
        getKeyName = "Z"
    Case VK_STARTKEY
        getKeyName = "Start Menu key"
    Case VK_CONTEXTKEY
        getKeyName = "Context Menu key"
    Case VK_NUMPAD0
        getKeyName = "0 (keypad)"
    Case VK_NUMPAD1
        getKeyName = "1 (keypad)"
    Case VK_NUMPAD2
        getKeyName = "2 (keypad)"
    Case VK_NUMPAD3
        getKeyName = "3 (keypad)"
    Case VK_NUMPAD4
        getKeyName = "4 (keypad)"
    Case VK_NUMPAD5
        getKeyName = "5 (keypad)"
    Case VK_NUMPAD6
        getKeyName = "6 (keypad)"
    Case VK_NUMPAD7
        getKeyName = "7 (keypad)"
    Case VK_NUMPAD8
        getKeyName = "8 (keypad)"
    Case VK_NUMPAD9
        getKeyName = "9 (keypad)"
    Case VK_MULTIPLY
        getKeyName = "*"
    Case VK_ADD
        getKeyName = "+"
    Case VK_DECIMAL
        getKeyName = ". (keypad)"
    Case VK_DIVIDE
        getKeyName = "/"
    Case VK_F1
        getKeyName = "F1"
    Case VK_F2
        getKeyName = "F2"
    Case VK_F3
        getKeyName = "F3"
    Case VK_F4
        getKeyName = "F4"
    Case VK_F5
        getKeyName = "F5"
    Case VK_F6
        getKeyName = "F6"
    Case VK_F7
        getKeyName = "F7"
    Case VK_F8
        getKeyName = "F8"
    Case VK_F9
        getKeyName = "F9"
    Case VK_F10
        getKeyName = "F10"
    Case VK_F11
        getKeyName = "F11"
    Case VK_F12
        getKeyName = "F12"
    Case VK_F13
        getKeyName = "F13"
    Case VK_F14
        getKeyName = "F14"
    Case VK_F15
        getKeyName = "F15"
    Case VK_F16
        getKeyName = "F16"
    Case VK_F17
        getKeyName = "F17"
    Case VK_F18
        getKeyName = "F18"
    Case VK_F19
        getKeyName = "F19"
    Case VK_F20
        getKeyName = "F20"
    Case VK_F21
        getKeyName = "F21"
    Case VK_F22
        getKeyName = "F22"
    Case VK_F23
        getKeyName = "F23"
    Case VK_F24
        getKeyName = "F24"
    Case VK_NUMLOCK
        getKeyName = "Num Lock"
    Case VK_OEM_SCROLL
        getKeyName = "Scroll Lock"
    Case VK_OEM_1
        getKeyName = ";"
    Case VK_OEM_PLUS
        getKeyName = "="
    Case VK_OEM_COMMA
        getKeyName = ","
    Case VK_OEM_MINUS
        getKeyName = "-"
    Case VK_OEM_PERIOD
        getKeyName = "."
    Case VK_OEM_2
        getKeyName = "/"
    Case VK_OEM_3
        getKeyName = "`"
    Case VK_OEM_4
        getKeyName = "["
    Case VK_OEM_5
        getKeyName = "\"
    Case VK_OEM_6
        getKeyName = "]"
    Case VK_OEM_7
        getKeyName = "'"
    Case VK_OEM_8
        getKeyName = "(unknown)"
    Case VK_OEM_102
        getKeyName = "< or | on IBM-compatible 102 enhanced non-U.S. keyboard"
    Case VK_OEM_RESET
        getKeyName = "Reset"
    Case VK_OEM_JUMP
        getKeyName = "Jump"
    Case VK_OEM_PA1
        getKeyName = "PA1"
    Case VK_OEM_PA2
        getKeyName = "PA2"
    Case VK_OEM_PA3
        getKeyName = "PA3"
    Case VK_OEM_WSCTRL
        getKeyName = "WSCTRL"
    Case VK_OEM_CUSEL
        getKeyName = "CUSEL"
    Case VK_OEM_ATTN
        getKeyName = "ATTN"
    Case VK_OEM_FINNISH
        getKeyName = "FINNISH"
    Case VK_OEM_COPY
        getKeyName = "COPY"
    Case VK_OEM_AUTO
        getKeyName = "AUTO"
    Case VK_OEM_ENLW
        getKeyName = "ENLW"
    Case VK_OEM_BACKTAB
        getKeyName = "BACKTAB"
    Case VK_ATTN
        getKeyName = "ATTN"
    Case VK_CRSEL
        getKeyName = "CRSEL"
    Case VK_EXSEL
        getKeyName = "EXSEL"
    Case VK_EREOF
        getKeyName = "EREOF"
    Case VK_PLAY
        getKeyName = "PLAY"
    Case VK_ZOOM
        getKeyName = "Zoom"
    Case VK_NONAME
        getKeyName = "NONAME"
    Case VK_PA1
        getKeyName = "PA1"
    Case VK_OEM_CLEAR
        getKeyName = "Clear"
    End Select
End Function
