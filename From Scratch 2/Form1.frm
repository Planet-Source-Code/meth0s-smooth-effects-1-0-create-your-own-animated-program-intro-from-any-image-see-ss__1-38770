VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reboundlength As Long
Dim imagetrace2 As Long
Dim imagetrace1 As Long
Dim totaltracecount As Long
Dim ismore(0 To 202) As Boolean
Dim nomore(0 To 202) As Boolean
Dim C1(0 To 202) As Long
Dim Y1(0 To 202) As Long
Dim X1(0 To 202) As Long

Public Function DrawGraph()
   X1(0) = 82
   Y1(0) = 52
   C1(0) = 16765907
   X1(1) = 92
   Y1(1) = 52
   C1(1) = 16711680
   X1(2) = 102
   Y1(2) = 52
   C1(2) = 16711680
   X1(3) = 112
   Y1(3) = 52
   C1(3) = 16712194
   X1(4) = 82
   Y1(4) = 56
   C1(4) = 16646145
   X1(5) = 92
   Y1(5) = 56
   C1(5) = 14286885
   X1(6) = 102
   Y1(6) = 56
   C1(6) = 1769700
   X1(7) = 112
   Y1(7) = 56
   C1(7) = 16646145
   X1(8) = 122
   Y1(8) = 56
   C1(8) = 16712194
   X1(9) = 72
   Y1(9) = 60
   C1(9) = 16711937
   X1(10) = 82
   Y1(10) = 60
   C1(10) = 983280
   X1(11) = 92
   Y1(11) = 60
   C1(11) = 255
   X1(12) = 102
   Y1(12) = 60
   C1(12) = 255
   X1(13) = 112
   Y1(13) = 60
   C1(13) = 255
   X1(14) = 122
   Y1(14) = 60
   C1(14) = 16646145
   X1(15) = 72
   Y1(15) = 64
   C1(15) = 16384005
   X1(16) = 82
   Y1(16) = 64
   C1(16) = 255
   X1(17) = 92
   Y1(17) = 64
   C1(17) = 255
   X1(18) = 102
   Y1(18) = 64
   C1(18) = 255
   X1(19) = 112
   Y1(19) = 64
   C1(19) = 255
   X1(20) = 122
   Y1(20) = 64
   C1(20) = 255
   X1(21) = 132
   Y1(21) = 64
   C1(21) = 16711937
   X1(22) = 62
   Y1(22) = 68
   C1(22) = 16712708
   X1(23) = 72
   Y1(23) = 68
   C1(23) = 255
   X1(24) = 82
   Y1(24) = 68
   C1(24) = 62985
   X1(25) = 92
   Y1(25) = 68
   C1(25) = 255
   X1(26) = 102
   Y1(26) = 68
   C1(26) = 255
   X1(27) = 112
   Y1(27) = 68
   C1(27) = 5355
   X1(28) = 122
   Y1(28) = 68
   C1(28) = 3825
   X1(29) = 132
   Y1(29) = 68
   C1(29) = 16580610
   X1(30) = 62
   Y1(30) = 72
   C1(30) = 16646145
   X1(31) = 72
   Y1(31) = 72
   C1(31) = 43350
   X1(32) = 82
   Y1(32) = 72
   C1(32) = 65280
   X1(33) = 92
   Y1(33) = 72
   C1(33) = 255
   X1(34) = 102
   Y1(34) = 72
   C1(34) = 255
   X1(35) = 112
   Y1(35) = 72
   C1(35) = 65280
   X1(36) = 122
   Y1(36) = 72
   C1(36) = 65280
   X1(37) = 132
   Y1(37) = 72
   C1(37) = 255
   X1(38) = 142
   Y1(38) = 72
   C1(38) = 16763851
   X1(39) = 62
   Y1(39) = 76
   C1(39) = 15138840
   X1(40) = 72
   Y1(40) = 76
   C1(40) = 65280
   X1(41) = 82
   Y1(41) = 76
   C1(41) = 65280
   X1(42) = 92
   Y1(42) = 76
   C1(42) = 59415
   X1(43) = 102
   Y1(43) = 76
   C1(43) = 255
   X1(44) = 112
   Y1(44) = 76
   C1(44) = 65280
   X1(45) = 122
   Y1(45) = 76
   C1(45) = 65280
   X1(46) = 132
   Y1(46) = 76
   C1(46) = 255
   X1(47) = 142
   Y1(47) = 76
   C1(47) = 16711937
   X1(48) = 62
   Y1(48) = 80
   C1(48) = 255
   X1(49) = 72
   Y1(49) = 80
   C1(49) = 65280
   X1(50) = 82
   Y1(50) = 80
   C1(50) = 65280
   X1(51) = 92
   Y1(51) = 80
   C1(51) = 65025
   X1(52) = 102
   Y1(52) = 80
   C1(52) = 255
   X1(53) = 112
   Y1(53) = 80
   C1(53) = 65280
   X1(54) = 122
   Y1(54) = 80
   C1(54) = 65280
   X1(55) = 132
   Y1(55) = 80
   C1(55) = 255
   X1(56) = 142
   Y1(56) = 80
   C1(56) = 16711680
   X1(57) = 52
   Y1(57) = 84
   C1(57) = 16723245
   X1(58) = 62
   Y1(58) = 84
   C1(58) = 255
   X1(59) = 72
   Y1(59) = 84
   C1(59) = 65280
   X1(60) = 82
   Y1(60) = 84
   C1(60) = 65280
   X1(61) = 92
   Y1(61) = 84
   C1(61) = 65280
   X1(62) = 102
   Y1(62) = 84
   C1(62) = 255
   X1(63) = 112
   Y1(63) = 84
   C1(63) = 65280
   X1(64) = 122
   Y1(64) = 84
   C1(64) = 65280
   X1(65) = 132
   Y1(65) = 84
   C1(65) = 255
   X1(66) = 142
   Y1(66) = 84
   C1(66) = 16646145
   X1(67) = 52
   Y1(67) = 88
   C1(67) = 16712194
   X1(68) = 62
   Y1(68) = 88
   C1(68) = 255
   X1(69) = 72
   Y1(69) = 88
   C1(69) = 65280
   X1(70) = 82
   Y1(70) = 88
   C1(70) = 65280
   X1(71) = 92
   Y1(71) = 88
   C1(71) = 64005
   X1(72) = 102
   Y1(72) = 88
   C1(72) = 255
   X1(73) = 112
   Y1(73) = 88
   C1(73) = 65280
   X1(74) = 122
   Y1(74) = 88
   C1(74) = 65280
   X1(75) = 132
   Y1(75) = 88
   C1(75) = 255
   X1(76) = 142
   Y1(76) = 88
   C1(76) = 13238325
   X1(77) = 52
   Y1(77) = 92
   C1(77) = 16711680
   X1(78) = 62
   Y1(78) = 92
   C1(78) = 255
   X1(79) = 72
   Y1(79) = 92
   C1(79) = 63240
   X1(80) = 82
   Y1(80) = 92
   C1(80) = 65280
   X1(81) = 92
   Y1(81) = 92
   C1(81) = 255
   X1(82) = 102
   Y1(82) = 92
   C1(82) = 255
   X1(83) = 112
   Y1(83) = 92
   C1(83) = 65280
   X1(84) = 122
   Y1(84) = 92
   C1(84) = 65280
   X1(85) = 132
   Y1(85) = 92
   C1(85) = 255
   X1(86) = 142
   Y1(86) = 92
   C1(86) = 255
   X1(87) = 52
   Y1(87) = 96
   C1(87) = 16711937
   X1(88) = 62
   Y1(88) = 96
   C1(88) = 255
   X1(89) = 72
   Y1(89) = 96
   C1(89) = 255
   X1(90) = 82
   Y1(90) = 96
   C1(90) = 65280
   X1(91) = 92
   Y1(91) = 96
   C1(91) = 255
   X1(92) = 102
   Y1(92) = 96
   C1(92) = 16764672
   X1(93) = 112
   Y1(93) = 96
   C1(93) = 56865
   X1(94) = 122
   Y1(94) = 96
   C1(94) = 52275
   X1(95) = 132
   Y1(95) = 96
   C1(95) = 255
   X1(96) = 142
   Y1(96) = 96
   C1(96) = 255
   X1(97) = 52
   Y1(97) = 100
   C1(97) = 16711680
   X1(98) = 62
   Y1(98) = 100
   C1(98) = 255
   X1(99) = 72
   Y1(99) = 100
   C1(99) = 255
   X1(100) = 82
   Y1(100) = 100
   C1(100) = 255
   X1(101) = 92
   Y1(101) = 100
   C1(101) = 255
   X1(102) = 102
   Y1(102) = 100
   C1(102) = 16763904
   X1(103) = 112
   Y1(103) = 100
   C1(103) = 255
   X1(104) = 122
   Y1(104) = 100
   C1(104) = 255
   X1(105) = 132
   Y1(105) = 100
   C1(105) = 9630926
   X1(106) = 142
   Y1(106) = 100
   C1(106) = 255
   X1(107) = 52
   Y1(107) = 104
   C1(107) = 16711937
   X1(108) = 62
   Y1(108) = 104
   C1(108) = 255
   X1(109) = 72
   Y1(109) = 104
   C1(109) = 10026700
   X1(110) = 82
   Y1(110) = 104
   C1(110) = 255
   X1(111) = 92
   Y1(111) = 104
   C1(111) = 255
   X1(112) = 102
   Y1(112) = 104
   C1(112) = 16772608
   X1(113) = 112
   Y1(113) = 104
   C1(113) = 255
   X1(114) = 122
   Y1(114) = 104
   C1(114) = 255
   X1(115) = 132
   Y1(115) = 104
   C1(115) = 10092492
   X1(116) = 142
   Y1(116) = 104
   C1(116) = 255
   X1(117) = 52
   Y1(117) = 108
   C1(117) = 16711937
   X1(118) = 62
   Y1(118) = 108
   C1(118) = 255
   X1(119) = 72
   Y1(119) = 108
   C1(119) = 10092492
   X1(120) = 82
   Y1(120) = 108
   C1(120) = 255
   X1(121) = 92
   Y1(121) = 108
   C1(121) = 255
   X1(122) = 102
   Y1(122) = 108
   C1(122) = 255
   X1(123) = 112
   Y1(123) = 108
   C1(123) = 255
   X1(124) = 122
   Y1(124) = 108
   C1(124) = 9433039
   X1(125) = 132
   Y1(125) = 108
   C1(125) = 10092492
   X1(126) = 142
   Y1(126) = 108
   C1(126) = 1507560
   X1(127) = 52
   Y1(127) = 112
   C1(127) = 16712194
   X1(128) = 62
   Y1(128) = 112
   C1(128) = 255
   X1(129) = 72
   Y1(129) = 112
   C1(129) = 10092492
   X1(130) = 82
   Y1(130) = 112
   C1(130) = 7058139
   X1(131) = 92
   Y1(131) = 112
   C1(131) = 255
   X1(132) = 102
   Y1(132) = 112
   C1(132) = 255
   X1(133) = 112
   Y1(133) = 112
   C1(133) = 255
   X1(134) = 122
   Y1(134) = 112
   C1(134) = 10092492
   X1(135) = 132
   Y1(135) = 112
   C1(135) = 10092492
   X1(136) = 142
   Y1(136) = 112
   C1(136) = 15794190
   X1(137) = 52
   Y1(137) = 116
   C1(137) = 16746375
   X1(138) = 62
   Y1(138) = 116
   C1(138) = 255
   X1(139) = 72
   Y1(139) = 116
   C1(139) = 10092492
   X1(140) = 82
   Y1(140) = 116
   C1(140) = 10092492
   X1(141) = 92
   Y1(141) = 116
   C1(141) = 255
   X1(142) = 102
   Y1(142) = 116
   C1(142) = 255
   X1(143) = 112
   Y1(143) = 116
   C1(143) = 8839378
   X1(144) = 122
   Y1(144) = 116
   C1(144) = 10092492
   X1(145) = 132
   Y1(145) = 116
   C1(145) = 10092492
   X1(146) = 142
   Y1(146) = 116
   C1(146) = 16646145
   X1(147) = 62
   Y1(147) = 120
   C1(147) = 255
   X1(148) = 72
   Y1(148) = 120
   C1(148) = 10092492
   X1(149) = 82
   Y1(149) = 120
   C1(149) = 10092492
   X1(150) = 92
   Y1(150) = 120
   C1(150) = 10092492
   X1(151) = 102
   Y1(151) = 120
   C1(151) = 7388122
   X1(152) = 112
   Y1(152) = 120
   C1(152) = 10092492
   X1(153) = 122
   Y1(153) = 120
   C1(153) = 10092492
   X1(154) = 132
   Y1(154) = 120
   C1(154) = 9828813
   X1(155) = 142
   Y1(155) = 120
   C1(155) = 16711937
   X1(156) = 62
   Y1(156) = 124
   C1(156) = 16384005
   X1(157) = 72
   Y1(157) = 124
   C1(157) = 10092492
   X1(158) = 82
   Y1(158) = 124
   C1(158) = 10092492
   X1(159) = 92
   Y1(159) = 124
   C1(159) = 10092492
   X1(160) = 102
   Y1(160) = 124
   C1(160) = 10092492
   X1(161) = 112
   Y1(161) = 124
   C1(161) = 10092492
   X1(162) = 122
   Y1(162) = 124
   C1(162) = 10092492
   X1(163) = 132
   Y1(163) = 124
   C1(163) = 255
   X1(164) = 142
   Y1(164) = 124
   C1(164) = 16712451
   X1(165) = 62
   Y1(165) = 128
   C1(165) = 16711937
   X1(166) = 72
   Y1(166) = 128
   C1(166) = 659964
   X1(167) = 82
   Y1(167) = 128
   C1(167) = 10092492
   X1(168) = 92
   Y1(168) = 128
   C1(168) = 10092492
   X1(169) = 102
   Y1(169) = 128
   C1(169) = 10092492
   X1(170) = 112
   Y1(170) = 128
   C1(170) = 10092492
   X1(171) = 122
   Y1(171) = 128
   C1(171) = 10092492
   X1(172) = 132
   Y1(172) = 128
   C1(172) = 2097375
   X1(173) = 62
   Y1(173) = 132
   C1(173) = 16716049
   X1(174) = 72
   Y1(174) = 132
   C1(174) = 255
   X1(175) = 82
   Y1(175) = 132
   C1(175) = 10026700
   X1(176) = 92
   Y1(176) = 132
   C1(176) = 10092492
   X1(177) = 102
   Y1(177) = 132
   C1(177) = 10092492
   X1(178) = 112
   Y1(178) = 132
   C1(178) = 10092492
   X1(179) = 122
   Y1(179) = 132
   C1(179) = 255
   X1(180) = 132
   Y1(180) = 132
   C1(180) = 16711680
   X1(181) = 72
   Y1(181) = 136
   C1(181) = 16646145
   X1(182) = 82
   Y1(182) = 136
   C1(182) = 255
   X1(183) = 92
   Y1(183) = 136
   C1(183) = 255
   X1(184) = 102
   Y1(184) = 136
   C1(184) = 3100655
   X1(185) = 112
   Y1(185) = 136
   C1(185) = 255
   X1(186) = 122
   Y1(186) = 136
   C1(186) = 255
   X1(187) = 132
   Y1(187) = 136
   C1(187) = 16712708
   X1(188) = 72
   Y1(188) = 140
   C1(188) = 16712451
   X1(189) = 82
   Y1(189) = 140
   C1(189) = 14680095
   X1(190) = 92
   Y1(190) = 140
   C1(190) = 255
   X1(191) = 102
   Y1(191) = 140
   C1(191) = 255
   X1(192) = 112
   Y1(192) = 140
   C1(192) = 255
   X1(193) = 122
   Y1(193) = 140
   C1(193) = 16711680
   X1(194) = 82
   Y1(194) = 144
   C1(194) = 16711937
   X1(195) = 92
   Y1(195) = 144
   C1(195) = 16646145
   X1(196) = 102
   Y1(196) = 144
   C1(196) = 16449540
   X1(197) = 112
   Y1(197) = 144
   C1(197) = 16646145
   X1(198) = 122
   Y1(198) = 144
   C1(198) = 16722731
   X1(199) = 92
   Y1(199) = 148
   C1(199) = 16712451
   X1(200) = 102
   Y1(200) = 148
   C1(200) = 16711937
   X1(201) = 112
   Y1(201) = 148
   C1(201) = 16736095
End Function
Private Sub Form_Load()
   Timer1.Enabled = False
   reboundlength = 2
   Picture1.ScaleMode = 3
   Picture1.AutoRedraw = True
   Me.ScaleMode = 3
   Call DrawGraph
   Timer1.Interval = 10
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   Picture1.Visible = False
   Picture1.Cls
For i = 0 To 201
    If nomore(i) = True Then
        X1(i) = X1(i) - reboundlength
        If X1(i) <= 0 Then
            nomore(i) = False
        End If
    Else
        If X1(i) >= Picture1.Width Then
            X1(i) = X1(i) - reboundlength
            nomore(i) = True
        Else
            X1(i) = X1(i) + reboundlength
        End If
    End If
    If ismore(i) = True Then
        Y1(i) = Y1(i) - reboundlength
        If Y1(i) <= 0 Then
            ismore(i) = False
        End If
    Else
        If Y1(i) >= Picture1.Height Then
            Y1(i) = Y1(i) - reboundlength
            ismore(i) = True
        Else
            Y1(i) = Y1(i) + reboundlength
        End If
    End If
    Picture1.PSet (X1(i), Y1(i)), C1(i)
Next i
   Picture1.Visible = True
End Sub

