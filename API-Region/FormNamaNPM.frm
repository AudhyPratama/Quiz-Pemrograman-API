VERSION 5.00
Begin VB.Form FormNamaNPM 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormNamaNPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()

bg = CreateRoundRectRgn(10, 30, 1200, 470, 20, 20) 'Background'
CombineRgn bg, bg, bg, 4

'################################################################################'
'Huruf A'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n1 = CreateRoundRectRgn(30, 80, 50, 170, 0, 0) 'Garis Vertikal Kiri'
n2 = CreateRoundRectRgn(50, 60, 80, 80, 0, 0) 'Garis Horizontal Tengah atas'
n3 = CreateRoundRectRgn(80, 80, 100, 170, 0, 0) 'Garis Vertikal Kanan'
n4 = CreateRoundRectRgn(30, 110, 100, 130, 0, 0) 'Garis Horizontal Tengah'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n1, 2
CombineRgn bg, bg, n2, 2
CombineRgn bg, bg, n3, 2
CombineRgn bg, bg, n4, 2

'################################################################################'
'Huruf U'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n5 = CreateRoundRectRgn(110, 60, 130, 150, 0, 0) 'Garis Vertikal Kiri'
n6 = CreateRoundRectRgn(130, 150, 160, 170, 0, 0) 'Garis Horizontal Bawah'
n7 = CreateRoundRectRgn(160, 60, 180, 150, 0, 0) 'Garis Vertikal Kanan'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n5, 2
CombineRgn bg, bg, n6, 2
CombineRgn bg, bg, n7, 2

'################################################################################'
'Huruf D'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n8 = CreateRoundRectRgn(190, 60, 210, 170, 0, 0) 'Garis Vertikal Kiri'
n9 = CreateRoundRectRgn(240, 80, 260, 150, 0, 0) 'Garis Vertikal Kanan'
n10 = CreateRoundRectRgn(190, 60, 240, 80, 0, 0) 'Garis Horizontal Tengah Atas'
n11 = CreateRoundRectRgn(190, 150, 240, 170, 0, 0) 'Garis Horizontal Tengah Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n8, 2
CombineRgn bg, bg, n9, 2
CombineRgn bg, bg, n10, 2
CombineRgn bg, bg, n11, 2

'################################################################################'
'Huruf H'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n12 = CreateRoundRectRgn(270, 60, 290, 170, 0, 0) 'Garis Vertikal Kiri'
n13 = CreateRoundRectRgn(270, 110, 330, 130, 0, 0) 'Garis Horizontal Tengah'
n14 = CreateRoundRectRgn(320, 60, 340, 170, 0, 0) 'Garis Vertikal Kanan'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n12, 2
CombineRgn bg, bg, n13, 2
CombineRgn bg, bg, n14, 2

'################################################################################'
'Huruf Y'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n15 = CreateRoundRectRgn(360, 60, 380, 100, 0, 0) 'Garis Vertikal Kiri'
n16 = CreateRoundRectRgn(360, 90, 440, 110, 0, 0) 'Garis Horizontal Tengah'
n17 = CreateRoundRectRgn(420, 60, 440, 100, 0, 0) 'Garis Vertikal Kanan'
n18 = CreateRoundRectRgn(390, 90, 410, 170, 0, 0) 'Garis Vertikal Tengah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n15, 2
CombineRgn bg, bg, n16, 2
CombineRgn bg, bg, n17, 2
CombineRgn bg, bg, n18, 2

'################################################################################'
'Huruf B'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n19 = CreateRoundRectRgn(480, 60, 500, 170, 0, 0) 'Garis Vertikal Kiri'
n20 = CreateRoundRectRgn(540, 80, 560, 150, 0, 0) 'Garis Vertikal Kanan'
n21 = CreateRoundRectRgn(480, 60, 540, 80, 0, 0) 'Garis Horizontal Atas'
n22 = CreateRoundRectRgn(480, 110, 540, 130, 0, 0) 'Garis Horizontal Tengah'
n23 = CreateRoundRectRgn(480, 150, 540, 170, 0, 0) 'Garis Horizontal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n19, 2
CombineRgn bg, bg, n20, 2
CombineRgn bg, bg, n21, 2
CombineRgn bg, bg, n22, 2
CombineRgn bg, bg, n23, 2


'################################################################################'
'Huruf R'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n24 = CreateRoundRectRgn(580, 60, 600, 170, 0, 0) 'Garis Vertikal Kiri'
n25 = CreateEllipticRgn(580, 60, 640, 130) 'Lingkaran Luar'
n26 = CreateEllipticRgn(590, 70, 630, 120) 'Lingkaran Dalam'
n27 = CreateRoundRectRgn(620, 120, 630, 170, 50, 0) 'Garis Vertikal Bawah'
n271 = CreateRoundRectRgn(620, 160, 640, 170, 50, 0) 'Garis Horizontal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n24, 2
CombineRgn bg, bg, n25, 2
CombineRgn bg, bg, n26, 4
CombineRgn bg, bg, n27, 2
CombineRgn bg, bg, n271, 2

'################################################################################'
'Huruf I'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n28 = CreateRoundRectRgn(660, 60, 680, 170, 0, 0) 'Garis Vertikal'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n28, 2

'################################################################################'
'Huruf L'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n29 = CreateRoundRectRgn(700, 60, 720, 170, 0, 0) 'Garis Vertikal'
n30 = CreateRoundRectRgn(700, 150, 760, 170, 0, 0) 'Garis Horizontal'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n29, 2
CombineRgn bg, bg, n30, 2

'################################################################################'
'Huruf L'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n31 = CreateRoundRectRgn(780, 60, 800, 170, 0, 0) 'Garis Vertikal'
n32 = CreateRoundRectRgn(780, 150, 840, 170, 0, 0) 'Garis Horizontal'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n31, 2
CombineRgn bg, bg, n32, 2

'################################################################################'
'Huruf I'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n33 = CreateRoundRectRgn(860, 60, 880, 170, 0, 0) 'Garis Vertikal'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n33, 2

'################################################################################'
'Huruf A'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n34 = CreateRoundRectRgn(900, 80, 920, 170, 0, 0) 'Garis Vertikal Kiri'
n35 = CreateRoundRectRgn(960, 80, 980, 170, 0, 0) 'Garis Vertikal Kanan'
n36 = CreateRoundRectRgn(920, 60, 960, 80, 0, 0) 'Garis Horizontal Tengah atas'
n37 = CreateRoundRectRgn(920, 110, 960, 130, 0, 0) 'Garis Horizontal Tengah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n34, 2
CombineRgn bg, bg, n35, 2
CombineRgn bg, bg, n36, 2
CombineRgn bg, bg, n37, 2

'################################################################################'
'Huruf N'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n38 = CreateRoundRectRgn(1000, 60, 1020, 170, 0, 0) 'Garis Vertikal Kiri'
n39 = CreateRoundRectRgn(1070, 60, 1090, 170, 0, 0) 'Garis Vertikal Kanan'
n39 = CreateRoundRectRgn(1070, 60, 1090, 170, 0, 0) 'Garis Vertikal Kanan'
n391 = CreateRoundRectRgn(1020, 60, 1030, 80, 0, 0) 'Garis Diagonal 1'
n392 = CreateRoundRectRgn(1030, 80, 1040, 100, 0, 0) 'Garis Diagonal 2'
n393 = CreateRoundRectRgn(1040, 100, 1050, 130, 0, 0) 'Garis Diagonal 3'
n394 = CreateRoundRectRgn(1050, 130, 1060, 150, 0, 0) 'Garis Diagonal 3'
n395 = CreateRoundRectRgn(1060, 150, 1070, 170, 0, 0) 'Garis Diagonal 3'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n38, 2
CombineRgn bg, bg, n39, 2
CombineRgn bg, bg, n391, 2
CombineRgn bg, bg, n392, 2
CombineRgn bg, bg, n393, 2
CombineRgn bg, bg, n394, 2
CombineRgn bg, bg, n395, 2

'################################################################################'
'Huruf T'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n40 = CreateRoundRectRgn(1130, 60, 1150, 170, 0, 0) 'Garis Vertikal Tengah'
n41 = CreateRoundRectRgn(1100, 60, 1180, 80, 0, 0) 'Garis Horizontal Atas'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n40, 2
CombineRgn bg, bg, n41, 2


'===============================================================BARIS BAWAHNYA==============================================================='


'################################################################################'
'Huruf P'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n42 = CreateRoundRectRgn(30, 200, 50, 310, 0, 0) 'Garis Vertikal Kiri'
n43 = CreateEllipticRgn(30, 200, 100, 270) 'Lingkaran Luar'
n44 = CreateEllipticRgn(40, 210, 90, 260) 'Lingkaran Dalam'


'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n42, 2
CombineRgn bg, bg, n43, 2
CombineRgn bg, bg, n44, 4

'################################################################################'
'Huruf R'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n45 = CreateRoundRectRgn(120, 200, 140, 310, 0, 0) 'Garis Vertikal Kiri'
n46 = CreateEllipticRgn(120, 200, 190, 270) 'Lingkaran Luar'
n47 = CreateEllipticRgn(130, 210, 180, 260) 'Lingkaran Dalam'
n48 = CreateRoundRectRgn(160, 260, 170, 310, 0, 0) 'Garis Vertikal Bawah'
n481 = CreateRoundRectRgn(160, 300, 180, 310, 0, 0) 'Garis Horizontal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n45, 2
CombineRgn bg, bg, n46, 2
CombineRgn bg, bg, n47, 4
CombineRgn bg, bg, n48, 2
CombineRgn bg, bg, n481, 2

'################################################################################'
'Huruf A'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n49 = CreateRoundRectRgn(210, 220, 230, 310, 0, 0) 'Garis Vertikal Kiri'
n50 = CreateRoundRectRgn(270, 220, 290, 310, 0, 0) 'Garis Vertikal Kanan'
n51 = CreateRoundRectRgn(230, 200, 270, 220, 0, 0) 'Garis Horizontal Tengah atas'
n52 = CreateRoundRectRgn(230, 260, 270, 280, 0, 0) 'Garis Horizontal Tengah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n49, 2
CombineRgn bg, bg, n50, 2
CombineRgn bg, bg, n51, 2
CombineRgn bg, bg, n52, 2

'################################################################################'
'Huruf T'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n53 = CreateRoundRectRgn(340, 210, 360, 310, 0, 0) 'Garis Vertikal Tengah'
n54 = CreateRoundRectRgn(310, 200, 390, 220, 0, 0) 'Garis Horizontal Atas'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n53, 2
CombineRgn bg, bg, n54, 2

'################################################################################'
'Huruf A'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n55 = CreateRoundRectRgn(410, 220, 430, 310, 0, 0) 'Garis Vertikal Kiri'
n56 = CreateRoundRectRgn(470, 220, 490, 310, 0, 0) 'Garis Vertikal Kanan'
n57 = CreateRoundRectRgn(430, 200, 470, 220, 0, 0) 'Garis Horizontal Tengah atas'
n58 = CreateRoundRectRgn(430, 260, 470, 280, 0, 0) 'Garis Horizontal Tengah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n55, 2
CombineRgn bg, bg, n56, 2
CombineRgn bg, bg, n57, 2
CombineRgn bg, bg, n58, 2

'################################################################################'
'Huruf M'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n59 = CreateRoundRectRgn(510, 220, 530, 310, 0, 0) 'Garis Vertikal Kiri'
n60 = CreateRoundRectRgn(560, 220, 580, 310, 0, 0) 'Garis Vertikal Tengah'
n61 = CreateRoundRectRgn(610, 220, 630, 310, 0, 0) 'Garis Vertikal Kanan'
n62 = CreateRoundRectRgn(530, 200, 560, 220, 0, 0) 'Garis Horizontal Kiri Atas'
n63 = CreateRoundRectRgn(580, 200, 610, 220, 0, 0) 'Garis Horizontal Kanan Atas'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n59, 2
CombineRgn bg, bg, n60, 2
CombineRgn bg, bg, n61, 2
CombineRgn bg, bg, n62, 2
CombineRgn bg, bg, n63, 2

'################################################################################'
'Huruf A'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n64 = CreateRoundRectRgn(650, 220, 670, 310, 0, 0) 'Garis Vertikal Kiri'
n65 = CreateRoundRectRgn(710, 220, 730, 310, 0, 0) 'Garis Vertikal Kanan'
n66 = CreateRoundRectRgn(670, 200, 710, 220, 0, 0) 'Garis Horizontal Tengah atas'
n67 = CreateRoundRectRgn(670, 260, 710, 280, 0, 0) 'Garis Horizontal Tengah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n64, 2
CombineRgn bg, bg, n65, 2
CombineRgn bg, bg, n66, 2
CombineRgn bg, bg, n67, 2


'===============================================================BARIS BAWAHNYA==============================================================='


'################################################################################'
'Angka 1'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n68 = CreateRoundRectRgn(30, 340, 50, 380, 10, 10) 'Garis Vertikal Atas'
n69 = CreateRoundRectRgn(30, 400, 50, 440, 10, 10) 'Garis Vertikal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n68, 2
CombineRgn bg, bg, n69, 2

'################################################################################'
'Angka 9'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n70 = CreateRoundRectRgn(70, 340, 90, 380, 10, 10) 'Garis Vertikal Atas'
n71 = CreateRoundRectRgn(90, 320, 130, 340, 10, 10) 'Garis Horizontal Atas'
n72 = CreateRoundRectRgn(130, 340, 150, 380, 10, 10) 'Garis Vertikal Kanan'
n73 = CreateRoundRectRgn(90, 380, 130, 400, 10, 10) 'Garis Horizontal tengah'
n74 = CreateRoundRectRgn(130, 400, 150, 440, 10, 10) 'Garis Vertikal Bawah'
n75 = CreateRoundRectRgn(90, 440, 130, 460, 10, 10) 'Garis Horizontal tengah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n70, 2
CombineRgn bg, bg, n71, 2
CombineRgn bg, bg, n72, 2
CombineRgn bg, bg, n73, 2
CombineRgn bg, bg, n74, 2
CombineRgn bg, bg, n75, 2

'################################################################################'
'Angka 0'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n76 = CreateRoundRectRgn(170, 340, 190, 380, 10, 10) 'Garis Vertikal Atas Kiri'
n77 = CreateRoundRectRgn(230, 340, 250, 380, 10, 10) 'Garis Vertikal Atas Kanan'
n78 = CreateRoundRectRgn(190, 320, 230, 340, 10, 10) 'Garis Horizontal Atas'
n79 = CreateRoundRectRgn(230, 400, 250, 440, 10, 10) 'Garis Vertikal Bawah Kanan'
n80 = CreateRoundRectRgn(190, 440, 230, 460, 10, 10) 'Garis Horizontal Bawah'
n81 = CreateRoundRectRgn(170, 400, 190, 440, 10, 10) 'Garis Vertikal Bawah Kiri'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n76, 2
CombineRgn bg, bg, n77, 2
CombineRgn bg, bg, n78, 2
CombineRgn bg, bg, n79, 2
CombineRgn bg, bg, n80, 2
CombineRgn bg, bg, n81, 2

'################################################################################'
'Angka 8'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n82 = CreateRoundRectRgn(270, 340, 290, 380, 10, 10) 'Garis Vertikal Atas Kiri'
n83 = CreateRoundRectRgn(330, 340, 350, 380, 10, 10) 'Garis Vertikal Atas Kanan'
n84 = CreateRoundRectRgn(290, 320, 330, 340, 10, 10) 'Garis Horizontal Atas'
n85 = CreateRoundRectRgn(330, 400, 350, 440, 10, 10) 'Garis Vertikal Bawah Kanan'
n86 = CreateRoundRectRgn(290, 440, 330, 460, 10, 10) 'Garis Horizontal Bawah'
n87 = CreateRoundRectRgn(270, 400, 290, 440, 10, 10) 'Garis Vertikal Bawah Kiri'
n88 = CreateRoundRectRgn(290, 380, 330, 400, 10, 10) 'Garis Horizontal Tengah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n82, 2
CombineRgn bg, bg, n83, 2
CombineRgn bg, bg, n84, 2
CombineRgn bg, bg, n85, 2
CombineRgn bg, bg, n86, 2
CombineRgn bg, bg, n87, 2
CombineRgn bg, bg, n88, 2

'################################################################################'
'Angka 1'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n68 = CreateRoundRectRgn(370, 340, 390, 380, 10, 10) 'Garis Vertikal Atas'
n69 = CreateRoundRectRgn(370, 400, 390, 440, 10, 10) 'Garis Vertikal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n68, 2
CombineRgn bg, bg, n69, 2

'################################################################################'
'Angka 0'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n76 = CreateRoundRectRgn(410, 340, 430, 380, 10, 10) 'Garis Vertikal Atas Kiri'
n77 = CreateRoundRectRgn(470, 340, 490, 380, 10, 10) 'Garis Vertikal Atas Kanan'
n78 = CreateRoundRectRgn(430, 320, 470, 340, 10, 10) 'Garis Horizontal Atas'
n79 = CreateRoundRectRgn(410, 400, 430, 440, 10, 10) 'Garis Vertikal Bawah Kanan'
n80 = CreateRoundRectRgn(430, 440, 470, 460, 10, 10) 'Garis Horizontal Bawah'
n81 = CreateRoundRectRgn(470, 400, 490, 440, 10, 10) 'Garis Vertikal Bawah Kiri'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n76, 2
CombineRgn bg, bg, n77, 2
CombineRgn bg, bg, n78, 2
CombineRgn bg, bg, n79, 2
CombineRgn bg, bg, n80, 2
CombineRgn bg, bg, n81, 2

'################################################################################'
'Angka 1'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n82 = CreateRoundRectRgn(510, 340, 530, 380, 10, 10) 'Garis Vertikal Atas'
n83 = CreateRoundRectRgn(510, 400, 530, 440, 10, 10) 'Garis Vertikal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n82, 2
CombineRgn bg, bg, n83, 2

'################################################################################'
'Angka 0'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n84 = CreateRoundRectRgn(550, 340, 570, 380, 10, 10) 'Garis Vertikal Atas Kiri'
n85 = CreateRoundRectRgn(610, 340, 630, 380, 10, 10) 'Garis Vertikal Atas Kanan'
n86 = CreateRoundRectRgn(570, 320, 610, 340, 10, 10) 'Garis Horizontal Atas'
n87 = CreateRoundRectRgn(610, 400, 630, 440, 10, 10) 'Garis Vertikal Bawah Kanan'
n88 = CreateRoundRectRgn(570, 440, 610, 460, 10, 10) 'Garis Horizontal Bawah'
n89 = CreateRoundRectRgn(550, 400, 570, 440, 10, 10) 'Garis Vertikal Bawah Kiri'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n84, 2
CombineRgn bg, bg, n85, 2
CombineRgn bg, bg, n86, 2
CombineRgn bg, bg, n87, 2
CombineRgn bg, bg, n88, 2
CombineRgn bg, bg, n89, 2

'################################################################################'
'Angka 1'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n90 = CreateRoundRectRgn(650, 340, 670, 380, 10, 10) 'Garis Vertikal Atas'
n91 = CreateRoundRectRgn(650, 400, 670, 440, 10, 10) 'Garis Vertikal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n90, 2
CombineRgn bg, bg, n91, 2

'################################################################################'
'Angka 2'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'

n92 = CreateRoundRectRgn(750, 340, 770, 380, 10, 10) 'Garis Vertikal Atas Kanan'
n93 = CreateRoundRectRgn(710, 320, 750, 340, 10, 10) 'Garis Horizontal Atas'
n94 = CreateRoundRectRgn(690, 400, 710, 440, 10, 10) 'Garis Vertikal Bawah Kiri'
n95 = CreateRoundRectRgn(710, 380, 750, 400, 10, 10) 'Garis Horizontal tengah'
n96 = CreateRoundRectRgn(710, 440, 750, 460, 10, 10) 'Garis Horizontal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n92, 2
CombineRgn bg, bg, n93, 2
CombineRgn bg, bg, n94, 2
CombineRgn bg, bg, n95, 2
CombineRgn bg, bg, n96, 2

'################################################################################'
'Angka 3'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'

n97 = CreateRoundRectRgn(850, 340, 870, 380, 10, 10) 'Garis Vertikal Atas Kanan'
n98 = CreateRoundRectRgn(810, 320, 850, 340, 10, 10) 'Garis Horizontal Atas'
n99 = CreateRoundRectRgn(850, 400, 870, 440, 10, 10) 'Garis Vertikal Bawah Kanan'
n100 = CreateRoundRectRgn(810, 380, 850, 400, 10, 10) 'Garis Horizontal tengah'
n101 = CreateRoundRectRgn(810, 440, 850, 460, 10, 10) 'Garis Horizontal Bawah'

'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n97, 2
CombineRgn bg, bg, n98, 2
CombineRgn bg, bg, n99, 2
CombineRgn bg, bg, n100, 2
CombineRgn bg, bg, n101, 2


'View Hasil Combine'
SetWindowRgn FormNamaNPM.hwnd, bg, True
End Sub

'Responsive Object -> Mouse Trigger'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage FormNamaNPM.hwnd, &HA1, 2, 0&
End Sub


