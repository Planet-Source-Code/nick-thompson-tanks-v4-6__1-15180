Attribute VB_Name = "Main"
'fps = frames per second or number of ___ per second
Global doturn As Integer 'countdown to doing the turns of tanks
Global docheck As Integer 'countdown to doing the check ofr winners
Global doreload(7) As Integer 'countdown to reloading
Global dochangeai As Long 'countdown to doing AI
Global leavenow As Byte 'tell timermain to exit if not 0
Global fps As Long 'fps of moving
Global fpsdraw As Long 'fps of drawing
Global i As Long
Global j As Long
Global domove As Single 'what moveopto has to get to
Global dodraw As Single 'what drawopto has to get to
Global moveopto As Single 'countdown to doing all moving/checks
Global drawopto As Single 'countdown to doing all drawing
Global checkspeed As Integer 'used in timerfps
Global slowdown As Long 'used to slow the program down if drawing faster than 60 to 70 fps

Global gspeed As Integer 'speed of moving of game
Global armour As Integer 'amount of armour
Global gofast As Byte 'moderate speed or not
Global dback As Byte 'Show background picture
Global reloadspeed As Integer 'reload speed
Global pstate(8) As Integer 'is player computer or human
Global a As Long 'User in For... Next loops etc
Global b As Long 'User in For... Next loops etc
Global ran As Single ' for randomize timer
Global ran2 As Single
Global lefton(8) As Byte
Global righton(8) As Byte
Global upon(8) As Byte
Global downon(8) As Byte
'Global slowdown As Long ' Used to run game at correct speed on all computers
Global ns As Integer 'number of shells
Global pfire(8) As Byte 'Unit fire
Global pxpos(8) As Single 'Tank x coordinates
Global pypos(8) As Single 'Tank y coordinates
Global ptxpos(8) As Single 'Tank x coordinates (temp)
Global ptypos(8) As Single 'Tank y coordinates (temp)
Global pdir(8) As Integer 'Tank Direction
Global pbdir(8) As Integer 'Bounce Direction
Global preloaded(8) As Byte 'Has tank reloaded?
Global ps(8) As Single 'Tank Speed
Global pbs(8) As Single 'Bounce Speed
Global ph(8) As Integer 'Tank Health
Global pt(8) As Integer 'Currant target (for AI)
Global sdir(1500) As Integer 'Shell direction starts 0, techniqually you can have 100 shells on screen at once - hey major slowdown
Global ss(1500) As Integer  'Shell speed - combination of tank speed and exit velocity
Global sxpos(1500) As Single 'shell x coord
Global sypos(1500) As Single 'shell y coord
Global sown(1500) As Integer 'Who fired shell?

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' Straight Copy
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest


Public Function ms(thiscontrol As Control, formwidth As Long)
  thiscontrol.Left = (formwidth - thiscontrol.Width) / 2
End Function
