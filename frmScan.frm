VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner Tutorial by Skane2004@aol.com"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sock1 
      Index           =   0
      Left            =   3000
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3480
      Top             =   1560
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scan Status"
      Height          =   2415
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ListBox lstOpenPorts 
         Height          =   1425
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Open Ports:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Configuration"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txthost 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtstartport 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtendport 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Text            =   "32767"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Ending Port"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Starting Port"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Target:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status - Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   3255
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example was designed by Sean Kane (skane2004@aol.com).  Please don't bug me online,
'if you really want then just e-mail me.  This example wasn't meant to be the quickest scanner possible,
'it was meant to teach you about winsock and how to use it.  Also, I think it uses methods that
'seperate the crappy programmers from the ones that are serious about Visual Basic, and don't
'just want to make proggies or whatever.  I've commented about every line here to give you a
'better idea of what I'm doing if you aren't sure.

'I really hope you learn from this, if not about winsock, about loading controls,
'centering forms by yourself not using VB6 or bas files.

Dim curport As Long 'This specifies the next port that needs to be scanned
Dim timertarget As Long 'This specifies the next control index the timer should check on

Private Sub cmdClear_Click()
    lstOpenPorts.Clear  'Clears all the ports already found
End Sub

Private Sub cmdStart_Click()
If cmdStart.Caption = "Start" Then  'Since the button is going to be a toggle button, make sure they aren't trying to stop the scan
cmdStart.Caption = "Stop"

curport = Val(txtstartport.Text) 'Sets the global variable (set up at the top) to the first port the user wants scanned
timertarget = 0                  'Sets the index of the control the timer should check

'Make sure all of the required fields are taken care of
If txthost.Text = "" Then   'They forgot to put in an address to scan
    MsgBox "You must have a target to scan."
    Exit Sub
End If
If txtstartport = "" Then txtstartport = "1"    'They forgot the starting port...so as a default it's 1
If txtendport = "" Then txtendport = "32767"    'They forgot the ending port...so as a default it's 32767

For i = 1 To 75    'We are starting a loop, using the variable i that will increase by one and go up to 75
    Load sock1(i)      'This will load another instance of our winsock control, but now our controls will be arrays
Next i
'What have we just done???  We made a loop that repeats 75 times, and in the loop
'we have told it to load another winsock (specifically, we told it to make an array
'of our winsock control named "sock1".  Now we have 51 controls for winsock.  Please
'note: In order to do this, you must have the original control as an array (it's got
'an index).  So if you look at the properties for sock1 on my form, it has an index
'value of 0.  Now, whenever I want to call any of my controls, I just have to call them
'by index number... so if I wanted to set a value for the tag of sock1 - index #54 I would
'type sock1(54).tag = "Blah blah blah"

'Using control arrays makes things go a lot quicker by using loops...you'll see
'if you stick with the code and keep going

For i = 0 To 75    'Note how now I have 0 to 75, not 1 to 75...remember we have a control with an index of 0 as the original control?  We have to program that just like all the others
    sock1(i).Close                      'Close the control
    sock1(i).RemoteHost = txthost.Text  'This sets all the controls to want to connect to the target the user specified
    sock1(i).RemotePort = curport       'Sets the port needed to be scanned
    sock1(i).Connect                    'Try to get it to connect
    curport = curport + 1               'Makes curport get larget by one for the next control
    lblStatus.Caption = "Status - Scanning port " & curport 'Sets the status message to say the port currently being scanned
Next i

Timer1.Enabled = True   'The timer is sort of like the clean up person...it looks for controls that haven't connected and are just sitting there and assigns them a new port
Exit Sub    'We must have it exit the entire sub because if we don't, it will move down and think that the user wanted it stopped, but we don't yet
End If

If cmdStart.Caption = "Stop" Then 'The user wants to stop the scan, we have to stop all the things happening
    cmdStart.Caption = "Start"      'Sets the caption back if the user wants to go again
    Timer1.Enabled = False  'Stop the timer
    
    For i = 1 To 75    'We are going to look through every control to find the last control that is hanging
        If sock1(i).State = sckError Then 'If the sock had an error (it couldn't connect)
            If sock1(i).RemotePort > Val(txtstartport.Text) Then    'We're looking for the biggest port number that has the state of sckError
                txtstartport.Text = sock1(i).RemotePort 'This sets the user's starting port as the last port that couldn't connect
            End If
        End If
    Next i
    
    lblStatus.Caption = "Status - Stopped" 'Sets the label status
    
    For i = 1 To 75    'Repeats 75 times
        sock1(i).Close  'Stops what their doing
        Unload sock1(i) 'Unloads the controls that we loaded earlier
    Next i
End If
End Sub

Private Sub Form_Load()
'Center the form so it looks nice

Me.Left = Screen.Width / 2 - Me.Width / 2
'^This is saying to set the left part of the window on: the vertical line that makes
'up the width of the screen divided by two (which is the center of the screen)..but
'we can't stop there, because we aren't setting the center of our window at the center
'of the screen.  Thus, we have to move the imaginary vertical line left (that's why
'we subtract... to go left is subtracting, to go right you add.  This is that way
'because when we place object, we place them as if the entire screen was a big coordinate
'plane (remember math class?)... where you are given points: x and y.  The larget X is,
'the further to the left you have to go.  Therefore, if you subtract from x, it moves
'left.  So, now that that's covered, we subtract the width of our window divided by two.
'Which basically means we subtract the distance from the left side to the middle.
'If you draw this in paint it makes a lot more sense.

'Now we have to center the form vertically...It's more or less the same as above except
'It deals with heights and tops, not widths and left sides respectively.
Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub sock1_Connect(Index As Integer)
lstOpenPorts.AddItem sock1(Index).RemotePort   'This adds the port number to our listbox, since it did connect
sock1(Index).Close  'Closes the control so our timer will see it and assign it a new port
End Sub

Private Sub Timer1_Timer()
'Remember: timertarget is the variable that was declared at the top that tells the timer
'what control index it should look at next.  The control index is the index number of a control

If sock1(timertarget).State <> sckConnected Or sock1(timertarget).State <> sckConnecting Or sock1(timertarget).State <> sckHostResolved Or sock1(timertarget).State <> sckResolvingHost Then    'If the control isn't connected, connecting, resolving it's host, or already resolved it's host then...(in other words, it's hanging)
    sock1(timertarget).Close    'Stop the control from doing whatever it's doing and close it
    sock1(timertarget).RemotePort = curport 'Set the remote port for the next port needed to be scanned
    sock1(timertarget).Connect  'Have it connect
    lblStatus.Caption = "Status - Scanning Port " & sock1(timertarget).RemotePort 'Sets the status to the new port
    curport = curport + 1 'Increment the variable for the next time we ask for a new port
End If

timertarget = timertarget + 1   'Increments the index the timer should look at
If timertarget = 75 Then timertarget = 0   'If we've gone too far on the increment, this sets it back to the first index
End Sub
