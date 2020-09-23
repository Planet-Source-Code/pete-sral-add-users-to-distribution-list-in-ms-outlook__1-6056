VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "Form1"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   750
      Left            =   195
      TabIndex        =   0
      Top             =   90
      Width           =   4365
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'note: I only take credit for making the class.  The source is
'from Microsoft's web site.  (Attached is the web file in the zip)

'You will need to install the CDO.dll in order for this too work
'which inturn means you also need to have Outlook installed on
'your machine.  You can read more about that in the CDO zip attached

Private Sub Command1_Click()

clsDL.DistributionList = "Mailing List"
clsDL.NewUserName = "Pete Sral"
clsDL.Email = "PETE@pjs-inc.com"
clsDL.Main

'one thing I did was create a collection and look through
'the collection and add a bunch at one time.  If you do
'this you will have to tweak the function that adds the
'DL b/c if you don't it creates a DL for each on you try
'to add.  If your interested in that code, please e-mail
'me at Pete@pjs-inc.com





End Sub

