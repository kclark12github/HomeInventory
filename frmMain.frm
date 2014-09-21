VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Home Inventory"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6576
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   6576
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnKFC As New ADODB.Connection
Private cmdKFC As New ADODB.Command
Const gstrProvider = "Microsoft.Jet.OLEDB.3.51"
'Const gstrConnectionString = "E:\WebShare\wwwroot\Access\KFC.mdb"
Const gstrRunTimeUserName = "admin"
Const gstrRunTimePassword = ""

Const gstrDB_Books = "E:\WebShare\wwwroot\Access\Books.mdb"
Const gstrDB_Books = "E:\WebShare\wwwroot\Access\Hobby.mdb"
Const gstrDB_Books = "E:\WebShare\wwwroot\Access\KFC.mdb"
Const gstrDB_Books = "E:\WebShare\wwwroot\Access\Music.mdb"
Const gstrDB_Books = "E:\WebShare\wwwroot\Access\Software.mdb"
Const gstrDB_Books = "E:\WebShare\wwwroot\Access\US Navy Ships.mdb"
Const gstrDB_Books = "E:\WebShare\wwwroot\Access\UserAccessInfo.mdb"
Const gstrDB_Books = "E:\WebShare\wwwroot\Access\VideoTapes.mdb"

Public DBcollection As DataBaseCollection
Private Sub Form_Load()
    DBcollection.Add
End Sub
