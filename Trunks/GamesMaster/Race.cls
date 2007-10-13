VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Race"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Race in the GalaxyNG game"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public RaceName As String
Public EMail As String
Public Password As String
Public Drive As Single
Public Weapons As Single
Public Shields As Single
Public Cargo As Single
Public mX As Single
Public mY As Single
Public mSize As Single
Public PlayerName As String
Public TeamName As String
Public Unused3 As String
Public Unused4 As String
Public Unused5 As String
Public MassLost As String
Public MassProduced As String
Public LastOrders As Long
Public flags As Long

Public Enum RaceFlag
    R_ANONYMOUS = 1          'Player want to stay anonymous
    R_AUTOUNLOAD = 2
    R_PRODTABLE = 4
    R_SORTGROUPS = 8
    R_GROUPFORECAST = 16
    R_PLANETFORECAST = 32
    R_SHIPTYPEFORECAST = 64
    R_ROUTESFORECAST = 128
    R_SORTED = 256           'used by the rateRace function
    R_COMPRESS = 512         'Compress turn reports before sending then
    R_GPLUS = 1024
    R_MACHINEREPORT = 2048   'Players ask for a "machine" turn report
    R_BATTLEPROTOCOL = 4096  'Include a battle protocol
    R_CRLF = 8192            'Make every /n a cr lf (not used, can be reused)
    R_XMLREPORT = 16384      'Players ask for an xml turn report
    R_DEAD = 32768           'player is no longer active
    R_TXTREPORT = 65536      'is a text report wanted?
End Enum

Public Property Get Flag(ByVal uFlag As RaceFlag) As Boolean
    Flag = ((flags And uFlag) = uFlag)
End Property

