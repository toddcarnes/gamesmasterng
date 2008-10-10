VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ShipType"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Public ShipType As ShipType
Public FromPlanet As Planet
Public WherePlanet As Planet

Public ShipTypeID As Long
Public Number As Long
Public Drive As Single
Public Weapons As Single
Public Shields As Single
Public Cargo As Single
Public LoadType As String
Public Load As Single
Public FromPlanetID As Long
Public WherePlanetID As Long
Public Dist As Single
Public Ships As Long
Public FleetID As Long
