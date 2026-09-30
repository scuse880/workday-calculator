VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Order As Long
Public Name As String
Public Birthdate As String
Public NeisPersonId As String
Public WorkStatusRecords As Collection
Public TripRecords As Collection
Public VacationWorkRecords As Collection
Public PersonnelChangeStart As Date
Public PersonnelChangeEnd As Date

Private Sub Class_Initialize()
    Set WorkStatusRecords = New Collection
    Set TripRecords = New Collection
    Set VacationWorkRecords = New Collection
End Sub
