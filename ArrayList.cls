VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Dim m_list As Collection

Private Sub Class_Initialize()
    Set m_list = New Collection
End Sub

Public Sub Add(Item, Optional Key, Optional Before, Optional After)
    Call m_list.Add(Item, Key, Before, After)
End Sub

Public Sub Remove(Index)
    Call m_list.Remove(Index)
End Sub

Public Sub Clear()
    Call m_list.Clear
End Sub

Public Function Count()
    Count = m_list.Count
End Function

Public Function Contains(Key)
    Contains = m_list.Contains(Key)
End Function

Public Function Item(Index)
Attribute Item.VB_UserMemId = 0
    Item = m_list.Item(Index)
End Function

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_list.[_NewEnum]
End Function
