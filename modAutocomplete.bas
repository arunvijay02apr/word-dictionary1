Attribute VB_Name = "modAutocomplete"
Option Explicit

' Inserts the top suggestion shown in the Autocomplete pane at the cursor.
' How it works: Word briefly selects the character before the cursor and then collapses again.
' The Autocomplete pane sees "1 character selected, then empty again" and inserts its top suggestion.

Public Sub AcceptSuggestion()
    Dim t As Single
    If Selection.Type <> wdSelectionIP Then Exit Sub
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    t = Timer
    Do While Timer - t < 0.6                        ' keep Word responsive so the pane can notice
        DoEvents
    Loop
    Selection.Collapse wdCollapseEnd
End Sub

' Run this ONCE to assign Ctrl+Alt+Z to AcceptSuggestion (saved in Normal.dotm).
Public Sub BindAcceptKey()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:="Normal.modAutocomplete.AcceptSuggestion", _
        KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyZ)
    MsgBox "Ctrl+Alt+Z is now assigned to AcceptSuggestion.", vbInformation
End Sub
