Attribute VB_Name = "ExtendersModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants and functions used by this program:
Public Const CB_ERR As Long = -1&
Public Const LB_ERR As Long = -1&
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_FINDSTRINGEXACT As Long = &H158&
Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Private Const EM_GETLINECOUNT As Long = &HBA&
Private Const EM_LINEFROMCHAR As Long = &HC9&
Private Const EM_LINELENGTH As Long = &HC1&
Private Const ERROR_SUCCESS As Long = 0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const LB_FINDSTRING As Long = &H18F&
Private Const LB_FINDSTRINGEXACT  As Long = &H1A2&
Private Const MAX_STRING As Long = 65535

Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function SendMessageA Lib "User32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



'The constants and structures used by this program.
Public Const AllItems As Long = -1   'Indicates that all items in a combo/list box are to be searched.

'This structure defines information about the lines in a text box.
Public Type TextLinesStr
   Count As Long           'Contains the number of lines.
   Current As Long         'Contains the current line number.
   FirstVisible As Long    'Contains the number of the first visible line.
   Length As Long          'Contains the current line's length.
End Type

'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

ErrorCode = Err.LastDllError
Err.Clear

   If Not ErrorCode = ERROR_SUCCESS Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_IGNORE_INSERTS Or FORMAT_MESSAGE_FROM_SYSTEM, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCrLf
      MsgBox Message, vbExclamation
   End If
   
   CheckForError = ReturnValue
End Function



'This procedure returns line information for the specified text box.
Public Function GetTextLines(TextBoxO As TextBox) As TextLinesStr
Dim TextLines As TextLinesStr

   With TextLines
      .Count = CheckForError(SendMessageA(TextBoxO.hwnd, EM_GETLINECOUNT, CLng(0), CLng(0)))
      .Current = CheckForError(SendMessageA(TextBoxO.hwnd, EM_LINEFROMCHAR, TextBoxO.SelStart, CLng(0)))
      .FirstVisible = CheckForError(SendMessageA(TextBoxO.hwnd, EM_GETFIRSTVISIBLELINE, CLng(0), CLng(0)))
      .Length = CheckForError(SendMessageA(TextBoxO.hwnd, EM_LINELENGTH, TextBoxO.SelStart, CLng(0)))
   End With
   
   GetTextLines = TextLines
End Function

'This procedure is executed when this program is started.
Public Sub Main()

End Sub


'This procedure returns index of the item containing the specified text in the specified combobox.
Public Function SearchComboBox(ComboBoxO As ComboBox, Text As String, Optional StartIndex As Long = AllItems, Optional AllText As Boolean = False) As Long
Dim Index As Long
Dim SearchType As Long

   If AllText Then SearchType = CB_FINDSTRINGEXACT Else SearchType = CB_FINDSTRING
   Index = CheckForError(SendMessageA(ComboBoxO.hwnd, SearchType, StartIndex, ByVal Text))
   
   SearchComboBox = Index
End Function


'This procedure returns index of the item containing the specified text in the specified listbox.
Public Function SearchListBox(ListBoxO As ListBox, Text As String, Optional StartIndex As Long = AllItems, Optional AllText As Boolean = False) As Long
Dim Index As Long
Dim SearchType As Long

   If AllText Then SearchType = LB_FINDSTRINGEXACT Else SearchType = LB_FINDSTRING
   Index = CheckForError(SendMessageA(ListBoxO.hwnd, SearchType, StartIndex, ByVal Text))
   
   SearchListBox = Index
End Function

