Attribute VB_Name = "Save_Program_Settings"

  Option Explicit

' ------------------------------------------------------------------------------------
' Saves all control values including the window size and position
' for the form used by this program.
'
' The registry entries are stored in:
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings

' ------------------------------------------------------------------------------------
' Determine window parameters for this computer

  Const WinHeight = "@Height@"
  Const WinWidth = "@Width@"
  Const WinTop = "@Top@"
  Const WinLeft = "@Left@"

' -----------------------------------------------
' Remember the current program interface settings

  Public Sub Store_Program_Settings _
 (ByVal Program_Name As String, Program_Form_Name As Form, _
  Optional Program_Form_Name_Position As Boolean, _
  Optional Form_Control_Error As Long)

  Dim Kounter As Long
  Dim Control_Name As String
  Dim Control_Value As Control
  Dim Control_Value_Error As Long
    
  On Error GoTo ERROR_HANDLER

' ------------------------------------
' Store current form position settings

  If Program_Form_Name_Position = True Then

     SaveSetting Program_Name, Program_Form_Name.Name, _
     WinHeight, Program_Form_Name.Height

     SaveSetting Program_Name, Program_Form_Name.Name, _
     WinWidth, Program_Form_Name.Width

     SaveSetting Program_Name, Program_Form_Name.Name, _
     WinTop, Program_Form_Name.Top

     SaveSetting Program_Name, Program_Form_Name.Name, _
     WinLeft, Program_Form_Name.Left

  End If

' ------------------------------------------------------
' Read and store the parameters of all the controls used
' on this form.
    
  For Each Control_Value In Program_Form_Name.Controls
        
  On Error Resume Next

  With Control_Value

  Kounter = .Index
  If Err.Number = 343 Then
     Control_Name = .Name
     Err.Clear
  Else
     Control_Name = .Name & ":" & Trim(CStr(Kounter))
  End If

  End With
        
  On Error GoTo ERROR_HANDLER

  If TypeOf Control_Value Is TextBox Then
     SaveSetting Program_Name, Program_Form_Name.Name, _
     Control_Name, Control_Value.Text

  ElseIf TypeOf Control_Value Is CheckBox _
      Or TypeOf Control_Value Is OptionButton Then
      SaveSetting Program_Name, Program_Form_Name.Name, _
      Control_Name, Control_Value.Value

   ElseIf TypeOf Control_Value Is ListBox _
   Or TypeOf Control_Value Is ComboBox Then
      SaveSetting Program_Name, Program_Form_Name.Name, _
      Control_Name, Make_Control_List(Control_Value, Control_Value_Error)

      If Control_Value_Error <> 0 Then Err.Raise Control_Value_Error

  End If
    
  Next Control_Value

  Set Control_Value = Nothing

ERROR_HANDLER:
    
  Form_Control_Error = Err.Number

  End Sub

' --------------------------------------------------------------
' Routine to restore previous program settings from the registry

  Public Sub Recall_Program_Settings _
 (ByVal Program_Name As String, Program_Form_Name As Form, _
  Optional Program_Form_Name_Position As Boolean, _
  Optional Form_Control_Error As Long)

  Dim Kounter As Long
  Dim Control_Name As String
  Dim Control_Value As Control
  Dim Control_Value_Error As Long
    
  On Error GoTo ERROR_HANDLER
    
' Recall previously stored form window position settings
  If Program_Form_Name_Position = True Then

     Program_Form_Name.Height = GetSetting(Program_Name, _
     Program_Form_Name.Name, WinHeight, _
     Program_Form_Name.Height)

     Program_Form_Name.Width = GetSetting(Program_Name, _
     Program_Form_Name.Name, WinWidth, Program_Form_Name.Width)

     Program_Form_Name.Top = GetSetting(Program_Name, _
     Program_Form_Name.Name, WinTop, Program_Form_Name.Top)

     Program_Form_Name.Left = GetSetting(Program_Name, _
     Program_Form_Name.Name, WinLeft, Program_Form_Name.Left)
  
  End If

' --------------------------------------------------------
' Read the parameter values of each of the form's controls
  For Each Control_Value In Program_Form_Name.Controls
        
  On Error Resume Next

  With Control_Value
       Kounter = .Index

  If Err.Number = 343 Then
     Control_Name = .Name
     Err.Clear
  Else
     Control_Name = .Name & ":" & Trim(CStr(Kounter))
  End If

  End With

  On Error GoTo ERROR_HANDLER

  If TypeOf Control_Value Is TextBox Then
        Control_Value.Text = GetSetting(Program_Name, _
        Program_Form_Name.Name, Control_Name, Control_Value.Text)

  ElseIf TypeOf Control_Value Is CheckBox _
      Or TypeOf Control_Value Is OptionButton Then
         Control_Value.Value = GetSetting(Program_Name, _
         Program_Form_Name.Name, Control_Name, Control_Value.Value)

  ElseIf TypeOf Control_Value Is ListBox _
      Or TypeOf Control_Value Is ComboBox Then
         Set_Control_Values Control_Value, GetSetting(Program_Name, _
         Program_Form_Name.Name, Control_Name, ""), Control_Value_Error

      If Control_Value_Error <> 0 Then Err.Raise Control_Value_Error
  End If
    
  Next Control_Value

  Set Control_Value = Nothing

ERROR_HANDLER:
    
    Form_Control_Error = Err.Number

End Sub

  Private Function Make_Control_List _
 (Form_Control As Control, Form_Control_Error As Long) As String

' Construct a string consisting of a list of all the
' controls contained in the form.

  Dim List_of_Controls As Variant
  Dim Kounter As Long
   
  On Error GoTo ERROR_HANDLER
     List_of_Controls = ""
    
  If TypeOf Form_Control Is ListBox _
     Or TypeOf Form_Control Is ComboBox Then

        With Form_Control
             For Kounter = 0 To .ListCount - 1
                 If List_of_Controls <> "" Then
                    List_of_Controls = List_of_Controls & vbVerticalTab
                 End If
                 List_of_Controls = List_of_Controls & .List(Kounter)
             Next Kounter
        End With

  End If

  Make_Control_List = List_of_Controls
    
ERROR_HANDLER:

  Form_Control_Error = Err.Number

  End Function

' ---------------------------------------------------------------
' Read list of control item values and write them back into the
' original form controls.

  Private Sub Set_Control_Values _
 (Form_Control As Control, Control_List_String As String, _
  Form_Control_Error As Long)

  Dim Control_List_Array As Variant
  Dim Kounter As Integer
    
  On Error GoTo ERROR_HANDLER

  If TypeOf Form_Control Is ListBox _
     Or TypeOf Form_Control Is ComboBox Then
    
     Form_Control.Clear
     Control_List_Array = Split(Control_List_String, vbVerticalTab)
        
     If IsArray(Control_List_Array) Then

        For Kounter = LBound(Control_List_Array) _
            To UBound(Control_List_Array)

            Form_Control.AddItem Control_List_Array(Kounter)
        Next Kounter

     Else
        Form_Control.AddItem Control_List_Array
     End If
        
  End If
    
ERROR_HANDLER:

  Form_Control_Error = Err.Number

  End Sub


