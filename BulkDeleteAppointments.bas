Attribute VB_Name = "BulkDeleteAppointments"
'' Copyright (c) 2014 by Joachim Schlosser
' http://www.schlosser.info
'
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without modification,
' are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice,
' this list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation
' and/or other materials provided with the distribution.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'

' Delete all selected calendar items and send the same message on all
Public Sub BulkDeleteAppointments()
    Dim individualItem As Object
    Dim cancelMsg As String
    cancelMsg = InputBox(Prompt:="Your cancel message please. There will be no confirmation.", _
          Title:="ENTER YOUR MESSAGE", Default:="I will be on vacation.")
    If (cancelMsg <> False) Then
        For Each individualItem In Application.ActiveExplorer.Selection
            DeleteItemWithDefaultMessage individualItem, cancelMsg
            'DisplayInfo individualItem
        Next individualItem
    End If
End Sub

' The core function for bulk delete
Sub DeleteItemWithDefaultMessage(oItem, cancelMsg)
    Dim strMessageClass As String
    Dim oAppointItem As Outlook.AppointmentItem
    Dim myMtg As Outlook.MeetingItem
    strMessageClass = oItem.MessageClass
    If (strMessageClass = "IPM.Appointment") Then       ' Only operate on Calendar Entry.
        Set oAppointItem = oItem
        If oAppointItem.Organizer = Outlook.Session.CurrentUser Then  ' If this is my own meeting
            oAppointItem.MeetingStatus = olMeetingCanceled
            oAppointItem.Body = cancelMsg
            oAppointItem.Save
            oAppointItem.Send
        Else                                            ' If I was invited to this meeting
            Set myMtg = oAppointItem.Respond(olMeetingDeclined, True, False)
            myMtg.Body = cancelMsg
            myMtg.Send
        End If
    End If
End Sub

Sub DisplayInfo(oItem As Object)
    Dim strMessageClass As String
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oContactItem As Outlook.ContactItem
    Dim oMailItem As Outlook.MailItem
    Dim oJournalItem As Outlook.JournalItem
    Dim oNoteItem As Outlook.NoteItem
    Dim oTaskItem As Outlook.TaskItem
    
    ' You need the message class to determine the type.
    strMessageClass = oItem.MessageClass
    
    If (strMessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oItem
        MsgBox oAppointItem.Subject
        MsgBox oAppointItem.Start
        MsgBox oAppointItem.Body
    ElseIf (strMessageClass = "IPM.Contact") Then       ' Contact Entry.
        Set oContactItem = oItem
        MsgBox oContactItem.FullName
        MsgBox oContactItem.Email1Address
    ElseIf (strMessageClass = "IPM.Note") Then          ' Mail Entry.
        Set oMailItem = oItem
        MsgBox oMailItem.Subject
        MsgBox oMailItem.Body
    ElseIf (strMessageClass = "IPM.Activity") Then      ' Journal Entry.
        Set oJournalItem = oItem
        MsgBox oJournalItem.Subject
        MsgBox oJournalItem.Actions
    ElseIf (strMessageClass = "IPM.StickyNote") Then    ' Notes Entry.
        Set oNoteItem = oItem
        MsgBox oNoteItem.Subject
        MsgBox oNoteItem.Body
    ElseIf (strMessageClass = "IPM.Task") Then          ' Tasks Entry.
        Set oTaskItem = oItem
        MsgBox oTaskItem.DueDate
        MsgBox oTaskItem.PercentComplete
    End If
    
End Sub

