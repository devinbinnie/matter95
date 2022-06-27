VERSION 5.00
Begin VB.Form MainChatForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matter95"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "MainChatForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer PostUpdateTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   0
   End
   Begin VB.ListBox TeamList 
      Height          =   645
      ItemData        =   "MainChatForm.frx":030A
      Left            =   240
      List            =   "MainChatForm.frx":030C
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox PostTextBox 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   6255
   End
   Begin VB.CommandButton SubmitButton 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox PostViewTextBox 
      Height          =   4215
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.ListBox ChannelList 
      Height          =   2985
      ItemData        =   "MainChatForm.frx":030E
      Left            =   240
      List            =   "MainChatForm.frx":0310
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Posts"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Write a post"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Channels:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Teams:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "MainChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colMyTeams As Collection
Dim colMyChannels As Collection
Dim objPosts As Dictionary
Dim dictChannelIndexToChannelId As Dictionary
Dim sLastPostId As String
Dim sCurrentChannelId As String

Private Sub Form_Load()
    PostViewTextBox.Locked = True
    PostUpdateTimer.Enabled = False
    MainChatForm.Caption = "Matter95 - " & MattermostClient.sServerURL
    Set dictChannelIndexToChannelId = New Dictionary
    
    Set colMyTeams = MattermostClient.GetMyTeams()
    'MsgBox (colMyTeams.Count)
    'MsgBox (JSON.GetParserErrors)
    Dim i As Integer
    For i = 1 To colMyTeams.Count
        Dim sTeamDisplayName As String
        sTeamDisplayName = colMyTeams(i)("display_name")
        TeamList.AddItem sTeamDisplayName

        'MsgBox (colMyTeams(i)("id"))
    Next i
End Sub

Private Sub Form_Unload(iCancel As Integer)
    End
End Sub


Private Sub PostViewTextBox_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub PostTextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SubmitPost
        KeyAscii = 0
    End If
End Sub

Private Sub SubmitPost()
    If PostTextBox.Text = "" Then
        Exit Sub
    End If
    
    Dim sMessageString As String
    sMessageString = PostTextBox.Text
    PostTextBox.Text = ""
    
    MattermostClient.CreatePost sCurrentChannelId, sMessageString
End Sub

Private Sub SubmitButton_Click()
    SubmitPost
End Sub

Private Sub TeamList_Click()
    PostUpdateTimer.Enabled = False
    PostTextBox.Enabled = False
    SubmitButton.Enabled = False
    ChannelList.Clear
    dictChannelIndexToChannelId.RemoveAll
    PostViewTextBox.Text = ""
    sLastPostId = ""
            
    Dim sTeamId As String
    Dim i As Integer
    For i = 0 To colMyTeams.Count - 1
        If TeamList.Selected(i) Then
            sTeamId = colMyTeams(i + 1)("id")
        End If
    Next i
    Set colMyChannels = MattermostClient.GetMyChannelsForTeam(sTeamId)
    
    Dim j As Integer
    For j = 1 To colMyChannels.Count
        Dim sChannelId As String
        Dim sChannelDisplayName As String
        sChannelId = colMyChannels(j)("id")
        sChannelDisplayName = colMyChannels(j)("display_name")
        
        If sChannelDisplayName = "" Then
            If colMyChannels(j)("type") = "D" Then
                Dim sChannelName As String
                sChannelName = colMyChannels(j)("name")
                
                Dim sOtherUserId As String
                Dim arrSplitChannelName() As String
                arrSplitChannelName = Split(sChannelName, "__")
                If arrSplitChannelName(0) = MattermostClient.objCurrentUser("id") Then
                    sOtherUserId = arrSplitChannelName(1)
                Else
                    sOtherUserId = arrSplitChannelName(0)
                End If
                
                Dim colOtherUserResult As Object
                Dim arrUsersIds(1) As String
                arrUsersIds(0) = sOtherUserId
                Set colOtherUserResult = MattermostClient.GetUsersByIds(arrUsersIds)
                
                sChannelDisplayName = "@" & colOtherUserResult(1)("username")
            End If
        End If
        
        dictChannelIndexToChannelId.Add j - 1, sChannelId
        ChannelList.AddItem sChannelDisplayName
    Next j
End Sub

Private Sub ChannelList_Click()
    PostUpdateTimer.Enabled = False
    PostTextBox.Enabled = False
    SubmitButton.Enabled = False
    PostViewTextBox.Text = ""
    sLastPostId = ""
            
    Dim sChannelId As String
    Dim i As Integer
    For i = 0 To colMyChannels.Count - 1
        If ChannelList.Selected(i) Then
            sChannelId = dictChannelIndexToChannelId(i)
        End If
    Next i
    
    sCurrentChannelId = sChannelId
    Set objPosts = MattermostClient.GetPostsForChannel(sChannelId)
    
    Dim j As Integer
    For j = objPosts("order").Count To 1 Step -1
        Dim objPost As Object
        Set objPost = objPosts("posts")(objPosts("order")(j))
        
        Dim sUserName As String
        If Not MattermostClient.dictUsersById.Exists(objPost("user_id")) Then
            Dim arrUsersIds(1) As String
            arrUsersIds(0) = objPost("user_id")
            MattermostClient.GetUsersByIds arrUsersIds
        End If
        
        If MattermostClient.dictUsersById.Exists(objPost("user_id")) Then
            Dim objUser As Object
            Set objUser = MattermostClient.dictUsersById(objPost("user_id"))
            sUserName = objUser("username")
        Else
            sUserName = objPost("user_id")
        End If
        
        PostViewTextBox.Text = PostViewTextBox.Text & sUserName & ": " & objPost("message") & vbNewLine & vbNewLine
        PostViewTextBox.SelStart = &HFFFF&
        sLastPostId = objPost("id")
    Next j
    
    PostUpdateTimer.Enabled = True
    PostTextBox.Enabled = True
    SubmitButton.Enabled = True
End Sub

Private Sub PostUpdateTimer_Timer()
    Dim objNewPosts As Object
    Set objNewPosts = MattermostClient.GetPostsForChannel(sCurrentChannelId, sLastPostId)
    
    If Not objNewPosts("posts").Count = 0 Then
        Dim j As Integer
        For j = objNewPosts("order").Count To 1 Step -1
            Dim objPost As Object
            Set objPost = objNewPosts("posts")(objNewPosts("order")(j))
            
            objPosts.Add objPost("id"), objPost
            
            Dim sUserName As String
            If Not MattermostClient.dictUsersById.Exists(objPost("user_id")) Then
                Dim arrUsersIds(1) As String
                arrUsersIds(0) = objPost("user_id")
                MattermostClient.GetUsersByIds arrUsersIds
            End If
            
            If MattermostClient.dictUsersById.Exists(objPost("user_id")) Then
                Dim objUser As Object
                Set objUser = MattermostClient.dictUsersById(objPost("user_id"))
                sUserName = objUser("username")
            Else
                sUserName = objPost("user_id")
            End If
            
            PostViewTextBox.Text = PostViewTextBox.Text & sUserName & ": " & objPost("message") & vbNewLine & vbNewLine
            PostViewTextBox.SelStart = &HFFFF&
            sLastPostId = objPost("id")
        Next j
    End If
End Sub
