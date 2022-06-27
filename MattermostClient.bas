Attribute VB_Name = "MattermostClient"
Public sLoginToken As String
Public sServerURL As String
Public objCurrentUser As Object

Public dictUsersById As Dictionary
Public colUsers As Collection

Public Sub Login(sUserName As String, sPassword As String)
    Dim sLoginResult As String
    Dim sHeaderResult As String
    If INet.PostRequest(sServerURL & "/api/v4/users/login", "{""device_id"":"""",""login_id"":""" & sUserName & """,""password"":""" & sPassword & """,""token"":""""}", sLoginResult, sHeaderResult) = False Then
        'MsgBox ("Failed to login")
    Else
        Dim aHeaders() As String
        aHeaders = Split(sHeaderResult, vbNewLine)
        Dim i As Integer
        For i = 0 To UBound(aHeaders, 1)
            If InStr(1, aHeaders(i), "Token") = 1 Then
                Dim aTokenHeader() As String
                aTokenHeader = Split(aHeaders(i), ": ")
                sLoginToken = aTokenHeader(1)
            End If
        Next i
        
        Set objCurrentUser = JSON.parse(sLoginResult)
    End If
End Sub

Public Function GetMyTeams() As Object
    Dim sMyTeamsResult As String
    If INet.GetRequest(sServerURL & "/api/v4/users/me/teams", sMyTeamsResult, sLoginToken) = False Then
        'MsgBox ("Failed to get teams")
    Else
        'MsgBox (sMyTeamsResult)
        Set GetMyTeams = JSON.parse(sMyTeamsResult)
    End If
End Function

Public Function GetMyChannelsForTeam(sTeamId As String)
    Dim sMyChannelsResult As String
    If INet.GetRequest(sServerURL & "/api/v4/users/me/teams/" & sTeamId & "/channels", sMyChannelsResult, sLoginToken) = False Then
        'MsgBox ("Failed to get channels")
    Else
        'MsgBox (sMyChannelsResult)
        Set GetMyChannelsForTeam = JSON.parse(sMyChannelsResult)
    End If
End Function

Public Function GetUsersByIds(sUserIds() As String)
    Dim sUsersResult As String
    Dim sHeaderResult As String
    If INet.PostRequest(sServerURL & "/api/v4/users/ids", "[""" & Join(sUserIds, """,""") & """]", sUsersResult, sHeaderResult, sLoginToken) = False Then
        'MsgBox ("Failed to get users")
    Else
        If dictUsersById Is Nothing Then
            Set dictUsersById = New Dictionary
        End If
        'MsgBox (sUsersResult)
        Dim colUserObjects As Collection
        Set colUserObjects = JSON.parse(sUsersResult)
        Dim i As Integer
        For i = 1 To colUserObjects.Count
            Dim sUserId As String
            Dim objUser As Object
            Set objUser = colUserObjects(i)
            sUserId = objUser("id")
            If dictUsersById.Exists(sUserId) Then
                dictUsersById.Remove sUserId
            End If
            dictUsersById.Add sUserId, objUser
        Next i
        Set GetUsersByIds = colUserObjects
    End If
End Function

Public Function GetPostsForChannel(sChannelId As String, Optional sAfter As String)
    Dim sPostsResult As String
    Dim sURL As String
    sURL = sServerURL & "/api/v4/channels/" & sChannelId & "/posts"
    If Not sAfter = "" Then
        sURL = sURL & "?after=" & sAfter
    End If
    If INet.GetRequest(sURL, sPostsResult, sLoginToken) = False Then
        'MsgBox ("Failed to get posts")
    Else
        'MsgBox (sPostsResult)
        Set GetPostsForChannel = JSON.parse(sPostsResult)
    End If
End Function

Function DateToTimeStampMs(d As Date) As Double
    On Error Resume Next
    Dim base As Date
    base = DateSerial(1970, 1, 1)
    DateToTimeStampMs = DateDiff("s", base, d) * 1000
End Function

Public Function CreatePost(sChannelId As String, sMessageString As String)
    Dim dTimeStamp As Double
    dTimeStamp = DateToTimeStampMs(Now)
    Dim sPostResult As String
    Dim sHeaderResult As String
    If INet.PostRequest(sServerURL & "/api/v4/posts", "{""channel_id"":""" & sChannelId & """,""pending_post_id"":""" & objCurrentUser("id") & ":" & dTimeStamp & """,""user_id"":""" & objCurrentUser("id") & """,""update_at"":" & dTimeStamp & ", ""message"":""" & sMessageString & """}", sPostResult, sHeaderResult, sLoginToken) = False Then
        'MsgBox ("Failed to get users")
    Else
        'MsgBox (sPostResult)
    End If
End Function
