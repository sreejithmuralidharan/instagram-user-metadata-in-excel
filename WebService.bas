Attribute VB_Name = "WebService"

Function APIGET(url)

    Dim WinHttp As Object
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    WinHttp.SetAutoLogonPolicy (0)
    WinHttp.Open "GET", url
    WinHttp.Send
    
    Dim success As Boolean
    success = WinHttp.WaitForResponse(5)
    If Not success Then
        MsgBox ("DOWNLOAD FAILED!")
        Exit Function
    End If
    
    APIGET = WinHttp.responseText


End Function


Function InstagramMeta(handle, meta)

url = "https://www.instagram.com/" & handle & "/?__a=1"
jsonRespose = APIGET(url)

Set Parsed = JsonConverter.ParseJson(jsonRespose)

Select Case meta
    Case "FullName"
        InstagramMeta = Parsed("graphql")("user")("full_name")
    Case "FollowedBy"
        InstagramMeta = Parsed("graphql")("user")("edge_followed_by")("count")
    Case "Following"
        InstagramMeta = Parsed("graphql")("user")("edge_follow")("count")
    Case "BusinessAccount"
        InstagramMeta = Parsed("graphql")("user")("is_business_account")
    Case "JoinedRecently"
        InstagramMeta = Parsed("graphql")("user")("is_joined_recently")
     Case "Category"
        InstagramMeta = Parsed("graphql")("user")("business_category_name")
     Case "Private"
        InstagramMeta = Parsed("graphql")("user")("is_private")
     Case "Verified"
        InstagramMeta = Parsed("graphql")("user")("is_verified")
End Select

End Function


Sub InstagramMetaTest()

url = "https://www.instagram.com/benlcollins/?__a=1"
jsonRespose = APIGET(url)

Set Parsed = JsonConverter.ParseJson(jsonRespose)

MsgBox

End Sub
