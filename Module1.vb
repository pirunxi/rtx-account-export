Module Module1

    Sub Main(ByVal args() As String)
        Dim detail
        If args.Length = 0 Then
            detail = False
        Else
            detail = True
        End If
        Dim form = New Form1(detail)
        form.Run()
        Console.ReadKey()
    End Sub


    Public Class Form1
        Private detail
        Private myBuddy
        Private objApi
        Private objker
        Private objBuddy
        Private objApp
        Private objHelper

        Public Sub New(de)
            detail = de
        End Sub


        Public Sub Run()
            objApi = CreateObject("RTXClient.RTXAPI")
            objker = objApi.GetObject("KernalRoot")
            objBuddy = objker.RTXBuddyManager
            objApp = objApi.GetObject("AppRoot")
            objHelper = objApp.GetAppObject("RTXHelper")
            myBuddy = objker.RTXGroupManager

            Dim items As New ArrayList

            GetBuddyList(10000001, "", items)
            For Each line In items
                Console.WriteLine(line)
            Next

        End Sub

        Private Sub GetBuddyList(ByVal depid As Integer, ByVal tabs As String, ByRef lines As ArrayList)
            Dim buddy = myBuddy.Group(depid)
            Dim count = buddy.Groups.count
            Dim indent = tabs + "    "
            For i = 1 To count
                Dim item = buddy.Groups.Item(i)
                lines.Add(tabs + "+++" + item.Name)
                Dim subDepid = item.id
                GetDepUser(subDepid, indent, lines)
                GetBuddyList(subDepid, indent, lines)
            Next
        End Sub

        Private Sub GetDepUser(ByVal depid As Integer, ByVal tabs As String, ByRef lines As ArrayList)
            Dim buddy = myBuddy.Group(depid)
            For i = 1 To buddy.Buddies.count
                Dim account = buddy.Buddies.Item(i).Account
                GetAccountInfo(account, tabs, lines)
            Next
        End Sub

        Private Function GetGender(gender) As String
            If gender = 0 Then
                Return "男"
            Else
                Return "女"
            End If
        End Function

        Private Sub GetAccountInfo(account, tabs, lines)
            Dim buddy = objBuddy.Buddy(account)
            Dim indent = tabs & "    "
            Dim info
            If detail Then
                info = indent & "==========================" & vbCrLf _
                    & indent & "Account:" & buddy.Account & vbCrLf _
                    & indent & "Name:" & buddy.name & vbCrLf _
                    & indent & "Gender:" & GetGender(buddy.Gender) & vbCrLf _
                    & indent & "Tel:" & buddy.Telephone & vbCrLf _
                    & indent & "Mobile:" & buddy.Mobile & vbCrLf _
                    & indent & "Email:" & buddy.Email & vbCrLf _
                    & indent & "Dept:" & objHelper.GetBuddyDept(account) & vbCrLf
                lines.Add(info)
            ElseIf buddy.Telephone.Length >= 8 Or buddy.Mobile.Length >= 11 Then
                info = indent & "==========================" & vbCrLf _
                    & indent & "Account:" & buddy.Account & vbCrLf _
                    & indent & "Name:" & buddy.name & vbCrLf _
                    & indent & "Gender:" & GetGender(buddy.Gender) & vbCrLf _
                    & indent & "Tel:" & buddy.Telephone & vbCrLf _
                    & indent & "Mobile:" & buddy.Mobile & vbCrLf _
                    & indent & "Email:" & buddy.Email & vbCrLf _
                    & indent & "Dept:" & objHelper.GetBuddyDept(account) & vbCrLf
                lines.Add(info)
            End If

        End Sub
    End Class

End Module
