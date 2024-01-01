<%
function GetRSValue(ars,afieldname)
   dim arstype
   GetRSValue=ars.Fields(afieldname).Value
   arsType=ars.Fields(afieldname).Type
   if IsNull(GetRSValue) then
      GetRSValue="-"
   else
      if (arsType=7) or (arsType=133) or (arsType=134) or (arsType=135) then
         GetRSValue=getDateTimeAsStr(GetRSValue)
      end if
      GetRSValue=trim(CStr(GetRSValue))
      if GetRSValue="" then
         GetRSValue="-"
      end if
   end if
end function

sub ExecuteSQLXML(asql)
   Response.ContentType="text/xml"
   atext=""
   call db.Open()
   set rs=db.GetRsClientForwardOnly(asql)
   if not rs.EOF and not rs.BOF then
      atext=GetRSValue(rs,"result")
   else
      Response.End
   end if
   call db.CloseRS(rs)
   call db.Close()
   if atext="-" then
      atext=""
   end if
   Response.Write atext
end sub

Class DBConnection
   Private aCommandTimeout,aConnectionTimeout
   Public cn,rs
   Private Sub Class_Initialize()
      aCommandTimeout=60
      aConnectionTimeout=45
   End Sub
   Private Sub Class_Terminate()
      Me.Close()
   End Sub
   Sub Open()
      On Error Resume Next
      Set Me.cn = Server.CreateObject("ADODB.Connection")
      Me.cn.CursorLocation=2 'adUseServer
      Me.cn.CommandTimeout=aCommandTimeout
      Me.cn.ConnectionTimeout=aConnectionTimeout
      Me.cn.Open Application("CONNECTION_STRING")
      If Err Then
         Set Me.cn = Nothing
         Err.Clear
         Response.End()
      End If
      On Error GoTo 0
   End Sub
   Sub Close()
      On Error Resume Next
      Me.cn.Close
      Set Me.cn = Nothing
      Err.Clear
      On Error Goto 0
   End Sub
   Sub BeginTrans()
      Me.cn.BeginTrans()
   End Sub
   Sub RollBackTrans()
      Me.cn.RollBackTrans()
   End Sub
   Sub CommitTrans()
      Me.cn.CommitTrans()
   End Sub
   Sub CloseRs(ars)
      If IsObject(ars) Then
         On Error Resume Next
         if ars.State<>0 then
            ars.Close
         end if
         Set ars = Nothing
         On Error Goto 0
      End If
   End Sub
   Function GetRs(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorType = 1 'adOpenKeyset
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRs=axrs
   End Function
   Function GetRsClientStatic(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 3 'adUseClient
      axrs.CursorType = 3 'adOpenStatic
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsClientStatic=axrs
   End Function
   Function GetRsClientKeyset(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 3 'adUseClient
      axrs.CursorType = 1 'adOpenKeyset
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsClientKeyset=axrs
   End Function
   Function GetRsClientDynamic(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 3 'adUseClient
      axrs.CursorType = 2 'adOpenDynamic
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsClientDynamic=axrs
   End Function
   Function GetRsClientForwardOnly(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 3 'adUseClient
      axrs.CursorType = 0 'adOpenForwardOnly
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsClientForwardOnly=axrs
   End Function
   Function GetRsServerStatic(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 2 'adUseServer
      axrs.CursorType = 3 'adOpenStatic
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsServerStatic=axrs
   End Function
   Function GetRsServerKeyset(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 2 'adUseServer
      axrs.CursorType = 1 'adOpenKeyset
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsServerKeyset=axrs
   End Function
   Function GetRsServerDynamic(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 2 'adUseServer
      axrs.CursorType = 2 'adOpenDynamic
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsServerDynamic=axrs
   End Function
   Function GetRsServerForwardOnly(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
      axrs.CursorLocation = 2 'adUseServer
      axrs.CursorType = 0 'adOpenForwardOnly
      axrs.LockType = 3 'adLockOptimistic
      axrs.Source = asql
      axrs.Open
      Set GetRsServerForwardOnly=axrs
   End Function
   sub SetItemVersion(aprocname,aoid,aaction)
      Set cmd_ih = Server.CreateObject("ADODB.Command")
      cmd_ih.CommandTimeout=600
      cmd_ih.ActiveConnection = Me.cn
      cmd_ih.ActiveConnection.CursorLocation = 3 'adUseClient
      cmd_ih.CommandType = &H0004 'adCmdStoredProc
      cmd_ih.CommandText = aprocname
      cmd_ih.Parameters.Refresh
      cmd_ih.Parameters("@OID").Value=CStr(aoid)
      cmd_ih.Parameters("@action").Value=CStr(aaction)
      cmd_ih.Execute
      set cmd_ih=nothing
   end sub
End Class
%>