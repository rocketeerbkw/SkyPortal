<%
	'-------------------------------------------------------------
	' DataCache 1.1
	' http://www.2enetworx.com/dev/articles/caching.asp
	'
	' File: cache.asp
	' Description: Cache Engine for Recorset serialization
	' Written By Hakan Eskici on Oct 21, 2003
	'
	' You may use the code for any purpose
	' But re-publishing is discouraged.
	' See License.txt for additional information	
	'-------------------------------------------------------------

Class SqlCache
 
	Public ConnString
	 
	Private stCache
	 
	Private Sub Class_Initialize()
		Set stCache = Server.CreateObject("AdoDB.Stream")
		stCache.Type = 1 'adTypeBinary
	end sub
	 
	Private Sub Class_Terminate()
		Set stCache = Nothing
	End Sub
	
	Public Function GetRecordset(SQL)
		Dim conn, rsTemp, sBuffer
		Set rsTemp = Server.CreateObject("AdoDB.Recordset")
		
		'Check whether the requested SQL is in cache
		If dDataCache.Exists(SQL) then
			'Already in cache
			stCache.Open
			sBuffer = dDataCache.Item(SQL)
			
			stCache.Write sBuffer
			stCache.Position = 0
			
			rsTemp.Open stCache
			
			stCache.Close
		Else
			'Not in cache
			Set connTemp = Server.CreateObject("ADODB.Connection")
			connTemp.Open ConnString
			rsTemp.CursorLocation = 2  'adUseServer
			rsTemp.Open SQL, connTemp, 3, 1, 1
			'3, 1, 1 = adOpenStatic, adLockReadOnly, adCmdText
			
			'Save recordset in the stream in ADTG format
			rsTemp.Save stCache, 0 'adPersistADTG
		
			'Serialize the stream as text
			sBuffer = stCache.Read(-1) 'adReadAll
			
			Application.Lock
		
			'Use the SQL statement as the unique item key
			dDataCache.Add SQL, sBuffer
		
			Application.Unlock
			stCache.Close
		End If
		
		Set GetRecordset = rsTemp
		
		Set rsTemp = nothing
	End Function
	
	Public Sub Remove(SQL) 
		Application.Lock 
		If dDataCache.Exists(SQL) Then  
			dDataCache.Remove SQL
		End If 
		Application.UnLock
		
	End Sub
	
	Public Sub RemoveAll 
		Application.Lock
		dDataCache.RemoveAll
		Application.Unlock 
	End Sub

	Public Sub ExpireTable(TableName) 
		'Get keys as an array
		aKeys = dDataCache.Keys
		
		'How many entries?
		iCount = dDataCache.Count
		
		'Enumerate cache
		For i = 0 To iCount - 1
		
			'Get SQL
			sSQL = aKeys(i)
			
			sItemTable = ""
			
			'Find the underlying table
			aSQL = Split(sSQL, " ")
			For iSQL = 0 To UBound(aSQL) - 1
				'Look for FROM keyword
				If LCase(aSQL(iSQL)) = "from" Then
					'Found it
					sItemTable = aSQL(iSQL + 1)
					Exit For
				End if
			Next
			
			If LCase(TableName) = LCase(sItemTable) Then
				'Remove the entry
				dDataCache.Remove sSQL
			End If 
		Next 
	End Sub 
End Class
%>