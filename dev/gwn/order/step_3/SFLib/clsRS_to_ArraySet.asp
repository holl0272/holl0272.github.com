<%
'--------------------------------------------------------------------------------------------
' C_ArraySet.cls
' Tom Kelleher Consulting, Inc.
'
' www.tkelleher.com
' kelleher@tkelleher.com
'--------------------------------------------------------------------------------------------
' Use free of charge, but think of me from time to time.
'--------------------------------------------------------------------------------------------
' This class object is to be used as a lightweight alternative to the standard recordset.
' It uses the ADO recordset's .GetRows() method to create an internal array of the 
' recordset's contents.  Thereafter, it provides the same .MoveNext, .MoveLast, .MovePrev,
' .EOF and .BOF methods as a recordset.  This makes it intuitive for developers familiar
' with the ADO methods to continue to use the same syntax with this object.
'
' The benefits of this object include the extremely small "weight" of it.  An empty ADO
' recordset can require as much as 100KB of memory, while this uses far less. This is 
' related to the huge number of infrequently used methods and object-model branches
' in the ADO recordset.  This object-model is completely flat.  So after loading the
' contents of the recordset into this object, we recommend immediately closing and
' Nothing-izing the recordset, to release its resources.
'--------------------------------------------------------------------------------------------
' 
' 
' Usage:
' 		Dim objArraySet, rs, i
'		
' 		Set objArraySet = New C_ArraySet
' 		Set rs = Get_StaticRS(SQL, objUser.ClientName)
'		
' 		Call objArraySet.Load(rs)
'		
' 		rs.Close			'Close and Nothing-ize recordset immediately
' 		Set rs = Nothing
'		
'		If Not objArraySet.EOF Then
'			
'			Do While Not objArraySet.EOF
'				Response.Write objArraySet.Fields("Last_Name") & ", " & objArraySet.Fields("First_Name") & "<br>"
'				objArraySet.MoveNext
'			Loop
' 
'		End If
'		
'		Set objArraySet = Nothing
'		
' Notes:
'		
'		- The .RecordCount property works no matter what type of recordset 
'		  provides the original data. (Depending on the cursor-type, certain 
'		  recordsets only return -1 as the value of .RecordCount)
'		
'		- The .MovePrev and .MoveFirst functions work even if the original
'		  recordset had a forward-only cursor
'		  
'		- Because VBscript classes don't support default properties, we can't
'		  use this shorthand -- objArraySet("First_Name") -- the way we can for 
'		  recordsets -- rs("First_Name").  With a recordset, this is actually
'		  shorthand for rs.Fields("First_Name").Value.  With this object, we
'		  have to use the .Fields() method, as shown above.
'		  
'--------------------------------------------------------------------------------------------

Class C_ArraySet
	
	Private bBOF				'Boolean to indicate if we are at beginning of record array
	Private bEOF				'Boolean to indicate if we are at end of record array
	Private intArrayUBound		'The UBound of the internal data array (# of rows - 1)
	Private intFieldCount		'The number of fields (columns) in the array
	Private intCursorLocation	'The number of the current record (row)
	
	Private arDataset			'The array containing the records
	Private arFieldNames		'A one-dimensional array containing the field names
	
	Private bRsLoaded			'Boolean indicating whether a recordset has been loaded
	Private bReadyForUse		'Boolean indicating if the object is ready for use
	
	Private ERR_NO_RS			'A pseudo-constant; VBscript doesn't do constants
	Private ERR_NO_RECORDS

	'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	' ADDED 11/16/04
	'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	'1) Added a global bDebug variable
	'2) Increment a global counter each time this class is initiated. This line prevents
	'the global variables "fld_" from conflicting with eachother using multiple instances of this class simotaneously...

	Private bDebug
	Private sInstanceName

	Public Property Let SetInstanceName(arg_InstanceName)
		sInstanceName = arg_InstanceName
	End Property
	'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	
	Private Sub Class_Initialize()

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		' ADDED 11/16/04
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		bDebug = False
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

		intArrayUBound = -1		
		intFieldCount = -1
		bReadyForUse = False
		bRsLoaded = False
		bBOF = True
		bEOF = True
		
		arDataset = Array()
		
		ERR_NO_RS 		= "RsTable not ready for use. Submit a valid recordset object, using the LoadRS() method."
		ERR_NO_RECORDS	= "RsTable contains no records."
		
	End Sub


	Private Sub Class_Terminate()

	End Sub
	
	'-------------------------------------------------
	
	Public Sub Load(ByRef rs)
	
		'Loads a recordset ByRef.  The KillRS parameter
		'is a Boolean; if True, the recordset is closed/Nothingized.
		
		
		bRsLoaded = True
		intArrayUBound = -1
		intCursorLocation = -1
		bReadyForUse = False
		
		If TypeName(rs) <> "Recordset" Then
			RaiseError "The Load() function requires a recordset."
			Exit Sub
		End If
		
        On Error Resume Next

        If Not rs.EOF Then
            arDataset = rs.GetRows()
        Else
            intArrayUBound = -1
            intFieldCount = 0
            bBOF = True
            bEOF = True
            intCursorLocation = 0
            Exit Sub
        End If

        If Err.Number <> 0 Then

            'DropKill "BAM"

            intArrayUBound = -1
            intFieldCount = 0
            bBOF = True
            bEOF = True
            intCursorLocation = 0

            'If KillRS Then KillRecordset rs
            Exit Sub

        End If

        On Error GoTo 0

        intArrayUBound = UBound(arDataset, 2)
        intFieldCount = rs.Fields.Count

        bBOF = False
        bEOF = False
        intCursorLocation = 0


        Dim i, ar()
        ReDim ar(intFieldCount - 1)

        '//////////////////////////////////////////////
        'DEBUG CODE:
        If (bDebug) Then
            wl("<BR><BR>BREAK<BR><BR>")
            'Response.Flush
        End If
        '//////////////////////////////////////////////

        For i = 0 To intFieldCount - 1

            '//////////////////////////////////////////////
            'DEBUG CODE:
            If (bDebug) Then
                wl("[Instance_" & sInstanceName & "_fld_" & rs.fields(i).Name & "]")
                'Response.Flush
            End If
            '//////////////////////////////////////////////

            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            ' ADDED 11/16/04
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            Execute("Instance_" & sInstanceName & "_fld_" & rs.fields(i).Name & " = " & i)
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

            ar(i) = rs.fields(i).Name
        Next

        arFieldNames = ar

        bReadyForUse = True
				
	End Sub
	


	Public Sub MoveNext()
		
		'Move from the current record to the next one.
	
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS

		intCursorLocation = intCursorLocation + 1
		
		If intCursorLocation > intArrayUBound Then
			bEOF = True
		End If
		
	End Sub


	Public Sub MovePrevious()
	
		'Move from the current record to the previous one.
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS
	
		If intCursorLocation < 1 Then
			bBOF = True
			Exit Sub
		Else
			bEOF = False 'KEVIN ADDED THIS CODE 3/12/06 - because otherwise if we reach the EOF, then do moveprevious, it used to think we're still at EOF which is incorrect.
		End If
		
		intCursorLocation = intCursorLocation - 1
	
	End Sub


	Public Sub MoveFirst()
	
		'Move from the current record to the first record.
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS
		
		intCursorLocation = 0
		bBOF = False
		bEOF = False
	
	End Sub


	Public Sub MoveLast()
	
		'Move from the current record to the last record.
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS
		
		intCursorLocation = intArrayUBound
		bBOF = False
		bEOF = False
	
	End Sub


	Public Sub Move(ByVal RecordNumber)
	
		'Move from the current record to the record
		'at the RecordNumber position in the array.
		'RecordNumber is always treated as an absolute
		'value, not a relative value.
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS
		
		If Not IsNumeric(RecordNumber) Then
			RaiseError "Move method requires an integer."
		End If
		
		If RecordNumber > intArrayUBound Then
			intCursorLocation = intArrayUBound
			bEOF = True
		ElseIf RecordNumber < 0 Then
			intCursorLocation = 0
			bBOF = True
		Else
			intCursorLocation = CInt(RecordNumber)
		End If
	
	End Sub


	Public Function EvenRow()

		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS
		
		If bBOF Then
			RaiseError "ArraySet is at BOF position."
			Exit Function
		End If
		
		If bEOF Then
			RaiseError "ArraySet is at EOF position."
			Exit Function
		End If
		
		
		
		EvenRow = IIF( intCursorLocation Mod 2, False, True)

	End Function


	Public Function Fields(ByVal SelectedField)
	
		'Used to pluck the desired array-cell value from
		'the current row/record.  If SelectedField is a
		'string, then the number of the associated column
		'is looked up using the Eval function.  If it is
		'numeric, then the lookup is bypassed.  As with
		'a recordset, using numbers is faster.
		
	
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS

		If bBOF Then
			RaiseError "Cursor is at BOF. Use a .MoveXXXX method to access fields."
			Exit Function
		End If
	
		If bEOF Then
			RaiseError "Cursor is at EOF. Use a .MoveXXXX method to access fields."
			Exit Function
		End If
		
		
	
		'Check if input is a number...
		If IsNumeric(SelectedField) Then
		
			If CInt(SelectedField) > intFieldCount Then
				RaiseError "Field #" & SelectedField & " doesn't exist."
				Exit Function
			End If
			
			Fields = arDataset(CInt(SelectedField), intCursorLocation)
			Exit Function
			
		End If
		
		
		'It must not be a number, so assume it's a string...
		
		Dim strTest, intOrdinal

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		' ADDED 11/16/04
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		strTest = "Instance_" & sInstanceName & "_fld_" & SelectedField
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

		If Len( Eval( strTest ) ) = 0 Then
		
			'No such variable exists, so this is not a valid field name

			RaiseError "No field with the name '" & SelectedField & "' exists."
			Exit Function

		Else
			'Such a variable does exist, so we use Eval to tease out its value
			intOrdinal = Eval( strTest )
		End If
		
		Fields = arDataset(intOrdinal, intCursorLocation)
		
	End Function
	

	Public Function FieldExists(ByVal FieldName)
		
		Dim strTest, intOrdinal
		
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		' ADDED 11/16/04
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		strTest = "Instance_" & sInstanceName & "_fld_" & FieldName
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

		If Len( Eval( strTest ) ) = 0 Then
		
			FieldExists = False

		Else
	
			FieldExists = True

		End If
		
		
	End Function
	

	
	
	
	Public Property Get EOF()
	
		'Indicates the recordset is at the end-of-file position
	
		If Not bRsLoaded Then RaiseError ERR_NO_RS
	
		EOF = bEOF
	
	End Property
	
	
	Public Property Get BOF()

		'Indicates the recordset is at the beginning-of-file position
	
		If Not bRsLoaded Then RaiseError ERR_NO_RS

		BOF = bBOF
	
	End Property
	
	
	Public Property Get IsEmpty()
	
		'A Boolean, which simplifies this standard recordset command:
		'     If rs.EOF and rs.BOF Then...
		'
		'  into the more intuitive:
		'     If objArraySet.IsEmpty...
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS

		IsEmpty = (bBOF And bEOF) 
		
	End Property
	
	
	Public Property Get RecordCount()
	
		'The number of records available.  This number is accurate
		'no matter the type of cursor used to create the original
		'recordset. Like the .NoRecords() method, it can be used
		'to replace:
		'     If rs.EOF and rs.BOF Then...
		'
		'  with the more intuitive:
		'     If objArraySet.RecordCount = 0...
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
	
		RecordCount = intArrayUBound + 1
		
	End Property
	
	
	Public Property Get AbsolutePosition()
	
		'The exact ordinal number of the current record
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
	
		AbsolutePosition = intCursorLocation 
		
	End Property
	
	
	Public Property Get RecordNumber()
	
		'Same as AbsolutePosition, but we add 1...
		
		RecordNumber = AbsolutePosition + 1
		
	End Property
	
	
	Public Property Get IsLastRecord()
		
		'Boolean indicating whether or not this is the
		'last record in the 
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		
		If intCursorLocation = intArrayUBound Then
			IsLastRecord = True
		Else
			IsLastRecord = False
		End If
		
	End Property
	
	
	Public Property Get FieldCount()
		
		'Returns the number of fields.  
		'Similar to rs.Fields.Count
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
	
		FieldCount = intFieldCount
		
	End Property
	
	
	Public Function GetRows()
	
		'Returns the internal array of records.  
		'Similar to rs.GetRows(), though this doesn't
		'take a numeric parameter and so always returns
		'the entire array.
		
		If Not bRsLoaded Then RaiseError ERR_NO_RS
		If intArrayUBound = -1 Then RaiseError ERR_NO_RECORDS

		GetRows = arDataset
		
	End Function
	
	
	
	



	Public Sub Persist(ByVal Method, ByVal UniqueID)
		
		'A new method (incomplete) to persist the arrays
		'to a Session or Application variable or to a file. 
		'This makes it posssible to use a single record table
		'across any number of pages, or even across sessions.
		
		Method = Trim(UCase(Method))
		Dim ar(1), vArray
		ar(0) = arFieldNames
		ar(1) = arDataset
		vArray = ar
	
		Select Case Method
			
				
			Case "APPLICATION"
			
				Application("C_ArraySet:" & UniqueID) = vArray
			
			Case Else
			
				Session("C_ArraySet:" & UniqueID) = vArray
				
		End Select
	
	End Sub
	
	
	Public Sub LoadPersisted(ByVal Method, ByVal UniqueID)
	
		'A new method (incomplete) to load persisted values
		'from file or Session variable to 
	
		Method = Trim(Method)
		Dim ar
	
		Select Case UCase(Method)
			
			Case "APPLICATION"
			
				' pull data from Application variable
				
				ar = Application("C_ArraySet:" & UniqueID)
				
			Case Else
			
				' pull data from Session variable 
				
				ar = Session("C_ArraySet:" & UniqueID)
				
		End Select
	
		If Not IsArray(ar) Then
			RaiseError "Could not find ArraySet with ID of '" & UniqueID & "' persisted to " & Method & "."
			Exit Sub
		End If

	
		arFieldNames = ar(0)
		arDataset = ar(1)				

		bBOF = False
		bEOF = False
		
		intCursorLocation = 0
		
		intArrayUBound = UBound( arDataset, 2 )
		intFieldCount = UBound(arFieldNames) 

		bRsLoaded = True
		bReadyForUse = True

	
	End Sub
	
	
	
	Public Sub DePersist(ByVal Method, ByVal UniqueID)
		
		'Destroys the persisted file or variable.
		
		Method = Trim(UCase(Method))
	
		Select Case Method
			
			Case "APPLICATION"
			
				' Destroy the Application variable
				Application("C_ArraySet:" & UniqueID) = Empty
				
			Case Else
			
				' Destroy the Session variable
				Session("C_ArraySet:" & UniqueID) = Empty
				
		End Select
	
	End Sub
	
	
	Public Function IsPersisted(ByVal Method, ByVal UniqueID)
		
		'Determine if such data was persisted
		
		Method = Trim(UCase(Method))
	
		Select Case Method
			
			Case "APPLICATION"
			
				' Check for the Application variable
				IsPersisted = Not IsEmpty( Application("C_ArraySet:" & UniqueID) )
				
			Case Else
			
				' Check for the Session variable
				IsPersisted = Not IsEmpty( Session("C_ArraySet:" & UniqueID) )
				
		End Select
	
	End Function
	
	
	
	'=========================================
	' Supporting functions
	'=========================================

	Private Sub RaiseError(ByVal Msg)
		Err.Raise vbObjectError + 99999, "C_ArraySet", Msg
	End Sub

	
	Private Sub Drop(ByVal txt)
		Response.Write txt & vbCrLf
	End Sub


End Class

%>