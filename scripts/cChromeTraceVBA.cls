Option Explicit

'/**
'*this function allows the use of the chrome tracer from apps script
'* implements a subset of https://docs.google.com/document/d/1CvAClvFfyA5R-PhYUmn5OOQtYMH4h6I0nSsKchNAySU/edit#heading=h.yr4qxyxotyw
'*/

'/**
'* a trace manager
'* @return {ChromeTrace} self
'*/
Private pEnums As cJobject
Private pItems As cJobject
Private pIndex As Long
Private pStartTime As Double

'/**
'* get the enums for this class
'* @return {cJobject} the enums
'*/
Public Function getEnums()
    If (pEnums Is Nothing) Then
        Set pEnums = JSONParse("{'ENUMS':" & _
            "{'SCOPES' : {'GLOBAL': 'g','PROCESS':'p', 'THREAD':'t' }," & _
                "'PH': {" & _
                    "'DURATION': {'BEGIN':'B','END':'E'}," & _
                    "'COMPLETE':'X'," & _
                    "'INSTANT':'i'," & _
                    "'COUNTER':'C'," & _
                    "'ASYNC': {'NESTABLESTTART':'b','NESTABLEINSTANT':'n','NESTABLEEND':'e'}," & _
                    "'FLOW': {'BEGIN':'s','STEP':'t','END':'f'}," & _
                    "'SAMPLE':'P'," & _
                    "'OBJECT': {'CREATED':'N','SNAPSHOT':'O','DESTROYED':'D'}," & _
                    "'METADATA':'M'," & _
                    "'MEMORYDUMP': { 'GLOBAL':'V', 'PROCESS':'v' }" & _
            "}}}")
    End If
    
    Set getEnums = pEnums
End Function

  
'/**
'* get all the items
'* @return {cJobject} array of trace items that have been detected
'*/
Public Function getItems()
    Set getItems = pItems
End Function
 
'/**
' * record a begin event
' * @param {string} name the item name
' * @param {cjobject} options template
' * @return {cjobject} the updated event
' */
Public Function start(name As String, Optional options As cJobject = Nothing) As cJobject
    Set start = addEvent(name, options, getEnums().cValue("ENUMS.PH.DURATION.BEGIN"))
End Function

'/**
' * record an end event - need to change name to endEvent as end is reserved in vba
' * @param {string} name the item name
' * @param {cJobject} options template
' * @return {cJobject} the updated event
' */
Public Function finish(name As String, Optional options As cJobject = Nothing) As cJobject
    Set finish = addEvent(name, options, getEnums().cValue("ENUMS.PH.DURATION.END"))
End Function

'/**
' * record an counter event
' * @param {string} name the item name
' * @param {cJobject} options template
' * @return {cJobject} the updated event
' */
Public Function counter(name As String, Optional options As cJobject = Nothing) As cJobject
    Set counter = addEvent(name, options, getEnums().cValue("ENUMS.PH.COUNTER"))
End Function

'/**
' * record an instant event
' * @param {string} name the item name
' * @param {cJobject} options template
' * @return {cJobject} the updated event
' */
Public Function instant(name As String, Optional options As cJobject = Nothing) As cJobject
    Set instant = addEvent(name, options, getEnums().cValue("ENUMS.PH.INSTANT"))
End Function

'/**
'* add an event
'* @param {string} name the item name
'* @param {cJobject} options template
'* @param {string} ph the type of event
'* @return {cJobject} the updated event
'*/
Public Function addEvent(name As String, Optional options As cJobject = Nothing, Optional ph As String = "") As cJobject
    Dim newOptions As cJobject, cv As cJobject
    Set newOptions = makeOptions(name, options)
    newOptions.add "ph", ph
    
    '// if it's a new name then we add ititem
    If (pItems.child(newOptions.toString("name")) Is Nothing) Then
        pItems.add(newOptions.toString("name")).addArray
    End If
    
    '// this is for resolving sort ties
    pIndex = pIndex + 1

    Set addEvent = pItems.child(newOptions.toString("name")).add(, chromeTraceEvent(newOptions, pIndex))
End Function
'/**
'* a chrome event
'* @param {cjobject} options template
'* @return {cjobject} the updated event
'*/
Private Function chromeTraceEvent(options As cJobject, index As Long) As cJobject
    '// not going to bother to make this a specic class - just use a cjobect
    Dim je As cJobject, op As String, jo As cJobject
    If (isSomething(options)) Then
        op = options.stringify
    End If

    '// set default options
    Set je = optionsExtend(op, _
    "{'name': 'traceEvent','cat': 'chrometrace','ph': '" & getEnums().cValue("ENUMS.PH.DURATION.BEGIN") & "'" & _
    ",'ts':" & Round(1000 * (tinyTime() - pStartTime)) & ",'pid': 1, 'tid': 1,'args': null, 's' : '" & getEnums().cValue("ENUMS.SCOPES.GLOBAL") & "'")

    Set jo = New cJobject
    With jo.init(Nothing)
        jo.add "index", pIndex
        jo.add "trace", je
        jo.add "sortKey", je.cValue("ts")
    End With
    
    Set chromeTraceEvent = jo
End Function

Private Function cleanPropertyName(name As String) As String
    cleanPropertyName = makeKey(name)
End Function

   
Private Function makeOptions(name As String, options As cJobject) As cJobject
    Dim jo As cJobject
    
    If (options Is Nothing) Then
        Set jo = New cJobject
        jo.init Nothing
    Else
        Set jo = options
    End If
    
    If (name = vbNullString) Then name = jo.toString("name")
    jo.add "name", cleanPropertyName(name)
    
    Set makeOptions = jo
End Function


'/**
'* dump the result
'* @param {string} drivePath the folder
'* @param {string} fileName the file name
'* @param {string} name the item - null for all
'* @return {ChromeTraceEvent} the updated event
'*/
Public Function dump(drivePath As String, Optional filename As String = vbNullString, Optional name As String = vbNullString)
     Dim item As cJobject, content As String, job As cJobject, joc As cJobject, arrJob As cJobject, jod As cJobject
     name = cleanPropertyName(name)

    '// need to play around with the objects to sort them ny timestamp
    Set item = New cJobject

    With item.init(Nothing).addArray
        For Each job In pItems.children
        
            ' its possible to just select one item
            If ((name <> vbNullString And makeKey(name) = job.key) Or name = vbNullString) Then
                For Each joc In job.children
                    With .add
                        For Each jod In joc.getObject().children
                            .add jod.key, jod.value
                        Next jod
                    End With
                Next joc
            End If
        Next job
    End With

     
   ' sort out the filename
   filename = conditionalAssignment(filename <> vbNullString, filename, _
    conditionalAssignment(name <> vbNullString, name, "chrometracing") & ".json")
     
   '// sort in to ts order
    item.sortByValue

    '// set the content
    Set arrJob = item.extractAndBlowArrayOfJobs

    content = arrJob.stringify
    
    'write to file
    writeToFolderFile drivePath, filename, content
    
    '// clean up

    arrJob.tearDown
    item.tearDown
    
End Function
Private Sub Class_Initialize()
    Set pItems = New cJobject
    pItems.init Nothing
    pIndex = 0
    pStartTime = tinyTime()
End Sub

Private Sub Class_Terminate()
    If (isSomething(pItems)) Then
        pItems.tearDown
    End If
End Sub