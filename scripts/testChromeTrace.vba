Option Explicit
Function testTrace()
    Dim trace As cChromeTraceVBA, i As Long, b64 As String
    Const LOOPSIZE = 1000
    Const text = "abcdefghijklmnop"
    Dim args As cJobject
    
    '// set this up once at the beginning
    Set args = JSONParse("{'args':{'count':0,'random':0}}")

    '// this is the tracing object
    Set trace = New cChromeTraceVBA
    
    '// start an overall trace
    trace.start "b64"
        
        '// strat a nested trace
        trace.start "encode"
        For i = 1 To LOOPSIZE
            b64 = Base64Encode(text)
            
            '// put out some sample values
            args.child("args.count").value = i
            args.child("args.random").value = Rnd() * LOOPSIZE
            trace.counter "countencode", args
        Next i
        
        '// finish nested trace
        trace.finish "encode"
        
        '// start another nested trace
        'trace.start "decode"
        For i = 1 To LOOPSIZE
            b64 = Base64Decode(text)
            
            '// some more sample values
            args.child("args.count").value = i
            args.child("args.random").value = Rnd() * LOOPSIZE
            trace.counter "countdecode", args
        Next i
        
        '// finish nested trace
        trace.finish "decode"
        
    '// finish overall trace
    trace.finish "b64"
    
    '// write default file name, current directory
    trace.dump "./"
    
    '// clean up
    args.tearDown
    
End Function