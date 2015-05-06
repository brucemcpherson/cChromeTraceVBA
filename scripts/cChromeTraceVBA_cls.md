# VBA Project: **cChromeTraceVBA**
## VBA Module: **[cChromeTraceVBA](/scripts/cChromeTraceVBA.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (cChromeTraceVBA) was automatically created on 5/6/2015 11:59:05 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cChromeTraceVBA

---
VBA Procedure: **getEnums**  
Type: **Function**  
Returns: **Variant**  
Return description: **the enums**  
Scope: **Public**  
Description: **get the enums for this class**  

*Public Function getEnums()*  

**no arguments required for this procedure**


---
VBA Procedure: **getItems**  
Type: **Function**  
Returns: **Variant**  
Return description: **array of trace items that have been detected**  
Scope: **Public**  
Description: **get all the items**  

*Public Function getItems()*  

**no arguments required for this procedure**


---
VBA Procedure: **start**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the updated event**  
Scope: **Public**  
Description: **record a begin event**  

*Public Function start(name As String, Optional options As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||the item name
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|template


---
VBA Procedure: **finish**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the updated event**  
Scope: **Public**  
Description: **record an end event - need to change name to endEvent as end is reserved in vba**  

*Public Function finish(name As String, Optional options As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||the item name
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|template


---
VBA Procedure: **counter**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the updated event**  
Scope: **Public**  
Description: **record an counter event**  

*Public Function counter(name As String, Optional options As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||the item name
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|template


---
VBA Procedure: **instant**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the updated event**  
Scope: **Public**  
Description: **record an instant event**  

*Public Function instant(name As String, Optional options As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||the item name
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|template


---
VBA Procedure: **addEvent**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the updated event**  
Scope: **Public**  
Description: **add an event**  

*Public Function addEvent(name As String, Optional options As cJobject = Nothing, Optional ph As String = "") As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||the item name
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|template
ph|String|True| ""|the type of event


---
VBA Procedure: **chromeTraceEvent**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the updated event**  
Scope: **Private**  
Description: **a chrome event**  

*Private Function chromeTraceEvent(options As cJobject, index As Long) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||template
index|Long|False||


---
VBA Procedure: **cleanPropertyName**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function cleanPropertyName(name As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||


---
VBA Procedure: **makeOptions**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeOptions(name As String, options As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **dump**  
Type: **Function**  
Returns: **Variant**  
Return description: **the updated event**  
Scope: **Public**  
Description: **dump the result**  

*Public Function dump(drivePath As String, Optional filename As String = vbNullString, Optional name As String = vbNullString)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
drivePath|String|False||the folder
filename|String|True| vbNullString|the file name
name|String|True| vbNullString|the item - null for all


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Terminate**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Terminate()*  

**no arguments required for this procedure**
