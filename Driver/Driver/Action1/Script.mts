Datatable.AddSheet "Module"
Datatable.ImportSheet "D:\KeywordDrivenFramework\Organizer\organizer.xlsx",1,"Module"
'Datatable.ImportSheet "D:\KeywordDrivenFrame\Organizer\organizer.xlsx",2,"TestCase"
'Datatable.ImportSheet "D:\KeywordDrivenFrame\Organizer\organizer.xlsx",3,"TestStep"

Services.StartTransaction "tr1"


mrowcount=Datatable.GetSheet("Module").GetRowCount
msgbox mrowcount
For i = 1 To mrowcount Step 1
Datatable.SetCurrentRow(i)
Modexe=Datatable("ModuleExe","Action1") 
'msgbox Modexe
If Modexe="Y" Then
	Modid=Datatable("ModuleID","Action1")       
	msgbox Modid
	trowcount=datatable.GetSheet("Action2").GetRowCount
	msgbox trowcount
	For j=1 To trowcount Step 1
	Datatable.SetCurrentRow(j)
	If Modid=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" Then
	testcaseid=Datatable("TestcaseId","Action2")
	msgbox testcaseid
		trowcount=Datatable.GetSheet("Action 3").GetRowCount
		msgbox trowcount
		For k = 1 to trowcount step 1
		datatable.SetCurrentRow(k)
		If testcaseid=Datatable("TestcaseId","Action 3") Then
		keyword=Datatable("Keyword","Action 3")
		msgbox keyword
		
		select case (keyword)
		
		Case "ln"
		Call Login("john", "hp")	
		
		Case "ca"
		Call Closeapp()	
		
		Case "oo"
		Call openorder()	
		
		Case "uo"
		Call Updateorder()
		
		Case "lnd"
		
		drowcount=datatable.GetSheet("Action4").GetRowCount
		
		For l= 1 To drowcount Step 1
		
			datatable.SetCurrentRow(l)
			
 			Call login(datatable("username","Action4"),datatable("password","Action4"))
			
			Call Closeapp()
		Next
		
		Case "ood"
		
		orrowcount=datatable.GetSheet("Action4").GetRowCount
		For m = 1 To orrowcount Step 1
			datatable.SetCurrentRow(m)
			Call openorder(datatable("orderno","Action4"))
			
		Next
	
		End Select
		
		End If
		Next
		
		End If
		Next
	
		End If
		Next



 @@ hightlight id_;_2083367080_;_script infofile_;_ZIP::ssf4.xml_;_
Services.EndTransaction "tr2"
 @@ hightlight id_;_2081056112_;_script infofile_;_ZIP::ssf9.xml_;_
 @@ hightlight id_;_2228676_;_script infofile_;_ZIP::ssf10.xml_;_
 @@ hightlight id_;_2228676_;_script infofile_;_ZIP::ssf11.xml_;_
 @@ hightlight id_;_2081047376_;_script infofile_;_ZIP::ssf17.xml_;_

