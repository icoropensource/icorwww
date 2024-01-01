<%
dim w3
w3=True
asql=GetInputValue("pivotsql")
if asql<>"" and IsObject(dPageParams) then
   dim filename,adeleteifexists,fs3,w4
   filename=GetInputValue("datasourcefilenamepersist") 'for persisted recordsets
   adeleteifexists=GetInputValue("datasourcefiledeleteifexists") 'remove old data?
   if filename="" then
      Application.Lock
      atmpfilecnt=Dession("TmpServerFileCounter")
      Dession("TmpServerFileCounter")=1+Dession("TmpServerFileCounter")
      Application.UnLock
      filename="xmlsrcdata_" & CStr(Dession.SessionID) & "_" & atmpfilecnt & ".bin"
   end if
   dPageParams("datasource")="/icormanager/output/" & filename
   dPageParams("datasourcefilename")=filename
   fname=Application("DefaultOutputPath") & "\" & filename
   Set fs3 = Server.CreateObject("Scripting.FileSystemObject")
   w4=fs3.FileExists(fname)
   if w4 and adeleteifexists<>"" then
      on error resume next
      fs3.DeleteFile fname,True
      w4=False
      on error goto 0
   end if
   if not w4 then
      'on error resume next
      'fs3.DeleteFile fname,True
      Set rsd3 = Server.CreateObject("ADODB.Recordset")
      rsd3.ActiveConnection = cn
      rsd3.CursorType = 3 'adUseClient
      rsd3.LockType = 3 'adLockOptimistic
      rsd3.Source=asql
      rsd3.Open
      if rsd3.RecordCount>0 then
         on error resume next
         application.lock
         rsd3.Save fname, 0
         application.unlock
         if rsd3.State<>0 then 'adStateClosed
            rsd3.Close
         end if
         set rsd3=Nothing
         on error goto 0
      else
         w3=False
      end if
   end if
   Set fs3=Nothing
end if
if w3 then 'BB1 - BLOK DLA CALEJ STRONY

%>
<BR>
<TABLE><CAPTION class='objectsviewcaption'><%=GetInputValue("pivotcaption")%></CAPTION><TR><TD>

<script LANGUAGE=JScript>
var aobjectchartname="",aobjectchartid="",aowcver=0,aowcvermin=0;
function insertOWC() {
   var oOWC;
   lowc=[
      ['OWC11.PivotTable','0002E55A-0000-0000-C000-000000000046','0002E559-0000-0000-C000-000000000046','0002E55D-0000-0000-C000-000000000046','OWC11.ChartSpace',11,0]
     ,['OWC10.PivotTable','0002E552-0000-0000-C000-000000000046','0002E551-0000-0000-C000-000000000046','0002E556-0000-0000-C000-000000000046','OWC10.ChartSpace',10,2]
    // ,['OWC10.PivotTable','0002E542-0000-0000-C000-000000000046','0002E541-0000-0000-C000-000000000046','0002E546-0000-0000-C000-000000000046','OWC10.ChartSpace',10,1]
     ,['OWC.PivotTable','0002E520-0000-0000-C000-000000000046','0002E510-0000-0000-C000-000000000046','0002E500-0000-0000-C000-000000000046','OWC.Chart',9,0]
   ];
   for (i in lowc) {
      try {
         oOWC=new ActiveXObject(lowc[i][0]);
         document.write('<div id=myinfo<%=GetInputValue("pivotid")%>><h1>Proszê czekaæ, trwa ³adowanie danych...</h1></div><div id=mypivot<%=GetInputValue("pivotid")%> style="display:none;"><object classid="clsid:'+lowc[i][1]+'" id="PTable<%=GetInputValue("pivotid")%>"></object></div>')
<%
if GetInputValue("pivotshowdetails")="1" then
%>       document.write('<FONT SIZE=2 COLOR=IndianRed><DIV ID=divDesc<%=GetInputValue("pivotid")%>></DIV></FONT><OBJECT classid="clsid:'+lowc[i][2]+'" id="SSheet<%=GetInputValue("pivotid")%>" style="visibility:hidden;"></OBJECT>')
<%
end if
%>
         aobjectchartname=lowc[i][4];
         aobjectchartid=lowc[i][3];
         aowcver=lowc[i][5];
         aowcvermin=lowc[i][6];
         return;
      } catch(e){
      }   
      oOWC=null;
   }  
  alert("Zainstaluj OWC!");
}
insertOWC();
</script>
</TD></TR>
<%


Application.Lock
atmpfilecounter=CStr(Dession("TmpClientFileCounter"))
Dession("TmpClientFileCounter")=1+Dession("TmpClientFileCounter")
if Dession("TmpClientFileCounter")>10 then
   Dession("TmpClientFileCounter")=1
end if
Application.UnLock

if GetInputValue("showchart")="99" then ' WYKRESY WYLACZONE!!!

%>
<TR><TD>
<HR>

<script language="javascript">
if (aobjectchartname!="") {
   var oOWC = new ActiveXObject(aobjectchartname);
   document.write('<div id=mychart<%=GetInputValue("pivotid")%> style="display:none;"><OBJECT CLASSID="clsid:'+aobjectchartid+'" id="ChartSpace<%=GetInputValue("pivotid")%>"></OBJECT></div>')
}
</script>

</TD></TR>
<TR><TD>
<div id=mychartoptions<%=GetInputValue("pivotid")%> style="display:none;">
<%
   If True Then
%>
<SELECT id=lstChartType<%=GetInputValue("pivotid")%>>
   <OPTION selected value=0>Kolumnowy Grupowany</OPTION>
   <OPTION value=1>Skumulowany Kolumnowy</OPTION>
   <OPTION value=2>100% Skumulowany Kolumnowy</OPTION>
   <OPTION value=3>S³upkowy Grupowany</OPTION>
   <OPTION value=4>Skumulowany S³upkowy</OPTION>
   <OPTION value=5>100% Skumulowany S³upkowy</OPTION>
   <OPTION value=6>Liniowy</OPTION>
   <OPTION value=8>Skumulowany Liniowy</OPTION>
   <OPTION value=10>100% Skumulowany Liniowy</OPTION>
   <OPTION value=7>Liniowy ze Znacznikami</OPTION>
   <OPTION value=9>Skumulowany Liniowy ze Znacznikami</OPTION>
   <OPTION value=11>100% Skumulowany Liniowy ze Znacznikami</OPTION>
   <OPTION value=12>Liniowy Wyg³adzony</OPTION>
   <OPTION value=14>Skumulowany Liniowy Wyg³adzony</OPTION>
   <OPTION value=16>100% Skumulowany Liniowy Wyg³adzony</OPTION>
   <OPTION value=13>Liniowy Wyg³adzony ze Znacznikami</OPTION>
   <OPTION value=15>Skumulowany Liniowy Wyg³adzony ze Znacznikami</OPTION>
   <OPTION value=17>100% Skumulowany Liniowy Wyg³adzony ze Znacznikami</OPTION>
   <OPTION value=18>Ko³owy</OPTION>
   <OPTION value=19>Ko³owy Rozsuniêty</OPTION>
   <OPTION value=20>Ko³owy Skumulowany</OPTION>
   <OPTION value=21>Rozproszony</OPTION>
   <OPTION value=25>Rozproszony z Liniami</OPTION>
   <OPTION value=24>Rozproszony z Liniami i Znacznikami</OPTION>
   <OPTION value=26>Rozproszony Powierzchniowy</OPTION>
   <OPTION value=23>Rozproszony z Wyg³adzonymi Liniami</OPTION>
   <OPTION value=22>Rozproszony z Wyg³adzonymi Liniami i Znacznikami</OPTION>
   <OPTION value=27>B¹belkowy</OPTION>
   <OPTION value=28>B¹belkowy z Liniami</OPTION>
   <OPTION value=29>Powierzchniowy</OPTION>
   <OPTION value=30>Powierzchniowy Skumulowany</OPTION>
   <OPTION value=31>100% Powierzchniowy Skumulowany</OPTION>
   <OPTION value=32>Pierœcieniowy</OPTION>
   <OPTION value=33>Pierœcieniowy Rozsuniêty</OPTION>
   <OPTION value=34>Radarowy z Liniami</OPTION>
   <OPTION value=35>Radarowy z Liniami i Znacznikami</OPTION>
   <OPTION value=36>Wype³niony Radarowy</OPTION>
   <OPTION value=37>Radarowy z Wyg³adzonymi Liniami</OPTION>
   <OPTION value=38>Radarowy z Wyg³adzonymi Liniami i Znacznikami</OPTION>
   <OPTION value=39>Gie³dowy Maks-Min-Zamkniêcie</OPTION>
   <OPTION value=40>Gie³dowy Otwarcie-Maks-Min-Zamkniêcie</OPTION>
   <OPTION value=41>Polarny</OPTION>
   <OPTION value=42>Polarny z Liniami</OPTION>
   <OPTION value=43>Polarny z Liniami i Znacznikami</OPTION>
   <OPTION value=44>Polarny z Wyg³adzonymi Liniami</OPTION>
   <OPTION value=45>Polarny z Wyg³adzonymi Liniami i Znacznikami</OPTION>
</SELECT>
<%
   Else
%>
<SELECT id=lstChartType<%=GetInputValue("pivotid")%>>
   <OPTION selected value=0>Clustered Column</OPTION>
   <OPTION value=1>Stacked Column</OPTION>
   <OPTION value=2>100% Stacked Column</OPTION>
   <OPTION value=3>Clustered Bar</OPTION>
   <OPTION value=4>Stacked Bar</OPTION>
   <OPTION value=5>100% Stacked Bar</OPTION>
   <OPTION value=6>Line</OPTION>
   <OPTION value=8>Stacked Line</OPTION>
   <OPTION value=10>100% Stacked Line</OPTION>
   <OPTION value=7>Line with Markers</OPTION>
   <OPTION value=9>Stacked Line with Markers</OPTION>
   <OPTION value=11>100% Stacked Line with Markers</OPTION>
   <OPTION value=12>Smooth Line</OPTION>
   <OPTION value=14>Stacked Smooth Line</OPTION>
   <OPTION value=16>100% Stacked Smooth Line</OPTION>
   <OPTION value=13>Smooth Line with Markers</OPTION>
   <OPTION value=15>Stacked Smooth Line with Markers</OPTION>
   <OPTION value=17>100% Stacked Smooth Line with Markers</OPTION>
   <OPTION value=18>Pie</OPTION>
   <OPTION value=19>Pie Exploded</OPTION>
   <OPTION value=20>Stacked Pie</OPTION>
   <OPTION value=21>Scatter</OPTION>
   <OPTION value=25>Scatter with Lines</OPTION>
   <OPTION value=24>Scatter with Markers and Lines</OPTION>
   <OPTION value=26>Filled Scatter</OPTION>
   <OPTION value=23>Scatter with Smooth Lines</OPTION>
   <OPTION value=22>Scatter with Markers and Smooth Lines</OPTION>
   <OPTION value=27>Bubble</OPTION>
   <OPTION value=28>Bubble with Lines</OPTION>
   <OPTION value=29>Area</OPTION>
   <OPTION value=30>Stacked Area</OPTION>
   <OPTION value=31>100% Stacked Area</OPTION>
   <OPTION value=32>Doughnut</OPTION>
   <OPTION value=33>Exploded Doughnut</OPTION>
   <OPTION value=34>Radar with Lines</OPTION>
   <OPTION value=35>Radar with Lines and Markers</OPTION>
   <OPTION value=36>Filled Radar</OPTION>
   <OPTION value=37>Radar with Smooth Lines</OPTION>
   <OPTION value=38>Radar with Smooth Lines and Markers</OPTION>
   <OPTION value=39>High-Low-Close</OPTION>
   <OPTION value=40>Open-High-Low-Close</OPTION>
   <OPTION value=41>Polar</OPTION>
   <OPTION value=42>Polar with Lines</OPTION>
   <OPTION value=43>Polar with Lines and Markers</OPTION>
   <OPTION value=44>Polar with Smooth Lines</OPTION>
   <OPTION value=45>Polar with Smooth Lines and Markers</OPTION>
</SELECT>
<%
   End If
%>
</div>
</TD></TR>
<%
end if
%>
</TABLE>

<SCRIPT LANGUAGE=VBScript>
Function GetTempFileName
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set tfolder = fso.GetSpecialFolder(2) 'TempFolder
   tpath=tfolder.Path
   GetTempFileName = fso.BuildPath(tpath, "pivot_tmp_data_<%=atmpfilecounter%>.bin")
   If fso.FileExists(GetTempFileName) Then
      fso.DeleteFile GetTempFileName,True
   End If
   Set tfolder=Nothing
   Set fso=Nothing
End Function

Function LoadData
   adatasource="<%=GetInputValue("datasource")%>"
   if instr(adatasource,"/")=0 then
      ld=split(window.location.pathname,"/")
      ld(ubound(ld))=adatasource
      adatasource=join(ld,"/")
   end if
   if 0=1 then
      afname=GetTempFileName()
      Set rs1 = CreateObject("ADODB.Recordset")
      rs1.open window.location.protocol & "//" & window.location.host & adatasource
      rs1.Save afname, 0 'adPersistADTG
      if rs1.State<>0 then 'adStateClosed
         rs1.Close
      end if
      set rs1=Nothing
   end if
   PTable<%=GetInputValue("pivotid")%>.ConnectionString = "provider=mspersist"
   Ptable<%=GetInputValue("pivotid")%>.CommandText = window.location.protocol & "//" & window.location.host & adatasource 'afname
   Ptable<%=GetInputValue("pivotid")%>.BackColor="Ivory"
   Ptable<%=GetInputValue("pivotid")%>.MaxHeight=600
'   Ptable<%=GetInputValue("pivotid")%>.MaxWidth=1024

'   PTable<%=GetInputValue("pivotid")%>.AutoFit = False
'   PTable<%=GetInputValue("pivotid")%>.Height = 600
'   PTable<%=GetInputValue("pivotid")%>.Width = 800

   dim stv_ColumnAxisName,stv_RowAxisName
   stv_ColumnAxisName=""
   stv_RowAxisName=""
   
   With Ptable<%=GetInputValue("pivotid")%>.ActiveView
<%if GetInputValue("rowaxis")<>"" then
%>
      .RowAxis.InsertFieldSet .FieldSets("<%=GetInputValue("rowaxis")%>")
      .FieldSets("<%=GetInputValue("rowaxis")%>").Fields(0).Expanded=False
     if .RowAxis.FieldSets("<%=GetInputValue("rowaxis")%>").fields.Count>1 then
        stv_RowAxisName="<%=GetInputValue("rowaxis")%>"
     end if 
<%end if   
if GetInputValue("rowaxis1")<>"" then
%>
      .RowAxis.InsertFieldSet .FieldSets("<%=GetInputValue("rowaxis1")%>")
      .FieldSets("<%=GetInputValue("rowaxis1")%>").Fields(0).Expanded=False   
     if .RowAxis.FieldSets("<%=GetInputValue("rowaxis1")%>").fields.Count>1 then   
        stv_RowAxisName="<%=GetInputValue("rowaxis1")%>"  
     end if   
<%end if
if GetInputValue("rowaxis2")<>"" then
%>
      .RowAxis.InsertFieldSet .FieldSets("<%=GetInputValue("rowaxis2")%>")
      .FieldSets("<%=GetInputValue("rowaxis2")%>").Fields(0).Expanded=False
     if .RowAxis.FieldSets("<%=GetInputValue("rowaxis2")%>").fields.Count>1 then
        stv_RowAxisName="<%=GetInputValue("rowaxis2")%>"
     end if
<%end if
if GetInputValue("rowaxis3")<>"" then
%>
      .RowAxis.InsertFieldSet .FieldSets("<%=GetInputValue("rowaxis3")%>")
      .FieldSets("<%=GetInputValue("rowaxis3")%>").Fields(0).Expanded=False
     if .RowAxis.FieldSets("<%=GetInputValue("rowaxis3")%>").fields.Count>1 then
        stv_RowAxisName="<%=GetInputValue("rowaxis3")%>"
     end if
<%end if
if GetInputValue("columnaxis")<>"" then
%>
      .ColumnAxis.InsertFieldSet .FieldSets("<%=GetInputValue("columnaxis")%>")
      .FieldSets("<%=GetInputValue("columnaxis")%>").Fields(0).Expanded=False
     if .ColumnAxis.FieldSets("<%=GetInputValue("columnaxis")%>").fields.Count>1 then
        stv_ColumnAxisName="<%=GetInputValue("columnaxis")%>"
     end if      
<%end if
if GetInputValue("columnaxis1")<>"" then
%>
      .ColumnAxis.InsertFieldSet .FieldSets("<%=GetInputValue("columnaxis1")%>")
      .FieldSets("<%=GetInputValue("columnaxis1")%>").Fields(0).Expanded=False
     if .ColumnAxis.FieldSets("<%=GetInputValue("columnaxis1")%>").fields.Count>1 then
        stv_ColumnAxisName="<%=GetInputValue("columnaxis1")%>"
     end if
<%end if
if GetInputValue("columnaxis2")<>"" then
%>
      .ColumnAxis.InsertFieldSet .FieldSets("<%=GetInputValue("columnaxis2")%>")
      .FieldSets("<%=GetInputValue("columnaxis2")%>").Fields(0).Expanded=False
     if .ColumnAxis.FieldSets("<%=GetInputValue("columnaxis2")%>").fields.Count>1 then
        stv_ColumnAxisName="<%=GetInputValue("columnaxis2")%>"
     end if
<%end if
if GetInputValue("columnaxis3")<>"" then
%>
      .ColumnAxis.InsertFieldSet .FieldSets("<%=GetInputValue("columnaxis3")%>")
      .FieldSets("<%=GetInputValue("columnaxis3")%>").Fields(0).Expanded=False      
     if .ColumnAxis.FieldSets("<%=GetInputValue("columnaxis3")%>").fields.Count>1 then   
        stv_ColumnAxisName="<%=GetInputValue("columnaxis3")%>"  
     end if   
<%end if
%>   
   
     
   if stv_ColumnAxisName<>"" then
<%if GetInputValue("excludecolumnaxisfield1")="1" then
%>
        .ColumnAxis.FieldSets(stv_ColumnAxisName).fields(0).IsIncluded=False
<%
end if
if GetInputValue("excludecolumnaxisfield2")="1" then
%>
        .ColumnAxis.FieldSets(stv_ColumnAxisName).fields(1).IsIncluded=False
<%
end if
if GetInputValue("excludecolumnaxisfield3")="1" then
%>
        .ColumnAxis.FieldSets(stv_ColumnAxisName).fields(2).IsIncluded=False
<%
end if
if GetInputValue("excludecolumnaxisfield4")="1" then
%>
        .ColumnAxis.FieldSets(stv_ColumnAxisName).fields(3).IsIncluded=False
<%end if
%>  
   end if
   
   if stv_RowAxisName<>"" then
<%if GetInputValue("excluderowaxisfield1")="1" then
%>
        .RowAxis.FieldSets(stv_RowAxisName).fields(0).IsIncluded=False
<%
end if
if GetInputValue("excluderowaxisfield2")="1" then
%>
        .RowAxis.FieldSets(stv_RowAxisName).fields(1).IsIncluded=False
<%
end if
if GetInputValue("excluderowaxisfield3")="1" then
%>
        .RowAxis.FieldSets(stv_RowAxisName).fields(2).IsIncluded=False
<%
end if
if GetInputValue("excluderowaxisfield4")="1" then
%>
        .RowAxis.FieldSets(stv_RowAxisName).fields(3).IsIncluded=False
<%end if
%>  
   end if  

      dim atotalfunction
      atotalfunction="<%=GetInputValue("dataaxistotalfunction")%>"
      if atotalfunction="" then
         atotalfunction="sum"
      end if
      aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionSum
      if atotalfunction="sum" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionSum
      elseif atotalfunction="count" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionCount
      elseif atotalfunction="max" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionMax
      elseif atotalfunction="min" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionMin
      end if
      Ptable<%=GetInputValue("pivotid")%>.AllowDetails = False
      .DataAxis.InsertTotal .AddTotal("<%=GetInputValue("dataaxis")%>",.FieldSets("<%=GetInputValue("dataaxisfieldset")%>").Fields(0), aplfunction) 
      dim anumberformat
      anumberformat="<%=GetInputValue("dataaxisnumberformat")%>"
      if anumberformat="" then
         anumberformat="[BLACK] #,##0;[RED]-#,##0"
      end if
      .Totals("<%=GetInputValue("dataaxis")%>").NumberFormat = anumberformat
<%
for i=1 to 11
   if GetInputValue("dataaxis" & CStr(i))<>"" then
%>
'      msgbox "i = <%=i%> "
      atotalfunction="<%=GetInputValue("dataaxistotalfunction" & CStr(i))%>"
      if atotalfunction="" then
         atotalfunction="sum"
      end if

      aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionSum
      if atotalfunction="sum" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionSum
      elseif atotalfunction="count" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionCount
      elseif atotalfunction="max" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionMax
      elseif atotalfunction="min" then
         aplfunction=Ptable<%=GetInputValue("pivotid")%>.Constants.plFunctionMin
      end if
      Ptable<%=GetInputValue("pivotid")%>.AllowDetails = False
      .DataAxis.InsertTotal .AddTotal("<%=GetInputValue("dataaxis" & CStr(i))%>",.FieldSets("<%=GetInputValue("dataaxisfieldset" & CStr(i))%>").Fields(0), aplfunction) 
      anumberformat="<%=GetInputValue("dataaxisnumberformat" & CStr(i))%>"
      if anumberformat="" then
         anumberformat="[BLACK] #,##0;[RED]-#,##0"
      end if
      .Totals("<%=GetInputValue("dataaxis" & CStr(i))%>").NumberFormat = anumberformat
<%
   end if
next

if GetInputValue("SubTotals")<>"" then
%> 
      .Fieldsets("<%=GetInputValue("rowaxis")%>").Fields(0).SubTotals(0) = <%=GetInputValue("SubTotals")%>
<%
end if
if GetInputValue("FilterAxisField")<>"" then
%> 
      .FilterAxis.Label.Visible = True
      .FilterAxis.InsertFieldSet .FieldSets("<%=GetInputValue("FilterAxisField")%>")
<%
else
%> 
      .FilterAxis.Label.Visible = False
<%
end if
%> 
   
      Ptable<%=GetInputValue("pivotid")%>.DisplayToolbar = True
      Ptable<%=GetInputValue("pivotid")%>.AllowPropertyToolbox = True
      .TitleBar.Visible = False
      .TotalBackColor = "Ivory"
      .FieldLabelBackColor = "DarkBlue"
      .FieldLabelFont.Color = "White"

      PTable<%=GetInputValue("pivotid")%>.AllowDetails = False
      'PTable.DisplayToolbar = False
      'PTable.AllowGrouping = False

<%
if GetInputValue("columnaxis")<>"" then
%>   
      .ColumnAxis.Fieldsets(0).Fields(0).SubTotalBackColor = "LightSteelBlue"
<%
end if
%>   
      .RowAxis.Fieldsets(0).Fields(0).SubTotalBackColor = "LightSteelBlue"
      '.MemberBackColor = "Lavender"
   End With
   
<%
if GetInputValue("showchart")="99" then  ' WYKRESY WYLACZONE!!!
%>
   
      set c = ChartSpace<%=GetInputValue("pivotid")%>.Constants
      ChartSpace<%=GetInputValue("pivotid")%>.Clear 
      
      Set oChart = ChartSpace<%=GetInputValue("pivotid")%>.Charts.Add
      oChart.type = c.chChartTypeColumnClustered
      
      oChart.Interior.Color="Ivory"
      
      oChart.HasLegend = true
      oChart.Legend.Position = c.chLegendPositionRight
      
      oChart.HasTitle = True
      oChart.Title.Font.Bold = True
      oChart.Title.Caption="<%=GetInputValue("charttitle")%>"
      
      Set ChartSpace<%=GetInputValue("pivotid")%>.DataSource = PTable<%=GetInputValue("pivotid")%>
      oChart.SetData c.chDimSeriesNames, 0, c.chPivotColumns
      oChart.SetData c.chDimCategories, 0, c.chPivotRows
      oChart.SetData c.chDimValues, 0, 0
   
   
      document.all("lstChartType<%=GetInputValue("pivotid")%>").value=<%=GetInputValue("chartid")%>
      oChart.Type=CLng(lstChartType<%=GetInputValue("pivotid")%>.value)

      document.all("mychart<%=GetInputValue("pivotid")%>").style.display=""
      document.all("mychartoptions<%=GetInputValue("pivotid")%>").style.display=""
<%
end if
%>
   document.all("myinfo<%=GetInputValue("pivotid")%>").style.display="none"
   document.all("mypivot<%=GetInputValue("pivotid")%>").style.display=""
<%
if GetInputValue("pivotshowdetails")="1" then
%>

   ' *** Jesli wystepuje blad z wyswietlaniem danych *** 
   'MsgBox "Dane s¹ gotowe do wyœwietlenia."
   call SetupSheet<%=GetInputValue("pivotid")%>
<%
end if
%>
End Function

<%
if GetInputValue("showchart")="99" then ' WYKRESY WYLACZONE!!!
%>
Sub lstChartType<%=GetInputValue("pivotid")%>_onChange()
    Dim oChart
    Dim c
    Dim nChartType
    set c = ChartSpace<%=GetInputValue("pivotid")%>.Constants
    Set oChart = ChartSpace<%=GetInputValue("pivotid")%>.Charts(0)
    nChartType = CLng(lstChartType<%=GetInputValue("pivotid")%>.value)
    oChart.Type = nChartType
End Sub
<%
end if
%>

<%
if GetInputValue("pivotshowdetails")="1" then
%>
Sub PTable<%=GetInputValue("pivotid")%>_SelectionChange()
   Dim oSelection
   Set oSelection = PTable<%=GetInputValue("pivotid")%>.Selection
   If TypeName(oSelection)="PivotAggregates" Then
      Dim oAggregate
      Set oAggregate = oSelection.Item(0)
      divDesc<%=GetInputValue("pivotid")%>.InnerHTML = ""
      call ShowDetails<%=GetInputValue("pivotid")%>(oAggregate.Cell)
   Else
      SSheet<%=GetInputValue("pivotid")%>.Cells.ClearContents
      SSheet<%=GetInputValue("pivotid")%>.style.visibility = "hidden"
'      divDesc<%=GetInputValue("pivotid")%>.InnerHTML = "Wybrana pozycja to: <B>" & typename(oSelection) & "</B><BR><BR><BR>Aby pokazaæ dane szczegó³owe, zaznacz <U>jedn¹ wartoœæ</U> w obszarze danych tabeli."
      divDesc<%=GetInputValue("pivotid")%>.InnerHTML = ""
   End If
End Sub
<%
end if
%>

window.setTimeout GetRef("LoadData")

</SCRIPT>
<%
end if 'BB1 - BLOK DLA CALEJ STRONY
%>