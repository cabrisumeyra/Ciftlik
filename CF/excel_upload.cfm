<cfparam name="form.FileName" default="">
<cfset FileName = form.FileName>

<cffile action = "upload" fileField = "file_11" destination = "#expandPath("ExcellDosyalar")#"  nameConflict = "Overwrite" result="resul"> 	


<cfspreadsheet action="read" src="#expandPath("ExcellDosyalar/excelim.xls")#" name="sheet" >



<cfoutput>

    <cfloop from="2" to="#sheet.rowcount#" index="row">
     <!--- multi row selection (edit based on excel sheet col relationship) --->
    <cfquery datasource="Ciftlik" result="rResponse">
        INSERT INTO Ciftlik.dbo.HayvanDetay ( HayvanNo, CiftlikID, HayvanCesidi, HayvanTuru, Cinsiyet, Mahsul) 
        VALUES (
            <cfqueryparam cfsqltype="cf_sql_integer" value="#SpreadsheetGetCellValue(sheet, row, 1)#" />,
            <cfqueryparam cfsqltype="cf_sql_integer" value="#SpreadsheetGetCellValue(sheet, row, 2)#" />,
            <cfqueryparam cfsqltype="cf_sql_nvarchar" value="#SpreadsheetGetCellValue(sheet, row, 3)#" />,
            <cfqueryparam cfsqltype="cf_sql_nvarchar" value="#SpreadsheetGetCellValue(sheet, row, 4)#" />,
            <cfqueryparam cfsqltype="cf_sql_nvarchar" value="#SpreadsheetGetCellValue(sheet, row, 5)#" />,
            <cfqueryparam cfsqltype="cf_sql_nvarchar" value="#SpreadsheetGetCellValue(sheet, row, 6)#" />
        )
    </cfquery>
</cfloop>
<h1>Helal olsun, tebrikler !</h1>
</cfoutput>











<script>
    $('#file_11').change(function(e){
        var fileName = e. target. files[0]. name;
        $("#FileName").val(fileName)
    });

</script>
