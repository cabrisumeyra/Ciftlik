<cfquery name="HayvanSayi" datasource="Ciftlik">
    SELECT
    HayvanNo, HayvanCesidi
    FROM 
    HayvanDetay
</cfquery>
<!--Hayvan Sayılarını gösteren tablo -->
<div class="container p-3 my-3 col-md-3 h-100">
    <span>
        <table widht="99%">
            <td align="left" height="40"><h4>Hayvan Sayım Bilgileri</h4></td>
    </table>
    <table class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
        <tr>
            <td><b>No</b></td>
            <td><b>Hayvan Çeşidi</b></td>
        </tr>
        <cfoutput query="HayvanSayi" startrow="#HayvanSayi.recordCount - 5#">
            <tr>
                <td>#HayvanNo#</td>
                <td>#HayvanCesidi#</td>
            </tr>
        </cfoutput>
    </table>
    </span>
</div>

<cfform method="post" enctype="multipart/form-data" action="excel_aktarma.cfm">
    <input type="file" name="file_11" id="file_11">
    <input type="hidden" name="FileName" id="FileName">
    <input type="hidden" name="is_submitted" id="is_submitted">
    <input type="submit">
</cfform>

<cfif isDefined("form.FileName") and len(form.FileName)>
    <cffile action="upload" fileField="file_11" destination="#expandPath("ExcellDosyalar")#" nameConflict="Overwrite" result="resul"> 	
    <cfspreadsheet action="read" src="#expandPath("ExcellDosyalar/#form.FileName#")#" name="sheet">
    <cfoutput>
        <cfloop from="2" to="#sheet.rowcount#" index="row">
            <!--- multi row selection (edit based on excel sheet col relationship) --->
            <cfquery datasource="Ciftlik" result="rResponse">
                INSERT INTO Ciftlik.dbo.HayvanDetay (HayvanNo, CiftlikID, HayvanCesidi, HayvanTuru, Cinsiyet, Mahsul) 
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
    </cfoutput>
</cfif>

<script>
    $('#file_11').change(function(e){
        var fileName = e.target.files[0].name;
        $("#FileName").val(fileName);
    });
</script>
