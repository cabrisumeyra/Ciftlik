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




<cffile action = "upload" fileField = "file_11" destination = "#expandPath("ExcellDosyalar")#"  nameConflict = "Overwrite" result="resul"> 	
<cfoutput>Dosya_Adi: #resul.ServerFile#</cfoutput>
<cfspreadsheet action="read" src="#expandPath("ExcellDosyalar/#resul.ServerFile#")#" name="sheet" >

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









