<cfcomponent>
    <cffunction name="UretimGuncelleme" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="MahsulCesidi" type="string" required="true">
        <cfargument name="ToplanmaDurumu" type="string" required="true">
        <cfargument name="GunlukUretimAdeti" type="integer" required="true">
        <cfargument name="HasarliMahsulAdedi" type="date" required="true">

        <cfquery name="uretimekleme" datasource="Ciftlik" >
            INSERT INTO Uretim (CiftlikID, MahsulCesidi, ToplanmaDurumu, GunlukUretimAdeti, HasarliMahsulAdedi, RecordDate)
            VALUES (
                <cfqueryparam value="#arguments.CiftlikID#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#arguments.MahsulCesidi#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#arguments.ToplanmaDurumu#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#arguments.GunlukUretimAdeti#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#arguments.HasarliMahsulAdedi#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#arguments.RecordDate#" cfsqltype="CF_SQL_DATE">
            )
        </cfquery>

        <cflocation url="/Uretim.cfm" addtoken="false">
    </cffunction>
</cfcomponent>

 