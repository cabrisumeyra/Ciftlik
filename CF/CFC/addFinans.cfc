<cfcomponent>
    <cffunction name="FinansEkleme" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="HayvanCesidi" type="string" required="true">
        <cfargument name="YemAlinan" type="integer" required="true">
        <cfargument name="YemHarcanan" type="date" required="true">
        <cfargument name="IlacAlinan" type="integer" required="true">
        <cfargument name="VeterinerMasraf" type="integer" required="true">
        <cfargument name="MahsulCesidi" type="string" required="true">
        <cfargument name="SatisAdetUcreti" type="integer" required="true">
        <cfargument name="GunlukSatisUcreti" type="integer" required="true">
        <cfargument name="GunlukSatisMiktari" type="integer" required="true">
        <cfargument name="Aciklama" type="string" required="true">
        <cfargument name="RecordDate" type="date" required="true">

        <cfquery name="CiftlikInto" datasource="Ciftlik" >
            INSERT INTO Finans (CiftlikID, HayvanCesidi, YemAlinan, YemHarcanan, IlacAlinan, VeterinerMasraf, MahsulCesidi, SatisAdetUcreti, GunlukSatisUcreti, GunlukSatisMiktari, Aciklama, RecordDate)
            VALUES (
                <cfqueryparam value="#CiftlikID#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#HayvanCesidi#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#YemAlinan#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#YemHarcanan#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#IlacAlinan#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#VeterinerMasraf#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#MahsulCesidi#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#SatisAdetUcreti#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#GunlukSatisUcreti#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#GunlukSatisMiktari#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#Aciklama#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#RecordDate#" cfsqltype="CF_SQL_DATE">
                )
        </cfquery>
        <cflocation url="/Finans.cfm" addtoken="false">
    </cffunction>
</cfcomponent>


