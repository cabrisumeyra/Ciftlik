<cfcomponent>
    <cffunction name="Bakimekleme" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="HayvanCesidi" type="string" required="true">
        <cfargument name="IlacAlinan" type="integer" required="true">
        <cfargument name="VeterinerMasraf" type="integer" required="true">
        <cfargument name="VeterinerRandevu" type="date" required="true">
        <cfargument name="Aciklama" type="string" required="true">
        <cfargument name="RecordDate" type="date" required="true">

        <cfquery name="deniyorGuncelleme" datasource="Ciftlik">
            INSERT INTO Finans (CiftlikID, HayvanCesidi, IlacAlinan, VeterinerMasraf, VeterinerRandevu, Aciklama, RecordDate)
        VALUES (
            <cfqueryparam value="#CiftlikID#" cfsqltype="CF_SQL_INTEGER">,
            <cfqueryparam value="#HayvanCesidi#" cfsqltype="CF_SQL_VARCHAR">,
            <cfqueryparam value="#IlacAlinan#" cfsqltype="CF_SQL_INTEGER">,
            <cfqueryparam value="#VeterinerMasraf#" cfsqltype="CF_SQL_INTEGER">,
            <cfqueryparam value="#VeterinerRandevu#" cfsqltype="CF_SQL_DATE">,
            <cfqueryparam value="#Aciklama#" cfsqltype="CF_SQL_VARCHAR">,
            <cfqueryparam value="#RecordDate#" cfsqltype="CF_SQL_DATE">
            )
        </cfquery>
        <cflocation url="/Bakim.cfm" addtoken="false">
    </cffunction>

    <cffunction name="BakimGuncelle" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="HayvanID" type="integer" required="true">
        <cfargument name="HayvanCesidi" type="string" required="true">
        <cfargument name="SaglikDurumu" type="string" required="false">
        <cfargument name="HastalikAdi" type="string" required="false">
        <cfargument name="AsiTarihi" type="date" required="false">
        <cfargument name="TedaviTarihi" type="date" required="false">
        <cfargument name="Aciklama" type="string" required="false">
        <cfargument name="History" type="date" required="false">

        <cfquery name="deniyorGuncelleme" datasource="Ciftlik">
            UPDATE HayvanDetay
            SET 
            SaglikDurumu = <cfqueryparam value="#SaglikDurumu#" cfsqltype="CF_SQL_VARCHAR">,
            HastalikAdi = <cfqueryparam value="#HastalikAdi#" cfsqltype="CF_SQL_VARCHAR">,
            AsiTarihi = <cfqueryparam value="#AsiTarihi#" cfsqltype="CF_SQL_DATE">,
            TedaviTarihi = <cfqueryparam value="#TedaviTarihi#" cfsqltype="CF_SQL_DATE">,
            Aciklama = <cfqueryparam value="#Aciklama#" cfsqltype="CF_SQL_VARCHAR">,
            History = <cfqueryparam value="#History#" cfsqltype="CF_SQL_DATE">
            WHERE HayvanID = <cfqueryparam value="#HayvanID#" cfsqltype="CF_SQL_INTEGER"> 
        </cfquery>
        <cflocation url="/Bakim.cfm" addtoken="false">
    </cffunction>
</cfcomponent>