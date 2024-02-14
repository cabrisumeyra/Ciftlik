<cfcomponent>
    <cffunction name="CiftlikEkleme" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="CiftlikAd" type="string" required="true">
        <cfargument name="KapaliAlan" type="integer" required="true">
        <cfargument name="AcikAlan" type="integer" required="true">
        <cfargument name="SahiplikDurumu" type="string" required="true">
        <cfargument name="HasatTarih" type="date" required="true">
        <cfargument name="SiloKapasite" type="integer" required="true">
        <cfargument name="CalisanSayisi" type="integer" required="true">
        <cfargument name="BinaSigortaDurumu" type="string" required="true">
        <cfargument name="RecordDate" type="date" required="true">

        <cfquery name="CiftlikInto" datasource="Ciftlik" >
            INSERT INTO CiftlikBilgileri (CiftlikID, CiftlikAd, KapaliAlan, AcikAlan, SahiplikDurumu, HasatTarih, SiloKapasite, CalisanSayisi, BinaSigortaDurumu, RecordDate)
            VALUES
            ( 
                <cfqueryparam value="#CiftlikID#"       cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#CiftlikAd#"       cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#KapaliAlan#"      cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#AcikAlan#"        cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#SahiplikDurumu#"  cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#HasatTarih#"      cfsqltype="CF_SQL_DATE">,
                <cfqueryparam value="#SiloKapasite#"    cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#CalisanSayisi#"   cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#BinaSigortaDurumu#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#RecordDate#"      cfsqltype="CF_SQL_DATE">
            )
        </cfquery>
        <cflocation url="/CiftlikBilgi.cfm" addtoken="false">
    </cffunction>
</cfcomponent>
 