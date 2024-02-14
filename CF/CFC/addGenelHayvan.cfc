<cfcomponent>
    <cffunction name="BeslenmeGuncelleme" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="HayvanTuru" type="string" required="true">
        <cfargument name="BeslenmeDurumu" type="string" required="true">
        <cfargument name="RecordDate" type="date" required="true">

        <cfquery name="beslenme" datasource="Ciftlik">
            UPDATE GenelHayvan
        SET BeslenmeDurumu = <cfqueryparam value="#BeslenmeDurumu#" cfsqltype="CF_SQL_VARCHAR">,
            RecordDate = <cfqueryparam value="#RecordDate#" cfsqltype="CF_SQL_DATE">
        WHERE HayvanTuru = <cfqueryparam value="#HayvanTuru#" cfsqltype="CF_SQL_VARCHAR"> AND CiftlikID = <cfqueryparam value="#CiftlikID#" cfsqltype="CF_SQL_INTEGER">
        </cfquery>
        <cflocation url="/GenelHayvan.cfm" addtoken="false"> 
    </cffunction>

    <cffunction name="addYeniHayvan" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="HayvanCesidi" type="string" required="true">
        <cfargument name="HayvanTuru" type="string" required="true">
        <cfargument name="HayvanNo" type="integer" required="true">
        <cfargument name="Cinsiyet" type="string" required="true">
        <cfargument name="Mahsul" type="string" required="true">
        <cfargument name="YavrulukDurumu" type="string" required="true">
        <cfargument name="RecordDate" type="date" required="true">
        <cfargument name="History" type="date" required="true">

        <cfquery name="beslenme" datasource="Ciftlik" >
            INSERT INTO HayvanDetay (CiftlikID,  HayvanCesidi, HayvanTuru, HayvanNo, Cinsiyet, Mahsul, YavrulukDurumu, RecordDate, History)
            VALUES
            (
                <cfqueryparam value="#CiftlikID#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#HayvanCesidi#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#HayvanTuru#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#HayvanNo#" cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#Cinsiyet#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#Mahsul#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#YavrulukDurumu#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#RecordDate#" cfsqltype="CF_SQL_DATE">,
                <cfqueryparam value="#History#" cfsqltype="CF_SQL_DATE">
            )
        </cfquery>

        <cfquery name="hayvancesidi1" datasource="Ciftlik" >
            SELECT HayvanCesidi
            FROM HayvanDetay
            WHERE HayvanNo = #HayvanNo#
        </cfquery>

        <cfoutput query="hayvancesidi1">
            <cfif #HayvanCesidi# eq 'Tavuk'>
                <cflocation url="/Hayvanlar/Tavuk.cfm" addtoken="false">
            <cfelseif #HayvanCesidi# eq 'Inek'>
                <cflocation url="/Hayvanlar/Inek.cfm" addtoken="false">
            <cfelseif #HayvanCesidi# eq 'Kopek'>
                <cflocation url="/Hayvanlar/Kopek.cfm" addtoken="false">
            </cfif>
        </cfoutput>
    </cffunction>

    <cffunction name="GebelikGuncelleme" access="remote" returntype="void">
        <cfargument name="GebelikDurumu" type="string" required="true">
        <cfargument name="HayvanNo" type="integer" required="true">
        <cfargument name="History" type="string" required="true">

        <cfquery name="gebe" datasource="Ciftlik" >
            UPDATE HayvanDetay
            SET GebelikDurumu = <cfqueryparam value="#GebelikDurumu#" cfsqltype="CF_SQL_VARCHAR">,
                History = <cfqueryparam value="#History#" cfsqltype="CF_SQL_DATE">
            WHERE HayvanNo = <cfqueryparam value="#HayvanNo#" cfsqltype="CF_SQL_INTEGER">
        </cfquery>

        <cfquery name="hayvancesidi1" datasource="Ciftlik" >
            SELECT HayvanCesidi
            FROM HayvanDetay
            WHERE HayvanNo = #HayvanNo#
        </cfquery>

        <cfoutput query="hayvancesidi1">
            <cfif #HayvanCesidi# eq 'Tavuk'>
                <cflocation url="/Hayvanlar/Tavuk.cfm" addtoken="false">
            <cfelseif #HayvanCesidi# eq 'Inek'>
                <cflocation url="/Hayvanlar/Inek.cfm" addtoken="false">
            <cfelseif #HayvanCesidi# eq 'Kopek'>
                <cflocation url="/Hayvanlar/Kopek.cfm" addtoken="false">
            </cfif>
        </cfoutput>
    </cffunction>

    <cffunction name="YavrulukGuncelleme" access="remote" returntype="void">
        <cfargument name="YavrulukDurumu" type="string" required="true">
        <cfargument name="HayvanNo" type="integer" required="true">
        <cfargument name="History" type="string" required="true">

        <cfquery name="yavru" datasource="Ciftlik" >
            UPDATE HayvanDetay
            SET YavrulukDurumu = <cfqueryparam value="#YavrulukDurumu#" cfsqltype="CF_SQL_VARCHAR">,
                History = <cfqueryparam value="#History#" cfsqltype="CF_SQL_DATE">
            WHERE HayvanNo = #HayvanNo#
        </cfquery>

        <cfquery name="hayvancesidi1" datasource="Ciftlik" >
            SELECT HayvanCesidi
            FROM HayvanDetay
            WHERE HayvanNo = #HayvanNo#
        </cfquery>

        <cfoutput query="hayvancesidi1">
            <cfif #HayvanCesidi# eq 'Tavuk'>
                <cflocation url="/Hayvanlar/Tavuk.cfm" addtoken="false">
            <cfelseif #HayvanCesidi# eq 'Inek'>
                <cflocation url="/Hayvanlar/Inek.cfm" addtoken="false">
            <cfelseif #HayvanCesidi# eq 'Kopek'>
                <cflocation url="/Hayvanlar/Kopek.cfm" addtoken="false">
            </cfif>
        </cfoutput>
    </cffunction>
</cfcomponent>