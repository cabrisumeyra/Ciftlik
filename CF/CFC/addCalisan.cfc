<cfcomponent>
    <cffunction name="CalisanEkle" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="CalisanIsim" type="string" required="true">
        <cfargument name="CalisanSoyad" type="string" required="true">
        <cfargument name="CalisanYas" type="integer" required="true"> 
        <cfargument name="Maas" type="integer" required="true">
        <cfargument name="GirisTarihi" type="date" required="true">

        <cfquery name="CalisanEkleme" datasource="Ciftlik" >
            INSERT INTO Calisan (CiftlikID, CalisanIsim, CalisanSoyad, CalisanYas, Maas, GirisTarihi)
            VALUES (
                <cfqueryparam value="#CiftlikID#"    cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#CalisanIsim#"  cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#CalisanSoyad#" cfsqltype="CF_SQL_VARCHAR">,
                <cfqueryparam value="#CalisanYas#"   cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#Maas#"         cfsqltype="CF_SQL_INTEGER">,
                <cfqueryparam value="#GirisTarihi#"  cfsqltype="CF_SQL_DATE">
            )
        </cfquery>

        <cflocation url="/Calisan.cfm" addtoken="false">
    </cffunction>

    <cffunction name="CikisEkle" access="remote" returntype="void">
        <cfargument name="CiftlikID" type="integer" required="true">
        <cfargument name="CalisanIsim" type="string" required="true">
        <cfargument name="CalisanSoyad" type="string" required="true">
        <cfargument name="CikisTarihi" type="date" required="true">

        <cfquery name="CikisEkleme" datasource="Ciftlik" >
            UPDATE Calisan 
            SET 
                CikisTarihi = <cfqueryparam value="#CikisTarihi#"  cfsqltype="CF_SQL_DATE">
                WHERE CiftlikID = <cfqueryparam value="#CiftlikID#"    cfsqltype="CF_SQL_INTEGER"> AND CalisanIsim = <cfqueryparam value="#CalisanIsim#"  cfsqltype="CF_SQL_VARCHAR"> AND CalisanSoyad = <cfqueryparam value="#CalisanSoyad#" cfsqltype="CF_SQL_VARCHAR">             
        </cfquery>

        <cflocation url="/Calisan.cfm" addtoken="false">
    </cffunction>

</cfcomponent>
