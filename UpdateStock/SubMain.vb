Imports Sap.Data.Hana
Imports SAPbobsCOM

Module SubMain

    Public SBOCompany As SAPbobsCOM.Company

    Sub Main()

        Conectar()
        Delete()
        Update()

    End Sub

    Public Function Conectar()

        Try

            SBOCompany = New SAPbobsCOM.Company

            SBOCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            SBOCompany.Server = My.Settings.Server
            SBOCompany.LicenseServer = My.Settings.LicenseServer
            SBOCompany.DbUserName = My.Settings.DbUserName
            SBOCompany.DbPassword = My.Settings.DbPassword

            SBOCompany.CompanyDB = My.Settings.CompanyDB

            SBOCompany.UserName = My.Settings.UserName
            SBOCompany.Password = My.Settings.Password

            SBOCompany.Connect()

        Catch ex As Exception

            MsgBox("Error al Conectar: " & ex.Message)

        End Try

    End Function


    Public Function Delete()

        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Delete from ""@STOCKWEB"""
            oRecSettxb.DoQuery(stQuerytxb)

        Catch ex As Exception

            MsgBox("Error al Borrar: " & ex.Message)

        End Try

    End Function


    Public Function Update()

        Dim oRecSettxb, oRecSettxb2 As SAPbobsCOM.Recordset
        Dim stQuerytxb, stQuerytxb2 As String
        Dim Number, Sku, Unit, WhsCode, Stock As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "select 

                    case when Length(Row_Number() Over())=5 then concat(0,Row_Number() Over())
                    when Length(Row_Number() Over())=4 then concat(00,Row_Number() Over())
                    when Length(Row_Number() Over())=3 then concat(000,Row_Number() Over())
                    when Length(Row_Number() Over())=2 then concat(0000,Row_Number() Over())
                    when Length(Row_Number() Over())=1 then concat(00000,Row_Number() Over())
                    end as ""Number"",

                    T0.""sku"", T0.""purchaseunit"", T0.""warehousecode""
                    ,sum(T0.""stock"")AS ""stock""

                    from (
                    SELECT T0.""ItemCode"" as""sku""

                    , CASE WHEN T0.""WhsCode"" IN('001','001A','001B','001C') THEN '001' ELSE T0.""WhsCode"" END as ""warehousecode""
                    , CASE 
	                WHEN T1.""SalUnitMsr"" = 'MTK' then 'm2'
	                WHEN T1.""SalUnitMsr"" = 'LM' then 'm'
	                WHEN T1.""SalUnitMsr"" = 'LTR' then 'litro'
	                WHEN T1.""SalUnitMsr"" = 'AS' then 'pza'
	                WHEN T1.""SalUnitMsr"" = 'H87' then 'pza'
	                ELSE T1.""SalUnitMsr"" 
	                END as""purchaseunit"" 
                    , T0.""OnHand""as""StockSinFormula""

                    , CASE 
	                WHEN T0.""Locked"" = 'Y' THEN '0'
	                ELSE T0.""OnHand"" 
	                END AS ""stock""
	
                    FROM OITW T0 
                    LEFT OUTER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode""
                    --LEFT OUTER JOIN OWHS T2 ON T0.""WhsCode"" = T2.""WhsCode""

                    WHERE T1.""validFor"" = 'Y'
                    AND T1.""ItemCode"" NOT LIKE 'MUES%'
                    AND T1.""ItemCode"" NOT LIKE 'Y%'
                    AND T1.""ItmsGrpCod"" IN ('101','102','103','104','135','139','107','108','136','111','110','112','113','115','114'
                    ,'116','131','126','133','117','118','119','141','122','123','124','134','138','144','146')
                    AND T0.""WhsCode"" NOT IN ('007','015','016','019','023','024','025','700','701'
                    ,'702','703','704','995','996','997','998','999') 
                    ) T0
                    where T0.""stock"" <> '0'
                    group by T0.""sku"", T0.""warehousecode"", T0.""purchaseunit""
                    ORDER BY 1,4, 2"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                For cont As Integer = 0 To oRecSettxb.RecordCount - 1

                    Number = oRecSettxb.Fields.Item("Number").Value
                    Sku = oRecSettxb.Fields.Item("sku").Value
                    Unit = oRecSettxb.Fields.Item("purchaseunit").Value
                    WhsCode = oRecSettxb.Fields.Item("warehousecode").Value
                    Stock = oRecSettxb.Fields.Item("stock").Value

                    oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    stQuerytxb2 = "INSERT INTO ""@STOCKWEB"" VALUES ('" & Number & "','" & Number & "','" & Sku & "','" & WhsCode & "','" & Unit & "'," & Stock & ")"
                    oRecSettxb2.DoQuery(stQuerytxb2)

                    oRecSettxb.MoveNext()

                Next

            End If

        Catch ex As Exception

            MsgBox("Error al Actualizar: " & ex.Message)

        End Try

    End Function


End Module
