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
            stQuerytxb = "Select * from ""AStockWEB"""
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
