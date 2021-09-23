Public Class UpdateOPOR


    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private Directorio As String


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Directorio = oCatchingEvents.csDirectory
    End Sub


    Public Function Update(ByVal DocNum As String, ByVal Cost1 As Double, ByVal Cost2 As Double, ByVal Cost3 As Double, ByVal Cost4 As Double, ByVal Cost5 As Double, ByVal Resubir As String)

        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim OPOR As SAPbobsCOM.Documents
        Dim POR1 As SAPbobsCOM.Document_Lines
        Dim DocEntry, Lineas As Integer
        'Dim LinePrice1, LinePrice2, CurrentPrice, Price1, Price2 As Decimal
        Dim CurrentPrice, Price1, Price2, Price3, Price4, Price5 As Decimal
        Dim Line, l As Integer
        Dim scl As Integer
        Dim SumQty, percent, LinePrice1, LinePrice2, LinePrice3, LinePrice4, LinePrice5, LstPrice1, LstPrice2, LstPrice3, LstPrice4, LstPrice5 As Double

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        OPOR = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

        Try

            stQueryH = "Select T0.""DocEntry"",T0.""DocTotal"", count(T1.""ItemCode"") as ""Lines"", sum(T1.""Quantity"") as ""SumQty"" From OPOR T0 Inner Join POR1 T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocNum""=" & DocNum & " Group by T0.""DocEntry"",T0.""DocTotal"""
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                Lineas = oRecSetH.Fields.Item("Lines").Value
                SumQty = oRecSetH.Fields.Item("SumQty").Value

                'LinePrice1 = Cost1 / Lineas
                'LinePrice1 = Decimal.Round(LinePrice1, 4)
                'LinePrice2 = Cost2 / Lineas
                'LinePrice2 = Decimal.Round(LinePrice2, 4)

                OPOR.GetByKey(DocEntry)
                POR1 = OPOR.Lines

                stQueryH3 = "Select Top 1 T0.""LineNum"" from POR1 T0 where T0.""DocEntry""=" & DocEntry & " order by ""LineNum"" Asc"
                oRecSetH3.DoQuery(stQueryH3)

                If oRecSetH3.RecordCount > 0 Then

                    Line = oRecSetH3.Fields.Item("LineNum").Value
                    l = Line + Lineas - 1
                    scl = 0

                    For i = Line To l

                        stQueryH2 = "Select T0.""Quantity"",T0.""Price"",T0.""U_OriginalPrice"" from POR1 T0 where T0.""DocEntry""=" & DocEntry & " And T0.""LineNum""=" & i '& " Order by T0.""NumLine"" Desc"
                        oRecSetH2.DoQuery(stQueryH2)

                        If oRecSetH2.RecordCount > 0 Then

                            percent = (oRecSetH2.Fields.Item("Quantity").Value * 100) / SumQty

                            If Cost1 = 0 Then
                                LstPrice1 = 0
                            Else
                                LinePrice1 = (percent * Cost1) / 100
                                Price1 = LinePrice1 / oRecSetH2.Fields.Item("Quantity").Value
                                LstPrice1 = Decimal.Round(Price1, 4)
                            End If

                            If Cost2 = 0 Then
                                LstPrice2 = 0
                            Else
                                LinePrice2 = (percent * Cost2) / 100
                                Price2 = LinePrice2 / oRecSetH2.Fields.Item("Quantity").Value
                                LstPrice2 = Decimal.Round(Price2, 4)
                            End If

                            If Cost3 = 0 Then
                                LstPrice3 = 0
                            Else
                                LinePrice3 = (percent * Cost3) / 100
                                Price3 = LinePrice3 / oRecSetH2.Fields.Item("Quantity").Value
                                LstPrice3 = Decimal.Round(Price3, 4)
                            End If

                            If Cost4 = 0 Then
                                LstPrice4 = 0
                            Else
                                LinePrice4 = (percent * Cost4) / 100
                                Price4 = LinePrice4 / oRecSetH2.Fields.Item("Quantity").Value
                                LstPrice4 = Decimal.Round(Price4, 4)
                            End If

                            If Cost5 = 0 Then
                                LstPrice5 = 0
                            Else
                                LinePrice5 = (percent * Cost5) / 100
                                Price5 = LinePrice5 / oRecSetH2.Fields.Item("Quantity").Value
                                LstPrice5 = Decimal.Round(Price5, 4)
                            End If

                            'Price1 = LinePrice1 / oRecSetH2.Fields.Item("Quantity").Value
                            'LstPrice1 = Decimal.Round(Price1, 4)
                            'Price2 = LinePrice2 / oRecSetH2.Fields.Item("Quantity").Value
                            'LstPrice2 = Decimal.Round(Price2, 4)

                            POR1.SetCurrentLine(scl)

                            If Resubir = 0 Then
                                OPOR.Lines.UserFields.Fields.Item("U_OriginalPrice").Value = oRecSetH2.Fields.Item("Price").Value
                            End If

                            OPOR.Lines.UserFields.Fields.Item("U_ExtraPrice1").Value = LstPrice1
                            OPOR.Lines.UserFields.Fields.Item("U_ExtraPrice2").Value = LstPrice2
                            OPOR.Lines.UserFields.Fields.Item("U_ExtraPrice3").Value = LstPrice3
                            OPOR.Lines.UserFields.Fields.Item("U_ExtraPrice4").Value = LstPrice4
                            OPOR.Lines.UserFields.Fields.Item("U_ExtraPrice5").Value = LstPrice5

                            If Resubir = 0 Then
                                CurrentPrice = oRecSetH2.Fields.Item("Price").Value + LstPrice1 + LstPrice2 + LstPrice3 + LstPrice4 + LstPrice5
                            Else
                                CurrentPrice = oRecSetH2.Fields.Item("U_OriginalPrice").Value + LstPrice1 + LstPrice2 + LstPrice3 + LstPrice4 + LstPrice5
                            End If

                            OPOR.Lines.UnitPrice = CurrentPrice

                            OPOR.Lines.Add()

                            scl = scl + 1

                        End If

                    Next

                End If

                OPOR.Update()

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Update. Error al actualizar el documento, " & ex.Message)

        End Try

    End Function


End Class
