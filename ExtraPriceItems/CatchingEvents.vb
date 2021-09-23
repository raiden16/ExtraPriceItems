Public Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF
    Dim DocNum, Resubir As String

    Public Sub New()

        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        setFilters()

    End Sub

    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        End Try
    End Sub

    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
            'Finally
        End Try
    End Sub

    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        Finally
            loRecSet = Nothing
        End Try
    End Sub

    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try
            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            lofilter.AddEx(142) '// FORMA Pedido Proveedores
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx("TekExtraCost")
            lofilter.AddEx(142) '// FORMA Pedido Proveedores

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub

    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then
        Else
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then
                Select Case pVal.FormTypeEx

                    Case 142                           '////// FORMA Factura
                        frmORDRControllerAfter(FormUID, pVal)

                    Case "TekExtraCost"
                        frmOPORControllerAfter(FormUID, pVal)

                End Select
            End If
        End If

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA Estados de cuenta externos
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmORDRControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim obotton As OPOR
        Dim oOPOR As FrmtekOPOR
        Dim stTabla As String
        Dim coForm As SAPbouiCOM.Form
        Dim oDatatable As SAPbouiCOM.DBDataSource
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim OriginalPrice As Double
        Dim Respuesta As Integer

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType

                '///// Carga de formas
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    obotton = New OPOR
                    obotton.addFormItems(FormUID)

                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case "btCost"

                            stTabla = "OPOR"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            DocNum = oDatatable.GetValue("DocNum", 0)

                            stQueryH = "Select Top 1 T1.""U_OriginalPrice"" From OPOR T0 Inner Join POR1 T1 on T1.""DocEntry""=T0.""DocEntry"" where ""DocNum""=" & DocNum & " Order by T1.""LineNum"" Asc"
                            oRecSetH.DoQuery(stQueryH)

                            If oRecSetH.RecordCount > 0 Then

                                OriginalPrice = oRecSetH.Fields.Item("U_OriginalPrice").Value

                                If OriginalPrice <> Nothing Or OriginalPrice <> 0 Then

                                    Respuesta = SBOApplication.MessageBox("Ya se agregaron costos extra, ¿Quieres agregarlos de nuevo?", Btn1Caption:="Si", Btn2Caption:="No")

                                    If Respuesta = 1 Then

                                        Resubir = 1
                                        oOPOR = New FrmtekOPOR
                                        oOPOR.openForm(csDirectory)

                                    End If

                                Else

                                    Resubir = 0
                                    oOPOR = New FrmtekOPOR
                                    oOPOR.openForm(csDirectory)

                                End If

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Facturacion Clientes. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub

    Private Sub frmOPORControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim coForm As SAPbouiCOM.Form
        Dim Cost11, Cost22, Cost33, Cost44, Cost55 As String
        Dim Cost1, Cost2, Cost3, Cost4, Cost5 As Double
        Dim oUpOPOR As UpdateOPOR
        Dim Respuesta As Integer

        Try

            Select Case pVal.EventType

                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case "1"

                            Respuesta = SBOApplication.MessageBox("¿Deseas agregar los costos extra?", Btn1Caption:="Si", Btn2Caption:="No")

                            If Respuesta = 1 Then

                                coForm = SBOApplication.Forms.Item(FormUID)
                                Cost11 = coForm.DataSources.UserDataSources.Item("dsCost1").Value
                                If Cost11 = Nothing Then
                                    Cost1 = 0
                                Else
                                    Cost1 = Cost11
                                End If

                                Cost22 = coForm.DataSources.UserDataSources.Item("dsCost2").Value
                                If Cost22 = Nothing Then
                                    Cost2 = 0
                                Else
                                    Cost2 = Cost22
                                End If

                                Cost33 = coForm.DataSources.UserDataSources.Item("dsCost3").Value
                                If Cost33 = Nothing Then
                                    Cost3 = 0
                                Else
                                    Cost3 = Cost33
                                End If

                                Cost44 = coForm.DataSources.UserDataSources.Item("dsCost4").Value
                                If Cost44 = Nothing Then
                                    Cost4 = 0
                                Else
                                    Cost4 = Cost44
                                End If

                                Cost55 = coForm.DataSources.UserDataSources.Item("dsCost4").Value
                                If Cost55 = Nothing Then
                                    Cost5 = 0
                                Else
                                    Cost5 = Cost55
                                End If

                                oUpOPOR = New UpdateOPOR
                                oUpOPOR.Update(DocNum, Cost1, Cost2, Cost3, Cost4, Cost5, Resubir)

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Facturacion Clientes. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub

End Class
