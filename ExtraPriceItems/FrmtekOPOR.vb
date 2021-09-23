﻿Imports System.Drawing

Public Class FrmtekOPOR

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openForm(ByVal psDirectory As String)

        Try

            csFormUID = "TekExtraCost"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If

            '--- Referencia de Forma
            setForm(csFormUID)

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmTratamientoPedidos. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources() As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources()
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources() As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid

        Try
            bindUserDataSources = 0

            loDS = coForm.DataSources.UserDataSources.Add("dsCost1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("4").Specific  'identifico mi combobox
            loText.DataBind.SetBound(True, "", "dsCost1")   ' uno mi userdatasources a mi combobox

            loDS = coForm.DataSources.UserDataSources.Add("dsCost2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("5").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsCost2")   ' uno mi userdatasources a mi caja de texto

            loDS = coForm.DataSources.UserDataSources.Add("dsCost3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("9").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsCost3")   ' uno mi userdatasources a mi caja de texto

            loDS = coForm.DataSources.UserDataSources.Add("dsCost4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("10").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsCost4")   ' uno mi userdatasources a mi caja de texto

            loDS = coForm.DataSources.UserDataSources.Add("dsCost5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("11").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsCost5")   ' uno mi userdatasources a mi caja de texto

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function

End Class
