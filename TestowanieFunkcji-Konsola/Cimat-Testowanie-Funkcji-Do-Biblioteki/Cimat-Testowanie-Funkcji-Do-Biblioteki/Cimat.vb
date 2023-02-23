Imports Inventor
Imports System.Runtime.InteropServices
Imports Autodesk.iLogic.Interfaces

Module Cimat
    Dim _invapp As Inventor.Application ' _invapp działą tak samo jak ThisApplication
    Sub Main()
        IsBended()
    End Sub


    Sub IsBended() 'Funkcja sprawdzająca czy blacha jest gięta
        getThisApplication()
        Dim oDoc As Document
        oDoc = _invapp.ActiveDocument

        If oDoc.DocumentSubType.DocumentSubTypeID = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
            Dim oPartDoc As PartDocument
            oPartDoc = _invapp.ActiveDocument
            Dim oPartDef As SheetMetalComponentDefinition
            oPartDef = oPartDoc.ComponentDefinition

            If oPartDef.Bends.Count >= 1 Then ' Jeżeli ma więcej niż jedno gięcie
                UpdateCustomProperty("Blacha gięta", "X") ' to ustaw X jak właściwość "Blacha gięta"
            Else
                UpdateCustomProperty("Blacha gięta", "") 'ustaw pustą właściwość "Blacha gięta"

            End If
        Else
        End If
    End Sub
    Sub SetSubpartsAsPhantomForPurcharesAssembly()
        getThisApplication()
        Dim oDoc As Document
        oDoc = _invapp.ActiveDocument

        If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim oAsmDoc As AssemblyDocument
            oAsmDoc = _invapp.ActiveDocument
            Dim oAsmDef As AssemblyComponentDefinition
            oAsmDef = oAsmDoc.ComponentDefinition

            If oAsmDef.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then

                Dim oBOM As BOM
                oBOM = oAsmDef.BOM
                oBOM.StructuredViewEnabled = True
                Dim oBOMView As BOMView
                oBOMView = oBOM.BOMViews.Item("Bez nazwy") 'Zakładka "Dane modelu"

                Try
                    For i = 1 To oBOMView.BOMRows.Count
                        If oBOMView.BOMRows.Item(i).BOMStructure <> BOMStructureEnum.kPhantomBOMStructure Then oBOMView.BOMRows.Item(i).BOMStructure = BOMStructureEnum.kPhantomBOMStructure
                    Next
                Catch
                    MsgBox("Nie można zmienić elementów złożenia na pozorne." & vbCrLf & vbCrLf & "Prawdopodobnie pliki będące elementem złożenia nie są wypisane z Vaulta." & vbCrLf & vbCrLf & "Wypisz pliki i ponownie zapisz złożenie.", MsgBoxStyle.Information, "Błąd")
                End Try

            End If

        End If

    End Sub
    Sub getThisApplication()

        Try ' Spróbuj pobrać aktuwną instancję Inventora
            Try
                _invapp = Marshal.GetActiveObject("Inventor.Application")
            Catch ' Jeśli ni ejest aktywna, stwórz ją
                ' Sesja Iventora
                Dim inventorAppType As Type = System.Type.GetTypeFromProgID("Inventor.Application")
                _invapp = System.Activator.CreateInstance(inventorAppType)
                'Musi być ustawiony widoczny
                _invapp.Visible = True
            End Try
        Catch
            MsgBox("Error: nie można utworzyć instancji Inventora")
        End Try
    End Sub
    Sub UpdateCustomProperty(nameProperty As String, valueParameter As String)
        Dim oDoc As Document
        oDoc = _invapp.ActiveDocument
        Dim customPropSet As PropertySet
        customPropSet = oDoc.PropertySets.Item("Właściwości użytkownika programu Inventor")
        Dim customProperty As [Property]
        Try
            customProperty = customPropSet.Item(nameProperty)
            customProperty.Value = valueParameter
        Catch ex As Exception
            customPropSet.Add("", nameProperty)
            customProperty = customPropSet.Item(nameProperty)
            customProperty.Value = valueParameter
        End Try
    End Sub

End Module