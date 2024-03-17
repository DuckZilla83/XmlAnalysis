'Script version 0.1 MVP sirve para el trabajo encomendado, se requiere automatización del proceso.

' Obtener la ruta del script actual
Dim currentScriptPath
Set objShell = CreateObject("WScript.Shell")
currentScriptPath = objShell.CurrentDirectory

' Ruta de la carpeta CFDI
Dim folderPath
folderPath = currentScriptPath & "\CFDI"

' Crear objeto FileSystem
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Crear archivo de salida
Set objOutputFile = objFSO.CreateTextFile(currentScriptPath & "\salida.txt", True)

' Declarar espacio de nombres
Const cfdiNamespace = "http://www.sat.gob.mx/cfd/4"

' Escribir encabezado en el archivo de salida
objOutputFile.WriteLine "Fecha|Folio|TipoRelacion|UUID|Descripcion|Importe|Impuesto|TipoFactor|Base|TasaOCuota|ImporteImpuesto|"

' Recorrer archivos XML en la carpeta CFDI
For Each objFile In objFSO.GetFolder(folderPath).Files
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "xml" Then
        ' Cargar el archivo XML
        Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")
        xmlDoc.async = False
        xmlDoc.setProperty "SelectionLanguage", "XPath"
        xmlDoc.setProperty "SelectionNamespaces", "xmlns:cfdi='" & cfdiNamespace & "'"
        xmlDoc.load objFile.Path

        ' Obtener el atributo Folio
        Dim folio, tipoRelacion, fechaDoc
        fechaDoc = xmlDoc.SelectSingleNode("//cfdi:Comprobante").getAttribute("Fecha")
        ' Convertir la fecha al formato dd-mm-yyyy
        fechaDoc = Mid(fechaDoc, 9, 2) & "-" & Mid(fechaDoc, 6, 2) & "-" & Mid(fechaDoc, 1, 4)
        folio = xmlDoc.SelectSingleNode("//cfdi:Comprobante").getAttribute("Folio")
        tipoRelacion = xmlDoc.SelectSingleNode("//cfdi:CfdiRelacionados").getAttribute("TipoRelacion")

        ' Obtener datos de los nodos cfdi:CfdiRelacionados y cfdi:Concepto
        Dim cfdiRelacionados
        Set cfdiRelacionados = xmlDoc.SelectNodes("//cfdi:CfdiRelacionados/cfdi:CfdiRelacionado")
        Dim conceptos
        Set conceptos = xmlDoc.SelectNodes("//cfdi:Conceptos/cfdi:Concepto")

        ' Escribir datos en el archivo de salida
        For Each cfdiRelacionado In cfdiRelacionados
            For Each concepto In conceptos
                objOutputFile.Write fechaDoc & "|" & folio & "|" & tipoRelacion & "|" & cfdiRelacionado.getAttribute("UUID") & "|"
                objOutputFile.Write concepto.getAttribute("Descripcion") & "|" & concepto.getAttribute("Importe") & "|"
                Dim traslado
                Set traslado = concepto.SelectSingleNode("cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado")
                objOutputFile.Write traslado.getAttribute("Impuesto") & "|" & traslado.getAttribute("TipoFactor") & "|"
                objOutputFile.Write traslado.getAttribute("Base") & "|" & traslado.getAttribute("TasaOCuota") & "|"
                objOutputFile.WriteLine traslado.getAttribute("Importe")
            Next
        Next
    End If
Next

' Cerrar archivo de salida
objOutputFile.Close
