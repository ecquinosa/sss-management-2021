
Imports System
Imports System.Xml
Imports System.Configuration
Imports System.Reflection
Imports System.IO
Imports System.Xml.XPath


Public Class MySettings
    Public xml_Filename, value_path, value_name As String
    Public Shared settings As New XmlWriterSettings()
    ' value_name  =  value name of the settings.
    ' value_path =   node path of the value.
    'xml_filename  =  the filename of the settings.

    Public Sub writeSettings(ByVal xml_Filename As String)
        ' Dim doc As XmlDocument = loadconfig
        Try
            If Not File.Exists(xml_Filename) Then
                Dim XmlWrt As XmlWriter = XmlWriter.Create(xml_Filename, settings)


                With XmlWrt
                    .WriteStartDocument()

                    .WriteComment("XML Settings.")
                    .WriteStartElement("Configuration")

                    .WriteStartElement("Settings")


                    .WriteStartElement("VPN_USERNAME")
                    .WriteString("")
                    .WriteEndElement()

                    .WriteStartElement("VPN_PASSWORD")
                    .WriteString("")
                    .WriteEndElement()


                    .WriteEndElement()
                    .WriteEndElement()
                    .WriteEndDocument()
                    .Flush()
                    .Close()
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function editSettings(ByVal xml_Filename As String, ByVal xml_path As String, ByVal value_name As String, ByVal value1 As String)
        Try
            'GC.Collect()
            Dim xd As New XmlDocument()

            xd.Load(xml_Filename)

            Dim nod As XmlNode = xd.SelectSingleNode(xml_path & "/" & value_name)  '"Configuration/Settings/FirstName"
            If nod IsNot Nothing Then
                nod.InnerXml = value1
            Else
                nod.InnerXml = value1
            End If

            xd.Save(xml_Filename)



        Catch ex As Exception


        End Try


    End Function

    Public Function readSettings(ByVal xml_filame As String, ByVal xml_path As String, ByVal value_name As String) As String
        Dim return_value As String
        Try
            Dim a As String
            Dim xd As New XmlDocument
            xd.Load(xml_filame)
            Dim nod As XmlNode = xd.SelectSingleNode(xml_path & "/" & value_name)
            If nod IsNot Nothing Then
                a = nod.InnerXml
                return_value = a
            Else
                return_value = Nothing
            End If
            xd.Save(xml_filame)
        Catch ex As Exception
            return_value = ex.Message
        End Try
        Return return_value

    End Function
End Class
