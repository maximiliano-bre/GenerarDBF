Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb
Imports System.IO


Public Class clsGenerarDBF

'origen de datos
    Public strcn As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=C:\CARPETA\Liquidacion_XXXX;" & _
              "Extended Properties='dBASE IV;'"

    Dim cnn As New OleDbConnection(strcn)

'destino del archivo DBF
    Const _FOLDERIMAGE As String = "C:\CARPETA\Liquidacion_XXXX\"

   
    Public Sub GenerarDBF(ByVal NroLiquidacion As String, ByVal dtArchivo As DataTable)


        If VerificarLiquidacion(NroLiquidacion) = False Then

            'Debo Crear la Liquidacion

            Dim n As Integer = CreateDbf(dtArchivo, "LIQ" & NroLiquidacion, cnn)

        End If

    End Sub



'VERIFICAR SI EXISTE 
    Public Function VerificarLiquidacion(ByVal NroLiquidacion As String) As Boolean

        If System.IO.File.Exists(_FOLDERIMAGE & "LIQ" & NroLiquidacion & ".DBF") Then

            Throw New ArgumentNullException("Liquidacion", "la liquidacion que intenta Crear ya existe")
            Return True

        Else
            Return False
        End If

    End Function





    Public Function VerificarArchivo(ByVal NroLiquidacion As String) As Boolean
        If System.IO.File.Exists(_FOLDERIMAGE & "LIQ" & NroLiquidacion & ".DBF") Then
            Return True
        Else
            Return False
        End If
    End Function
	
	
    Private Function CreateDbf(ByVal dt As DataTable, _
                          ByVal tableName As String, _
                          ByVal cnn As OleDbConnection) As Integer
      
        If (dt Is Nothing) Then _
            Throw New ArgumentNullException("dt", _
                "El objeto no es válido")

        If (String.IsNullOrEmpty(tableName)) Then _
            Throw New ArgumentNullException("tableName", _
                "No se ha especificado el nombre de la tabla.")

        If (cnn Is Nothing) Then _
            Throw New ArgumentNullException("cnn", _
                "El objeto Connection no es válido.")

        Dim sql As New System.Text.StringBuilder(256)

        Dim sql1 As New System.Text.StringBuilder(256)

        For Each row As DataRow In dt.Rows
            row.SetAdded()
        Next

        Try

            For Each dc As DataColumn In dt.Columns

                sql.Append(GetDataTypeSql(dc))
            Next

            Dim STRSQL = "CREATE TABLE " & tableName & "(LIQUIDACION VARCHAR(10),DNI Varchar(10),CE VARCHAR(10),MONTO Varchar(10))"

            Using cnn
              
                Dim cmd As New OleDbCommand(STRSQL, cnn)
                cnn.Open()
                cmd.ExecuteNonQuery()
                cmd.CommandText = String.Format("SELECT * FROM [{0}]", tableName)

                Dim da As New OleDbDataAdapter(cmd)

                Dim cb As New OleDbCommandBuilder(da)

                cb.QuotePrefix = "["
                cb.QuoteSuffix = "]"

                da.InsertCommand = cb.GetInsertCommand()

                Return da.Update(dt)

            End Using
            sql = Nothing
        Catch ex As Exception

            Throw New ArgumentNullException("Base ", _
               " - Fallo en la Creacion de la Base" & ex.Message)
        End Try

    End Function

    Private Function GetDataTypeSql(ByVal dc As DataColumn) As String
        Dim columnName As String = dc.ColumnName
        Dim dataType As String
        Dim maxLength As Int32
        Select Case dc.DataType.Name
            Case "Boolean"
                dataType = "bit"

            Case "Byte", "SByte"
                dataType = "tinyint"

            Case "Char"
                dataType = "nchar"
                maxLength = dc.MaxLength

            Case "DateTime"
                dataType = "datetime"

            Case "Decimal"
                dataType = "decimal (10, 2)"

            Case "Double"
                dataType = "decimal (8, 2)"
                

            Case "Int16", "UInt16"
                dataType = "NUMERIC (5,0)"

            Case "Int32", "UInt32"
                ' dataType = "int"
                dataType = "nvarchar"
                maxLength = "10"
 
            Case "Int64", "UInt64"
                dataType = "bigint"

            Case "Object", "Byte[]"
                dataType = "image"

            Case "Single"
                dataType = "float"

            Case "string"
                dataType = "nchar"
                maxLength = "15"
            Case Else   ' String
                If (dc.MaxLength = 536870910) Then
                    dataType = "memo"
                Else
                    dataType = "nvarchar"
                    maxLength = "15"
                End If
        End Select
        If (maxLength > 0) Then
            Return String.Format("[{0}] {1} ({2}),", columnName, dataType, maxLength)
        Else
            Return String.Format("[{0}] {1},", columnName, dataType)
        End If
    End Function

    Public Function CantidadRegistros(ByVal tableName As String, _
                     ByVal strcn As String) As Integer

        Dim sSelect As String = "SELECT * FROM " & tableName
        Dim ds As New DataSet
        Dim Cantidad As Integer = 0
        Using dbConn As New System.Data.OleDb.OleDbConnection(strcn)
            Try
                dbConn.Open()
                Dim da As New System.Data.OleDb.OleDbDataAdapter(sSelect, dbConn)
                Dim dt As New DataTable
                da.Fill(dt)

                Cantidad = dt.Rows.Count()

                dbConn.Close()

                Return Cantidad

            Catch ex As Exception

                Throw New ArgumentNullException("Base", "Fallo en la Cantidad de Registros" + ex.Message)
                Exit Function

                Return Cantidad

            End Try
        End Using

    End Function

    Public Function TotalImporte(ByVal tableName As String, _
                       ByVal strcn As String) As Decimal

        Dim sSelect As String = "SELECT * FROM " & tableName
        Dim ds As New DataSet
        Dim Cantidad As Integer = 0
        Dim icont As Integer
        Dim MontoT As Decimal

        Dim valor As String
        Using dbConn As New System.Data.OleDb.OleDbConnection(strcn)
            Try
                dbConn.Open()
                Dim da As New System.Data.OleDb.OleDbDataAdapter(sSelect, dbConn)
                Dim dt As New DataTable
                da.Fill(dt)

                Cantidad = dt.Rows.Count()
                
                For icont = 0 To Cantidad - 1
                    valor = Decimal.Parse(dt.Rows(icont)("monto").Replace(".", ","))
                    MontoT = MontoT + CType(valor, Decimal)
                   
                Next
                dbConn.Close()
                Return MontoT
            Catch ex As Exception
                Throw New ArgumentNullException("TotalImporte", "Fallo en Calcular el Total del Importe" + ex.Message)
                Exit Function

                Return MontoT

            End Try
        End Using

    End Function

    Public Function BorrarLiquidacion(ByVal NroLiquidacion As String) As String
   
        Try
            File.Delete(_FOLDERIMAGE + "LIQ" & NroLiquidacion & ".DBF")
            Return "El Archivo fue Borrado, debera genera otro archivo !!!"
        Catch ex As Exception
            Return "No se pudo eliminar el archivo," & ex.Message
        End Try

    End Function


    Public Function Levantar_DBF(ByVal tableName As String) As DataTable

        Dim sSelect As String = "SELECT * FROM " & tableName & ".DBF"
        Dim ds As New DataSet
        Dim Cantidad As Integer = 0

        Using dbConn As New System.Data.OleDb.OleDbConnection(strcn)
            Try
                dbConn.Open()
                Dim da As New System.Data.OleDb.OleDbDataAdapter(sSelect, dbConn)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Catch ex As Exception
                Throw New ArgumentNullException("TotalImporte", "Fallo en Calcular el Total del Importe" + ex.Message)
                Exit Function
 
            End Try
        End Using

    End Function

End Class
