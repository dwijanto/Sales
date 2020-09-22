Imports System.IO
Imports System.Text
Imports Sales.PublicClass
Imports Sales.SharedClass
Imports Microsoft.Office.Interop
Public Class FormConvertEncoding
    Dim openFileDialog1 As New OpenFileDialog
    Dim sb As New stringbuilder
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'get filename
        If openFileDialog1.ShowDialog = DialogResult.OK Then
            Try
                Dim i As Integer = 0
                Dim sr As StreamReader = New StreamReader(openFileDialog1.FileName, System.Text.Encoding.Unicode)
                Dim line As String
                Dim sw As New StreamWriter("c:\tmp\utf8.txt", False, System.Text.Encoding.UTF8)
                Do

                    line = sr.ReadLine
                    If i > 0 Then
                        Try
                            If Not IsNothing(line) Then
                                If line.Length > 0 Then
                                    Debug.WriteLine(line)
                                    sw.WriteLine(line)
                                End If
                            End If
                            
                        Catch ex As Exception
                            Debug.WriteLine(ex.Message)
                        End Try
                        

                    End If
                    i += 1

                Loop Until line Is Nothing
                sr.Close()
                sw.Close()


                Dim sr2 As StreamReader = New StreamReader("c:\tmp\utf8.txt", System.Text.Encoding.UTF8)
                i = 0
                Do
                    i += 1
                    If i = 611 Then
                        Debug.Print("hello")
                    End If
                    line = sr2.ReadLine
                    If Not IsNothing(line) Then
                        If line.Length > 0 Then
                            sb.Append(line & vbCrLf)
                        End If
                    End If
                    
                Loop Until line Is Nothing
                sr2.Close()
                If sb.Length > 0 Then
                    Dim myret As Boolean = False
                    ' cmmf,productid,sbu,prodfamily,cdesc, brand
                    Dim mystr As String = "delete from sales.custcny;"
                    DbAdapter1.ExNonQuery(mystr.ToString)
                    Dim sqlstr = "copy sales.custcny(companyname,customerid) from stdin with null as 'Null';"
                    Dim errmessage = DbAdapter1.copyUTF(sqlstr, sb.ToString, myret)
                    If Not myret Then
                        Debug.WriteLine("The file could not be read:")
                        Exit Sub
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine("The file could not be read:")
                Debug.WriteLine(ex.Message)
            End Try
        End If
    End Sub
End Class