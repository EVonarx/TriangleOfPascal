Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

       



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim iDim As Integer

        iDim = TextBox2.Text

        Dim iResult(iDim) As Integer
        Dim iPreviousResult(iDim) As Integer
        Dim sString As String


        'Initialisation of the array
        iResult(0) = 1
        For iI = 1 To iDim - 1
            iResult(iI) = 0
        Next

        For iJ = 0 To iDim - 1

            'Copy the last result in the temporar variable iPreviousResult
            For iI = 0 To iDim - 1
                iPreviousResult(iI) = iResult(iI)
            Next

            'Calculate the Pascal triangle
            sString = iResult(0)
            For iI = 1 To iDim - 1
                iResult(iI) = iPreviousResult(iI) + iPreviousResult(iI - 1)
                'Store the calculated result for the calculated column
                If iResult(iI) <> 0 Then
                    sString = sString & vbTab & +iResult(iI)
                End If
            Next

            'Display the calculated result for a whole line
            TextBox1.Text = TextBox1.Text & sString & vbCrLf

        Next

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim iResult(10, 10) As Integer
        Dim sString As String = ""


        For iI = 0 To 9
            iResult(iI, 1) = 1
            iResult(iI, iI) = 1
        Next

        For iLine = 2 To 9
            For iColumn = 1 To (iLine - 2)
                iResult(iLine, iColumn) = iResult(iLine - 1, iColumn) + iResult(iLine - 1, iColumn - 1)
                sString = sString & vbTab & +iResult(iLine, iColumn)
            Next

            TextBox1.Text = TextBox1.Text & sString & vbCrLf
            sString = ""

        Next




    End Sub
End Class
