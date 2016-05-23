Module Module1

    Public YMM(2, 0) As String
    Dim PERCENTAGE As Integer
    Dim TOTALLINES As Integer

    Public Sub LOADVEHICLES()
        Dim buffer(2) As String
        Dim i As Integer = 0


        Dim TEMPSTR() As String = IO.File.ReadAllLines("C:/SoundBoard/YMM.txt")

        For Each myLine In TEMPSTR
            ReDim Preserve YMM(2, i)
            buffer = myLine.Split(vbTab)
            YMM(0, i) = buffer(0)
            YMM(1, i) = buffer(1)
            YMM(2, i) = buffer(2)
            Console.WriteLine(i & " " & YMM(0, i) & " " & YMM(1, i) & " " & YMM(2, i))
            i += 1

        Next


    End Sub
End Module
