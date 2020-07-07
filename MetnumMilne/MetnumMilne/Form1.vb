Imports NCalc.Domain
Imports NCalc


Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' If TextBox1.Text = "" & TextBox2.Text = "" & TextBox3.Text = "" & TextBox4.Text = "" & TextBox5.Text = "" &
        '    TextBox6.Text = "" & TextBox7.Text = "" & TextBox8.Text = "" & TextBox9.Text = "" & TextBox10.Text = "" &
        '   TextBox11.Text = "" & TextBox12.Text = "" Then
        'MessageBox.Show("COMPLETE ALL THE TEXTBOX TO CONTINUE!")
        'Else
        Dim h As Single = Single.Parse(TextBox1.Text)
        Dim fungsi As Expression = New Expression(TextBox2.Text)

        Dim x(10) As Double
        Dim y(10) As Double
        Dim yEksak(10) As Double
        Dim yAksen(10) As Double

        Dim yAksen0 As Expression
        yAksen0 = fungsi
        Dim x0 As Single = Single.Parse(TextBox3.Text)
        yAksen0.Parameters.Add("x", x0)
        x(0) = x0
        Dim y0 As Single = Single.Parse(TextBox4.Text)
        yAksen0.Parameters.Add("y", y0)
        y(0) = y0
        yEksak(0) = y0
        'yAksen0.Parameters.Add("y", -4.547302)
        ' y(0) = -4.547302
        Dim yAksen0Result = yAksen0.Evaluate
        yAksen(0) = yAksen0Result

        Dim yAksen1 As Expression
        yAksen1 = fungsi
        Dim x1 As Single = Single.Parse(TextBox5.Text)
        yAksen1.Parameters.Clear()
        yAksen1.Parameters.Add("x", x1)
        x(1) = x1
        Dim y1 As Single = Single.Parse(TextBox6.Text)
        yAksen1.Parameters.Add("y", y1)
        y(1) = y1
        yEksak(1) = y1
        Dim yAksen1Result = yAksen1.Evaluate
        yAksen(1) = yAksen1Result

        Dim yAksen2 As Expression
        yAksen2 = yAksen1
        Dim x2 As Single = Single.Parse(TextBox7.Text)
        x(2) = x2
        yAksen2.Parameters.Clear()
        yAksen2.Parameters.Add("x", x2)
        Dim y2 As Single = Single.Parse(TextBox8.Text)
        yAksen2.Parameters.Add("y", y2)
        y(2) = y2
        yEksak(2) = y2
        ' yAksen2.Parameters.Add("y", -0.3929953)
        'y(2) = -0.3929953
        Dim yAksen2Result = yAksen2.Evaluate
        yAksen(2) = yAksen2Result


        Dim yAksen3 As Expression
        yAksen3 = fungsi
        Dim x3 As Single = Single.Parse(TextBox9.Text)
        x(3) = x3
        yAksen3.Parameters.Clear()
        yAksen3.Parameters.Add("x", x3)
        Dim y3 As Single = Single.Parse(TextBox10.Text)
        yAksen3.Parameters.Add("y", y3)
        y(3) = y3
        yEksak(3) = y3
        Dim yAksen3Result = yAksen3.Evaluate
        yAksen(3) = yAksen3Result
        yAksen3.Parameters.Clear()

        Dim yTrue As Single = Single.Parse(TextBox12.Text)
        yEksak(4) = yTrue

        Dim iterasi As Single = Single.Parse(TextBox11.Text)
        Dim iterasike As Integer
        iterasike = (iterasi / h) - 1

        Dim yAksenx As Expression
        For i As Integer = 3 To iterasike
            'Milne Predictor
            y(i + 1) = y(i - 3) + ((4 * h) / 3) * ((2 * yAksen(i)) - yAksen(i - 1) + (2 * yAksen(i - 2)))
            Label1.Text = y(i + 1)

            'Milne Corrector
            x(i + 1) = x(i) + h

            yAksenx = fungsi
            yAksenx.Parameters.Add("x", x(i + 1))
            yAksenx.Parameters.Add("y", y(i + 1))
            Dim yAksenxResult = yAksenx.Evaluate
            yAksenx.Parameters.Clear()
            yAksen(i + 1) = yAksenxResult
            y(i + 1) = y(i - 1) + (h / 3) * (yAksen(i - 1) + 4 * yAksen(i) + yAksen(i + 1))
            Label3.Text = y(i + 1)
        Next

        Me.Chart1.Series("Nilai Eksak").Points.AddXY(x(0), y(0))
        Me.Chart1.Series("Nilai Eksak").Points.AddXY(x(1), y(1))
        Me.Chart1.Series("Nilai Eksak").Points.AddXY(x(2), y(2))
        Me.Chart1.Series("Nilai Eksak").Points.AddXY(x(3), y(3))
        Me.Chart1.Series("Nilai Eksak").Points.AddXY(x(4), yTrue)

        Me.Chart1.Series("Nilai Milne").Points.AddXY(x(0), y(0))
        Me.Chart1.Series("Nilai Milne").Points.AddXY(x(1), y(1))
        Me.Chart1.Series("Nilai Milne").Points.AddXY(x(2), y(2))
        Me.Chart1.Series("Nilai Milne").Points.AddXY(x(3), y(3))
        Me.Chart1.Series("Nilai Milne").Points.AddXY(x(4), y(4))

        Dim eror(10) As Double
        eror(0) = Math.Abs((y(0) - (y(0))) / y(0)) * 100
        Label17.Text = eror(0)
        eror(1) = Math.Abs((y(1) - (y(1))) / y(1)) * 100
        Label20.Text = eror(1)
        eror(2) = Math.Abs((y(2) - (y(2))) / y(2)) * 100
        Label22.Text = eror(2)
        eror(3) = Math.Abs((y(3) - (y(3))) / y(3)) * 100
        Label24.Text = eror(3)
        eror(4) = Math.Abs((yTrue - (y(4))) / yTrue) * 100
        Label26.Text = eror(4)
        'End If
    End Sub

End Class
