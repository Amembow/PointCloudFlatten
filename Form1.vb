Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click '点群の変換処理

        Dim fso As Object
        fso = CreateObject("Scripting.FileSystemObject")

        Dim buf As String
        Dim Arr() As String　'CSVファイル名の配列
        Dim j As Long
        Dim para As Long = 16　'並列処理数(可変)

        j = 0
        buf = Dir(TextBox3.Text & "\*.csv")　'TextBox3に記載されたフォルダ内の内、”csv”ファイルを抽出する。(拡張子可変)

        Do While buf <> ""　'ファイル名を配列に格納
            j = j + 1
            ReDim Preserve Arr(j)
            Arr(j) = buf
            Console.WriteLine(buf)
            buf = Dir()
        Loop


        Dim count As Long
        count = Arr.Length　'CSVファイル数

        Dim pos As String

        Dim k As Long
        Dim l As Long

        k = count \ para 'ファイル数を並行処理数で割ったもの、何度繰り返し処理をすればいいかの変数



        With OpenFileDialog2　'補間済みPOSのオープンダイアログ
            .Multiselect = False
            .Title = "POSの入力"
            .Filter = "posファイル|*.pos"
        End With

        Dim lineStr1 As String
        Dim SP1() As String
        Dim inputFile1 As Object
        Dim PArr(,) As String
        'Dim TempArr() As String

        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = count

        If OpenFileDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then　'補間済みPOSを配列に格納する処理

            pos = OpenFileDialog2.FileName
            inputFile1 = fso.OpenTextFile(pos, 1, False, 0)


            Dim LineCount1 As Integer = fso.OpenTextFile(pos, 8).Line   'UBound(rt.ReadToEnd.Split(Chr(13)))
            Console.WriteLine(LineCount1)

            ReDim PArr(LineCount1, 10)

            For m = 1 To LineCount1

                lineStr1 = inputFile1.ReadLine
                SP1 = lineStr1.Split(" ") 'POSの区切り文字はココ、可変
                'Console.WriteLine(SP1)
                For n = 0 To 10

                    PArr(m, n) = SP1(n)


                    'Console.WriteLine(PArr(m, n))
                Next n

            Next m

        End If



        For l = 1 To k　'並列処理。処理内容は別のプロシージャ
            Parallel.For((l - 1) * para + 1, para * l + 1,
                 Sub(i)
                     'この中が並列処理される
                     Console.WriteLine(i)
                     Dim fname As String = Arr(i)

                     Call Process(PArr, fname)　'POSの中身の配列、点群CSVファイル名、路面高毎のRGB配列が引数

                 End Sub
        )
            ProgressBar1.Value = l * para

        Next l



        Parallel.For((l - 1) * para + 1, count,　'並列処理。並列処理数で割った時のあまり分はこの中で処理する。処理内容は別のプロシージャ
                Sub(i)
                    'この中が並列処理される
                    Console.WriteLine(i)
                    Dim fname As String = Arr(i)

                    Call Process(PArr, fname)　'POSの中身の配列、点群CSVファイル名、路面高毎のRGB配列が引数

                End Sub
        )

        ProgressBar1.Value = count


        MessageBox.Show("End")

        '＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿終了処理

        fso = Nothing

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click　'POSを補完するプロシージャ

        Dim fso As Object
        fso = CreateObject("Scripting.FileSystemObject")
        Dim pos As String
        Dim inputFile As Object
        Dim lineStr As String
        Dim SP() As String
        Dim time As Double
        Dim hight As Double
        Dim i As Long
        Dim a As Long
        Dim time10 As Double
        Dim hight10 As Double
        Dim outputFile As Object
        outputFile = fso.OpenTextFile(TextBox2.Text & "\Modified.POS", 2, True)
        Dim TempArr() As String

        Dim lat As Double
        Dim lon As Double


        With OpenFileDialog2
            .Multiselect = False
            .Title = "POSの入力"
            .Filter = "posファイル|*.pos"
        End With

        If OpenFileDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then

            pos = OpenFileDialog2.FileName
            inputFile = fso.OpenTextFile(pos, 1, False, 0)

            Do Until inputFile.AtEndOfStream

                lineStr = inputFile.ReadLine

                SP = lineStr.Split("	")



                i = i + 1
                If CheckBox2.Checked = False Then

                    '横断平滑化に関しては使えない。ピッチと期首包囲の補間を行う必要がある。
                    '----------------------------------------------------------------------------------------------------------
                    If i <> 1 Then
                        time10 = CDbl(SP(0)) - time
                        hight10 = CDbl(SP(3)) - hight


                        For a = 1 To 19

                            outputFile.WriteLine(time + (time10 / 20) * a & " " & hight + (hight10 / 20) * a)
                            Console.WriteLine(time + (time10 / 20) * a & " " & hight + (hight10 / 20) * a)
                        Next a

                    End If

                    outputFile.WriteLine(SP(0) & " " & SP(3))
                    Console.WriteLine(SP(0) & " " & SP(3))
                    '--------------------------------------------------------------------------------------------------------
                Else

                    TempArr = SP(1).Split(New [Char]() {"N", "d", "m", "s", "E"})
                    lat = CStr(CDbl(TempArr(1)) + CDbl(TempArr(2)) / 60 + CDbl(TempArr(3)) / 60 / 60)

                    TempArr = SP(2).Split(New [Char]() {"N", "d", "m", "s", "E"})
                    lon = CStr(CDbl(TempArr(1)) + CDbl(TempArr(2)) / 60 + CDbl(TempArr(3)) / 60 / 60)

                    outputFile.WriteLine(SP(0) & " " & lat & " " & lon & " " & SP(3) & " " & SP(4) & " " & SP(5) & " " & SP(6) & " " & SP(7) & " " & SP(8) & " " & SP(9) & " " & SP(10))
                    'Console.WriteLine(SP(0) & " " & lat & " " & lon & " " & SP(3) & " " & SP(4) & " " & SP(5) & " " & SP(6) & " " & SP(7) & " " & SP(8) & " " & SP(9) & " " & SP(10))

                End If



                time = CDbl(SP(0))
                hight = CDbl(SP(3))



            Loop

        End If
        MessageBox.Show("End")

    End Sub

    Sub Process(PArr(,) As String, fname As String)　'点群CSVに路面高からRGB情報を付与するプロシージャ
        Dim fso As Object
        fso = CreateObject("Scripting.FileSystemObject")

        Dim outputFile As Object
        Dim inputFile As Object
        Dim lineStr As String
        Dim SP() As String

        Dim hight As Double
        Dim Dist As Double

        Dim HightDif As Double
        Dim Sys As Long = TextBox1.Text

        Dim n As Long
        Dim x As Double
        Dim y As Double
        Dim x1 As Double
        Dim y1 As Double
        Dim x2 As Double
        Dim y2 As Double
        Dim a As Double, b As Double, c As Double
        Dim angle As Double


        Dim Xr As Double


        inputFile = fso.OpenTextFile(TextBox3.Text & "\" & fname, 1, False, 0)　'
        outputFile = fso.OpenTextFile(TextBox2.Text & "\" & fname, 2, True)

        lineStr = inputFile.ReadLine
        SP = lineStr.Split(",")

        Dim flag As String

        n = 1
        Do Until inputFile.AtEndOfStream

            If CDbl(PArr(n, 0)) > CDbl(SP(4)) Then 'POSと点群CSVのGPSTimeを比較、POSのGPSTimeの方が大きい場合のみ処理を行う

                lineStr = inputFile.ReadLine
                SP = lineStr.Split(",")

                hight = CDbl(SP(2)) - CDbl(PArr(n, 3)) + 2.036

                '２点を通る１次方程式の求め方
                '点(x1,y1)と点(x2,y2)を通る直線の１次方程式は
                '(y1 - y2) * x - (x1 - x2) * y + (x1 - x2) * y1 - (y1 - y2) * x1 = 0

                y1 = calc_x(PArr(n - 1, 1), PArr(n - 1, 2), Sys)
                y2 = calc_x(PArr(n, 1), PArr(n, 2), Sys)
                x1 = calc_y(PArr(n - 1, 1), PArr(n - 1, 2), Sys)
                x2 = calc_y(PArr(n, 1), PArr(n, 2), Sys)


                a = y1 - y2
                b = (x1 - x2) * -1
                c = (x1 - x2) * y1 - (y1 - y2) * x1

                'a* x + b * y + c = 0
                'b* y = -c - a * x
                'y = -(c + a * x) / b

                'a* x = -b * y - c
                'x = -(b * y + c) / a

                '以下平滑化処理______________________________________________________________________________________________________
                If CheckBox1.Checked = True Then
                    x = calc_x(PArr(n, 1), PArr(n, 2), Sys) - CDbl(SP(0))
                    y = calc_y(PArr(n, 1), PArr(n, 2), Sys) - CDbl(SP(1))

                    Dist = Math.Sqrt(x ^ 2 + y ^ 2)

                    Xr = Math.Abs(a * CDbl(SP(1)) + b * CDbl(SP(0)) + c) / Math.Sqrt(a ^ 2 + b ^ 2)

                    angle = PArr(n, 4)




                    'アフィン変換Ver""""""""""""""""""""""""""""""""""""""""""""""""""""""

                    If PArr(n, 6) <= 90 Or PArr(n, 6) >= 270 Then
                        If CDbl(SP(1)) >= -(b * CDbl(SP(0)) + c) / a Then
                            HightDif = Affine(Xr, hight, angle)
                            flag = "a"
                        Else
                            HightDif = Affine(-Xr, hight, angle)
                            flag = "b"
                        End If
                    Else
                        If CDbl(SP(1)) >= -(b * CDbl(SP(0)) + c) / a Then
                            HightDif = Affine(-Xr, hight, angle)
                            flag = "c"
                        Else
                            HightDif = Affine(Xr, hight, angle)
                            flag = "d"
                        End If
                    End If



                    '""""""""""""""""""""""""""""""""""""""""""""""""""""""


                End If
                '以上平滑化処理______________________________________________________________________________________________________

                If HightDif <> NaN Then
                    outputFile.WriteLine(SP(0) & "," & SP(1) & "," & HightDif & "," & SP(3) & "," & SP(4) & "," & hight & "," & Dist & "," & Xr & "," & angle & "," & flag)
                End If

                '↑出力は可変
            Else

                    n = n + 1

            End If

        Loop

        fso = Nothing

    End Sub


    Dim NaN As Double

    '系番号から座標系x原点とy原点を取得するための配列
    Const TargetX As String = "0,33,33,36,33,36,36,36,36,36,40,44,44,44,26,26,26,26,20,26"
    Const TargetY As String = "0,129.5,131,132.16666666666667,133.5,134.33333333333333,136,137.16666666666667,138.5,139.83333333333333,140.83333333333333,140.25,142.25,144.25,142,127.5,124,131,136,154"

    '定数 (a, F: 世界測地系-測地基準系1980（GRS80）楕円体)
    Dim m0 As Double
    Dim a As Double
    Dim F As Double

    'セル関数として呼ばれる--平面直角座標の緯度を返す
    Public Function calc_x(phi_deg As Double, lambda_deg As Double, groupNum As Integer) As Double
        '    Attribute calc_x.VB_Description = "十進法緯度・経度を指定された平面直角座標の緯度（Ｘ座標）に変換します。"
        '    Attribute calc_x.VB_ProcData.VB_Invoke_Func = " \n14"
        calc_x = Split(calc_xy(phi_deg, lambda_deg, CDbl(Split(TargetX, ",")(groupNum)), CDbl(Split(TargetY, ",")(groupNum))), ",")(0)
    End Function

    ''セル関数として呼ばれる--平面直角座標の経度を返す
    Public Function calc_y(phi_deg As Double, lambda_deg As Double, groupNum As Integer) As Double
        '    Attribute calc_y.VB_Description = "十進法緯度・経度を指定された平面直角座標の経度（Y座標）に変換します。"
        '    Attribute calc_y.VB_ProcData.VB_Invoke_Func = " \n14"
        calc_y = Split(calc_xy(phi_deg, lambda_deg, CDbl(Split(TargetX, ",")(groupNum)), CDbl(Split(TargetY, ",")(groupNum))), ",")(1)
    End Function
    '
    Private Function calc_xy(phi_deg As Double, lambda_deg As Double, phi0_deg As Double, lambda0_deg As Double) As String
        '緯度経度を平面直角座標に変換する
        'input:
        ' (phi_deg, lambda_deg): 変換したい緯度・経度[ 十進法度]（分・秒でなく小数）
        ' (phi0_deg, lambda0_deg): 平面直角座標系原点の緯度・経度[ 十進法度]（分・秒でなく小数）
        'output:
        ' x: 変換後の平面直角座標[m]
        ' y: 変換後の平面直角座標[m]
        '緯度経度・平面直角座標系原点をラジアンに直す
        '
        NaN = CDbl(999999999)
        m0 = CDbl(0.9999)
        a = CDbl(6378137)
        F = CDbl(298.257222101)
        '
        Dim phi_rad As Double, lambda_rad As Double, phi0_rad As Double, lambda0_rad As Double
        '
        phi_rad = fnradians(phi_deg)
        lambda_rad = fnradians(lambda_deg)
        phi0_rad = fnradians(phi0_deg)
        lambda0_rad = fnradians(lambda0_deg)
        '
        '① n, A_i, alpha_iの計算
        Dim n As Double
        Dim A_array() As Double, alpha_array() As Double
        n = 1.0# / (2.0# * F - 1.0#)
        A_array = fnA_array(n)
        alpha_array = fnalpha_array(n)
        '
        '②S_, A_の計算
        Dim A_ As Double, S_ As Double
        A_ = ((m0 * a) / (1.0# + n)) * A_array(0)  'ｍ
        S_ = ((m0 * a) / (1.0# + n)) * (A_array(0) * phi0_rad + np_dot(A_array, np_sin(2.0# * phi0_rad, np_arange(1, 5)))) 'ｍ
        '
        '③ lambda_c, lambda_sの計算
        Dim lambda_c As Double, lambda_s As Double
        lambda_c = Math.Cos(lambda_rad - lambda0_rad)
        lambda_s = Math.Sin(lambda_rad - lambda0_rad)
        '
        '④t, t_の計算
        Dim t As Double, t_ As Double
        t = fnsinh(fnarctanh(Math.Sin(phi_rad)) - ((2.0# * Math.Sqrt(n)) / (1.0# + n)) * fnarctanh(((2.0# * Math.Sqrt(n)) / (1.0# + n)) * Math.Sin(phi_rad)))
        t_ = Math.Sqrt(1.0# + t * t)
        '
        ' ⑤ xi', eta'の計算
        Dim xi2 As Double, eta2 As Double
        xi2 = Math.Atan(t / lambda_c)    ' [rad]
        eta2 = fnarctanh(lambda_s / t_)
        '
        '⑥ x, yの計算
        Dim x As Double, y As Double
        x = A_ * (xi2 + np_sum(np_multiply(alpha_array, np_multiply(np_sin(2.0# * xi2, np_arange(1, 5)), np_cosh(2.0# * eta2, np_arange(1, 5)))))) - S_ ' ｍ
        y = A_ * (eta2 + np_sum(np_multiply(alpha_array, np_multiply(np_cos(2.0# * xi2, np_arange(1, 5)), np_sinh(2.0# * eta2, np_arange(1, 5))))))     ' ｍ
        '
        calc_xy = x & "," & y

    End Function
    '
    Function fnA_array(dn As Double) As Double()
        Dim Ar(5) As Double
        Ar(0) = 1.0# + (dn ^ 2) / 4.0# + (dn ^ 4) / 64.0#
        Ar(1) = -(3.0# / 2.0#) * (dn - (dn ^ 3) / 8.0# - (dn ^ 5) / 64.0#)
        Ar(2) = (15.0# / 16.0#) * (dn ^ 2 - (dn ^ 4) / 4.0#)
        Ar(3) = -(35.0# / 48.0#) * (dn ^ 3 - (5.0# / 16.0#) * (dn ^ 5))
        Ar(4) = (315.0# / 512.0#) * (dn ^ 4)
        Ar(5) = -(693.0# / 1280.0#) * (dn ^ 5)
        '
        fnA_array = Ar
    End Function
    '
    Function fnalpha_array(dn As Double) As Double()
        Dim Ar(5) As Double
        Ar(0) = NaN  'NaN
        Ar(1) = (1.0# / 2.0#) * dn - (2.0# / 3.0#) * (dn ^ 2) + (5.0# / 16.0#) * (dn ^ 3) + (41.0# / 180.0#) * (dn ^ 4) - (127.0# / 288.0#) * (dn ^ 5)
        Ar(2) = (13.0# / 48.0#) * (dn ^ 2) - (3.0# / 5.0#) * (dn ^ 3) + (557.0# / 1440.0#) * (dn ^ 4) + (281.0# / 630.0#) * (dn ^ 5)
        Ar(3) = (61.0# / 240.0#) * (dn ^ 3) - (103.0# / 140.0#) * (dn ^ 4) + (15061.0# / 26880.0#) * (dn ^ 5)
        Ar(4) = (49561.0# / 161280.0#) * (dn ^ 4) - (179.0# / 168.0#) * (dn ^ 5)
        Ar(5) = (34729.0# / 80640.0#) * (dn ^ 5)
        '
        fnalpha_array = Ar
    End Function
    '
    Function np_dot(da1() As Double, da2() As Double, Optional sta As Integer = 1) As Double
        Dim rt As Double
        Dim i As Integer
        rt = 0#
        For i = sta To UBound(da1)
            If da1(i) <> NaN And da2(i) <> NaN Then rt = rt + da1(i) * da2(i)
        Next i
        '
        np_dot = rt
    End Function
    '
    Function np_sum(da1() As Double, Optional sta As Integer = 1) As Double
        Dim rt As Double
        Dim i As Integer
        rt = 0#
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt = rt + da1(i)
        Next i
        '
        np_sum = rt
    End Function
    '
    Function np_multiply(da1() As Double, da2() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN And da2(i) <> NaN Then rt(i) = da1(i) * da2(i)
        Next i
        '
        np_multiply = rt
    End Function
    '
    Function np_cos(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = Math.Cos(dn * da1(i))
        Next i
        '
        np_cos = rt
    End Function
    '
    Function np_sin(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = Math.Sin(dn * da1(i))
        Next i
        '
        np_sin = rt
    End Function
    '
    Function np_cosh(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = fncosh(dn * da1(i))
        Next i
        '
        np_cosh = rt
    End Function
    '
    Function np_sinh(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = fnsinh(dn * da1(i))
        Next i
        '
        np_sinh = rt
    End Function
    '
    Function np_arange(sta As Integer, sto As Integer, Optional ste As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(sto)
        For i = sta To sto Step ste
            rt(i) = CDbl(i)
        Next i
        '
        np_arange = rt
    End Function
    '
    Function fnradians(dn As Double) As Double
        fnradians = dn / 45.0# * Math.Atan(1.0#)
    End Function
    '
    Function fncosh(dn As Double) As Double
        fncosh = (Math.Exp(dn) + Math.Exp(-dn)) / 2.0#
    End Function
    '
    Function fnsinh(dn As Double) As Double
        fnsinh = (Math.Exp(dn) - Math.Exp(-dn)) / 2.0#
    End Function
    '
    Function fnarctanh(dn As Double) As Double
        fnarctanh = Math.Log((1.0# + dn) / (1.0# - dn)) / 2.0#
    End Function


    Public Function Affine(ByVal x1 As Double, ByVal y1 As Double, ByVal p As Double) As Double
        Dim angle As Double, x As Double, y As Double

        angle = Math.PI * p / 180
        'x = x1 * Math.Cos(angle) - y1 * Math.Sin(angle)
        y = x1 * Math.Sin(angle) + y1 * Math.Cos(angle)

        'Return x.ToString("F6") & "," & y.ToString("F6")
        Return y
    End Function


End Class
