﻿Public Class Form1
    Dim j As Object
    Dim myworkbook As Microsoft.Office.Interop.Word.Document
    Dim myword As Microsoft.Office.Interop.Word.Application    ''表2
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook2 As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet ''临时用空EXCEL
    Dim KCH, JZND, KZDD, JGLX, HNTJLQ, SFCZDT, PMXZ, DDP, PJQT, DCCG, JCGZ, ZDCG, GKB, GZZ, GD, CS, ztnl, cpjg As String
    Dim strDirect As String
    Dim mFileInfo As System.IO.FileInfo
    Dim mDir As System.IO.DirectoryInfo
    Dim mDirInfo As System.IO.DirectoryInfo
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        OpenFileDialog2.Filter = "所有文件|*.*" ''文件筛选器
        OpenFileDialog2.ShowDialog()
        If OpenFileDialog2.FileName <> "OpenFileDialog2" Then
            TextBox3.Text = "文件已选择：" & OpenFileDialog2.FileName '1
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Filter = "所有文件|*.*" ''文件筛选器
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "OpenFileDialog1" Then
            TextBox2.Text = "文件已选择：" & OpenFileDialog1.FileName '1
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        myexcel = CreateObject("Excel.application")
        myexcel.Visible = True
        myworkbook2 = myexcel.Workbooks.Open(OpenFileDialog1.FileName)
        myworksheet = myworkbook2.Worksheets("Sheet1")
        myworkbook3 = myexcel.Workbooks.Open(OpenFileDialog2.FileName)
        myworksheet2 = myworkbook3.Worksheets("Sheet1")
        myword = CreateObject("word.application")
        myword.Visible = True
        'fs = CreateObject("Scripting.FileSystemObject")
        'a = fs.Getfolder(FolderBrowserDialog1.SelectedPath)
        'b = a.subFolders
        'xunhuan = 0
        'For Each i In b
        '    c = i.files
        '    For Each j In c
        '        If j.name Like "*表2-5*" Then
        '            table2_5()
        '            shuchu()
        '        ElseIf j.name Like "*表2-4*" Then
        '            table2_4()
        '            shuchu()
        '        End If
        '    Next
        'Next
        GetAllFiles(FolderBrowserDialog1.SelectedPath)
    End Sub

    Private Function GetAllFiles(ByVal strDirect As String) As String '搜索所有目录下的文件,运用递归，若是文件则遍历，若是文件夹则打开
        If Not (strDirect Is Nothing) Then
            mDirInfo = New System.IO.DirectoryInfo(strDirect)
            Try
                For Each mFileInfo In mDirInfo.GetFiles("*")
                    'Debug.Print(mFileInfo.FullName)
                    'Debug.Print(mFileInfo.FullName)
                    If mFileInfo.FullName Like "*表2-4*" Then
                        table2_4()
                        shuchu()
                    ElseIf mFileInfo.FullName Like "*表2-5*" Then
                        table2_5()
                        shuchu()
                    End If
                Next
                For Each mDir In mDirInfo.GetDirectories
                    GetAllFiles(mDir.FullName)
                Next
            Catch ex As System.IO.DirectoryNotFoundException
                Debug.Print("目录没找到：" + ex.Message)
            End Try
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click ''文件筛选器
        FolderBrowserDialog1.ShowDialog()
        MessageBox.Show(FolderBrowserDialog1.SelectedPath)
    End Sub

    Dim CZQH, QL, LBxs, sfczcc As String ''表2-5
    Dim myworkbook3 As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet2 As Microsoft.Office.Interop.Excel.Worksheet ''输出文件
    Dim xunhuan As Integer
    Dim fs, a, b, c As Object
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Function table2_4()
        myworkbook = myword.Documents.Open(mFileInfo.FullName)
        myworkbook.Tables(1).Select()
        myword.Selection.Copy()
        myworksheet.Activate()
        myworksheet.Range("A1").Select()
        Do While myworksheet.Range("A1").Value Is Nothing
            Try
                myexcel.ActiveSheet.PASTE
            Catch ex As Exception
            End Try
        Loop
        ''年代
        KCH = myworksheet.Range("C1").Value
        If myworksheet.Range("E4").Value Like "*R1978*" Then
            JZND = "E"
        ElseIf myworksheet.Range("E4").Value Like "*R1979*" Then
            JZND = "D"
        ElseIf myworksheet.Range("E5").Value Like "*R1990*" Then
            JZND = "C"
        ElseIf myworksheet.Range("E5").Value Like "*R2002*" Then
            JZND = "B"
        Else
            JZND = "A"
        End If
        ''地段
        If myworksheet.Range("E7").Value Like "*R不利地段*" Then
            KZDD = "D"
        ElseIf myworksheet.Range("E7").Value Like "*R危险地段*" Then
            KZDD = "E"
        ElseIf myworksheet.Range("E7").Value Like "*R一般地段*" Then
            KZDD = "B"
        Else
            KZDD = "A"
        End If
        ''结构
        If myworksheet.Range("C8").Value Like "*R砌体*" Then
            JGLX = "QT"
        ElseIf myworksheet.Range("C8").Value Like "*R内框架*" Then
            JGLX = "KJ"
        ElseIf myworksheet.Range("C8").Value Like "*R钢筋*" Then
            JGLX = "GH"
        Else
            JGLX = "OTHER"
        End If
        ''混凝土墙
        If myworksheet.Range("E9").Value Like "*R是*" Then
            PJQT = "QT"
        ElseIf myworksheet.Range("E9").Value Like "*R否*" Then
            PJQT = "KJ"
        End If
        ''高宽比
        If myworksheet.Range("E10").Value Like "*R小*" Then
            GKB = "A"
        ElseIf myworksheet.Range("E10").Value Like "*R1~*" Then
            GKB = "B"
        ElseIf myworksheet.Range("E10").Value Like "*R2~*" Then
            GKB = "C"
        ElseIf myworksheet.Range("E10").Value Like "*R4~*" Then
            GKB = "D"
        ElseIf myworksheet.Range("E10").Value Like "*R大*" Then
            GKB = "E"
        End If
        ''多塔
        If myworksheet.Range("E11").Value Like "*R有*" Then
            GZZ = "B"
        Else
            GZZ = "E"
        End If
        ''平面形状
        If myworksheet.Range("E12").Value Like "*R矩形*" Then
            GZZ = "A"
        ElseIf myworksheet.Range("E12").Value Like "*R正*" Then
            GZZ = "B"
        ElseIf myworksheet.Range("E12").Value Like "*R十字*" Then
            GZZ = "C"
        ElseIf myworksheet.Range("E12").Value Like "*RL形*" Then
            GZZ = "D"
        ElseIf myworksheet.Range("E12").Value Like "*RU*" Then
            GZZ = "E"
        End If
        ''大底盘
        If myworksheet.Range("E13").Value Like "*R不" Then
            CZQH = "B"
        ElseIf myworksheet.Range("E13").Value Like "*R存*" Then
            CZQH = "D"
        End If
        ''配筋砌体
        If myworksheet.Range("E14").Value Like "*R否*" Then
            PJQT = "D"
        Else
            PJQT = "A"
        End If
        ''底层拆改
        If myworksheet.Range("E15").Value Like "*R否*" Then
            DCCG = "B"
        ElseIf myworksheet.Range("E15").Value Like "*R是*" Then
            DCCG = "A"
        End If
        ''加层改造
        If myworksheet.Range("E16").Value Like "*R无加层*" Then
            JCGZ = "B"
        ElseIf myworksheet.Range("E16").Value Like "*R加一层*" Then
            JCGZ = "C"
        ElseIf myworksheet.Range("E16").Value Like "*R加两层*" Then
            JCGZ = "E"
        End If
        ''最大层高
        If myworksheet.Range("E17").Value Like "*R2.8*" Then
            ZDCG = "B"
        ElseIf myworksheet.Range("E17").Value Like "*R3.3*" Then
            ZDCG = "C"
        ElseIf myworksheet.Range("E17").Value Like "*R3.6*" Then
            ZDCG = "D"
        ElseIf myworksheet.Range("E17").Value Like "*大于*" Then
            ZDCG = "E"
        End If
        ''高宽比
        If myworksheet.Range("E18").Value Like "*R小*" Then
            GKB = "A"
        ElseIf myworksheet.Range("E18").Value Like "*R1~*" Then
            GKB = "B"
        ElseIf myworksheet.Range("E18").Value Like "*R1.5*" Then
            GKB = "C"
        ElseIf myworksheet.Range("E18").Value Like "*R2~*" Then
            GKB = "D"
        ElseIf myworksheet.Range("E18").Value Like "*R大*" Then
            GKB = "E"
        End If
        ''构造柱
        If myworksheet.Range("E19").Value Like "*R有*" Then
            GZZ = "B"
        ElseIf myworksheet.Range("E19").Value Like "*R无*" Then
            GZZ = "E"
        End If
        ''高度
        If myworksheet.Range("E20").Value Like "*R9*" Then
            GD = "A"
        ElseIf myworksheet.Range("E20").Value Like "*R12*" Then
            GD = "B"
        ElseIf myworksheet.Range("E20").Value Like "*R15*" Then
            GD = "C"
        ElseIf myworksheet.Range("E20").Value Like "*R18*" Then
            GD = "D"
        ElseIf myworksheet.Range("E20").Value Like "*R21*" Then
            GD = "E"
        End If
        ''层数
        If myworksheet.Range("E21").Value Like "*R3*" Then
            CS = "A"
        ElseIf myworksheet.Range("E21").Value Like "*R4*" Then
            CS = "B"
        ElseIf myworksheet.Range("E21").Value Like "*R5*" Then
            CS = "C"
        ElseIf myworksheet.Range("E21").Value Like "*R6*" Then
            CS = "D"
        ElseIf myworksheet.Range("E21").Value Like "*R7*" Then
            CS = "E"
        End If
        ''承重墙厚
        If myworksheet.Range("E22").Value Like "*R370*" Then
            CZQH = "A"
        ElseIf myworksheet.Range("E22").Value Like "*R240*" Then
            CZQH = "B"
        ElseIf myworksheet.Range("E22").Value Like "*R190*" Then
            CZQH = "C"
        ElseIf myworksheet.Range("E22").Value Like "*R小于190*" Then
            CZQH = "D"
        End If
        ''圈梁
        If myworksheet.Range("E23").Value Like "*R无*" Then
            QL = "E"
        Else
            QL = "A"
        End If
        ''楼板形式
        If myworksheet.Range("E24").Value Like "*R现浇板*" Then
            LBxs = "A"
        ElseIf myworksheet.Range("E24").Value Like "*R预制板*" Then
            LBxs = "D"
        ElseIf myworksheet.Range("E24").Value Like "*R木屋架*" Then
            LBxs = "E"
        End If
        ''存在错层
        If myworksheet.Range("E25").Value Like "*R不存在*" Then
            sfczcc = "B"
        Else
            sfczcc = "D"
        End If
        ''整体能力
        If myworksheet.Range("C26").Value Like "*R良*" Then
            cpjg = "A"
        ElseIf myworksheet.Range("C26").Value Like "*R中*" Then
            cpjg = "B"
        Else
            cpjg = "C"
        End If
    End Function
    Function table2_5()
        myworkbook = myword.Documents.Open(mFileInfo.FullName)
        myworkbook.Tables(1).Select()
        myword.Selection.Copy()
        myworksheet.Activate()
        myworksheet.Range("A1").Select()
        myexcel.ActiveSheet.PASTE
        KCH = myworksheet.Range("D1").Value
        If myworksheet.Range("G4").Value Like "*R1978*" Then
            JZND = "E"
        ElseIf myworksheet.Range("G4").Value Like "*R1979*" Then
            JZND = "D"
        ElseIf myworksheet.Range("G5").Value Like "*R1990*" Then
            JZND = "C"
        ElseIf myworksheet.Range("G5").Value Like "*R2002*" Then
            JZND = "B"
        Else
            JZND = "A"
        End If
        If myworksheet.Range("G7").Value Like "*R不利地段*" Then
            KZDD = "D"
        ElseIf myworksheet.Range("G7").Value Like "*R危险地段*" Then
            KZDD = "E"
        ElseIf myworksheet.Range("G7").Value Like "*R一般地段*" Then
            KZDD = "B"
        Else
            KZDD = "A"
        End If
        JGLX = "ZJF"
        If myworksheet.Range("G9").Value Like "*R否*" Then
            PJQT = "D"
        Else
            PJQT = "A"
        End If
        If myworksheet.Range("G10").Value Like "*R否*" Then
            DCCG = "B"
        Else
            DCCG = "A"
        End If
        If myworksheet.Range("G11").Value Like "*R无加层*" Then
            JCGZ = "B"
        ElseIf myworksheet.Range("G11").Value Like "*R加一层*" Then
            JCGZ = "C"
        ElseIf myworksheet.Range("G11").Value Like "*R加两层*" Then
            JCGZ = "E"
        End If
        If myworksheet.Range("G12").Value Like "*R2.8*" Then
            ZDCG = "B"
        ElseIf myworksheet.Range("G12").Value Like "*R3.3*" Then
            ZDCG = "C"
        ElseIf myworksheet.Range("G12").Value Like "*R3.6*" Then
            ZDCG = "D"
        ElseIf myworksheet.Range("G12").Value Like "*大于*" Then
            ZDCG = "E"
        End If
        If myworksheet.Range("G13").Value Like "*R*无" Then
            GZZ = "E"
        Else
            GZZ = "B"
        End If
        If myworksheet.Range("G14").Value Like "*R9*" Then
            GD = "A"
        ElseIf myworksheet.Range("G14").Value Like "*R12*" Then
            GD = "B"
        ElseIf myworksheet.Range("G14").Value Like "*R15*" Then
            GD = "C"
        ElseIf myworksheet.Range("G14").Value Like "*R18*" Then
            GD = "D"
        ElseIf myworksheet.Range("G14").Value Like "*R21*" Then
            GD = "E"
        End If
        If myworksheet.Range("G15").Value Like "*R3*" Then
            CS = "A"
        ElseIf myworksheet.Range("G15").Value Like "*R4*" Then
            CS = "B"
        ElseIf myworksheet.Range("G15").Value Like "*R5*" Then
            CS = "C"
        ElseIf myworksheet.Range("G15").Value Like "*R6*" Then
            CS = "D"
        ElseIf myworksheet.Range("G15").Value Like "*R7*" Then
            CS = "E"
        End If
        If myworksheet.Range("G16").Value Like "*R370*" Then
            CZQH = "A"
        ElseIf myworksheet.Range("G16").Value Like "*R240*" Then
            CZQH = "B"
        ElseIf myworksheet.Range("G16").Value Like "*R190*" Then
            CZQH = "C"
        ElseIf myworksheet.Range("G16").Value Like "*R小于190*" Then
            CZQH = "D"
        End If
        If myworksheet.Range("G17").Value Like "*R无*" Then
            QL = "E"
        Else
            QL = "A"
        End If
        If myworksheet.Range("G18").Value Like "*R现浇板*" Then
            LBxs = "A"
        ElseIf myworksheet.Range("G18").Value Like "*R预制板*" Then
            LBxs = "D"
        ElseIf myworksheet.Range("G18").Value Like "*R木屋架*" Then
            LBxs = "E"
        End If
        If myworksheet.Range("G19").Value Like "*R不存在*" Then
            sfczcc = "B"
        Else
            sfczcc = "D"
        End If
        If myworksheet.Range("G20").Value Like "*R良*" Then
            cpjg = "A"
        ElseIf myworksheet.Range("G20").Value Like "*R中*" Then
            cpjg = "B"
        Else
            cpjg = "C"
        End If
    End Function
    Function shuchu()
        myworksheet2.Range("A" & 5 + xunhuan).Value = KCH
        myworksheet2.Range("F" & 5 + xunhuan).Value = JGLX
        myworksheet2.Range("G" & 5 + xunhuan).Value = JZND
        myworksheet2.Range("H" & 5 + xunhuan).Value = KZDD
        myworksheet2.Range("I" & 5 + xunhuan).Value = PJQT
        myworksheet2.Range("J" & 5 + xunhuan).Value = DCCG
        myworksheet2.Range("K" & 5 + xunhuan).Value = JCGZ
        myworksheet2.Range("L" & 5 + xunhuan).Value = ZDCG
        myworksheet2.Range("M" & 5 + xunhuan).Value = GKB
        myworksheet2.Range("N" & 5 + xunhuan).Value = GZZ
        myworksheet2.Range("O" & 5 + xunhuan).Value = GD
        myworksheet2.Range("P" & 5 + xunhuan).Value = CS
        myworksheet2.Range("Q" & 5 + xunhuan).Value = PMXZ
        myworksheet2.Range("R" & 5 + xunhuan).Value = CZQH
        myworksheet2.Range("S" & 5 + xunhuan).Value = QL
        myworksheet2.Range("T" & 5 + xunhuan).Value = LBxs
        myworksheet2.Range("U" & 5 + xunhuan).Value = sfczcc
        myworksheet2.Range("V" & 5 + xunhuan).Value = cpjg
        xunhuan = xunhuan + 1
        myworksheet.Activate()
        myworksheet.Cells.Select()
        myexcel.Selection.delete
        KCH = ""
        JGLX = ""
        JZND = ""
        KZDD = ""
        PJQT = ""
        DCCG = ""
        JCGZ = ""
        ZDCG = ""
        GKB = ""
        GZZ = ""
        GD = ""
        CS = ""
        PMXZ = ""
        CZQH = ""
        QL = ""
        LBxs = ""
        sfczcc = ""
        cpjg = ""
        myworkbook.Close()
    End Function
End Class
