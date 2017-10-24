
Imports System
Imports System.IO
Imports System.Collections
Imports MapWinGIS

Public Class Form1
    Inherits System.Windows.Forms.Form
    Implements MapWinGIS.ICallback
    Dim shpManhol As String
    Dim shpPipe As String
    Dim shpBlock As String
    Dim MHList As New ArrayList
    Dim IntPnt As New ArrayList
    Dim Level1 As New ArrayList
    Dim arr As New ArrayList
    Dim Count As Integer
    Dim Rpt(400) As MapWinGIS.Point
    Dim F As Boolean
    Dim Man_No, Main_Man As Single
    Dim Que As Queue(Of Object)
    Dim QueueNumber As Queue(Of Object)
    Dim StackNumber As Stack(Of Object)
    Dim AreaList As Queue(Of Object)
    Dim FN As Integer
    Dim FID(10000) As Integer
    Public Fd2_Indx As Integer
    Dim NetName As String
    Dim Ic As Integer
    Dim Man_Str(2) As String
    Dim Man2_Str(2) As String
    Dim xxx, yyy, xx2, yy2 As Double
    Dim xxxStr As String
    Dim Man() As Double
    Dim xx(2), yy(2) As Double
    Dim Tol As Double
    Dim Elev_No As Integer
    Dim Total As Integer
    Dim NoFd As Integer
    Dim AreaFd2_Indx As Integer
    Dim TotalArea As Double
    Dim PipeNo As Integer
    Dim LyrHnd1 As Integer
    Dim LyrHnd2 As Integer
    Dim LyrHnd0 As Integer
    Dim LyrHnd3 As Integer
    Dim Err As Boolean
    Dim l() As Single
    

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End

    End Sub


    Private Sub ExitItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitItem.Click

        End
    End Sub

    Private Sub ReadShp()
        Dim sf1 As New MapWinGIS.Shapefile
        Dim pt As New MapWinGIS.Point
        sf1.Open(shpManhol)

        Dim i As Integer

        Dim Dk(10000) As Double
        Dim st As String
        Dim j As Integer = 0
        ' For Each manhole in Main Line


        For i = 0 To sf1.NumShapes - 1
            If (sf1.CellValue(FN, i).ToString = "" Or sf1.CellValue(FN, i).ToString = "0" Or Len(sf1.CellValue(FN, i).ToString) > 5) Then GoTo A10
            st = sf1.CellValue(FN, i)
            Dk(j) = Convert.ToDouble(st)
            FID(j) = i

            j += 1
A10:

        Next
        Main_Man = Int(Dk(j - 1)) + 1
        Ascending(Dk, j - 1)
        sf1.StartEditingTable()
        For i = 0 To j - 1
            pt = sf1.QuickPoint(FID(i), 0)
            MHList.Add(pt)
            IntPnt.Add(pt)
            st = sf1.CellValue(FN, FID(i))

            st = NetName + st
            sf1.EditCellValue(Fd2_Indx, FID(i), st)
        Next
        sf1.StopEditingTable()
        sf1.Close()

    End Sub

    Private Sub Ascending(ByVal D() As Double, ByVal num As Integer)
        Dim i, j As Integer
        Dim Temp1 As Double
        Dim Temp2 As Integer
        For i = 0 To num
            For j = 0 To num
                If D(i) < D(j) Then
                    Temp1 = D(i)
                    D(i) = D(j)
                    D(j) = Temp1
                    Temp2 = FID(i)
                    FID(i) = FID(j)
                    FID(j) = Temp2
                End If

            Next
        Next

      

    End Sub


    Private Function AngleF(ByVal p1 As MapWinGIS.Point, ByVal p2 As MapWinGIS.Point, ByVal Center As MapWinGIS.Point) As Double

       
        Dim angle1 As Double = Math.Atan2(p1.y - Center.y, p1.x - Center.x) * 180 / Math.PI
        Dim angle2 As Double = Math.Atan2(p2.y - Center.Y, p2.x - Center.X) * 180 / Math.PI
        AngleF = (angle1 - angle2)

    End Function
    Private Function Check(ByVal pt As MapWinGIS.Point) As Boolean
        Dim p As New MapWinGIS.Point
        For Each p In IntPnt
            If (pt.x = p.x And pt.y = p.y) Then
                Check = True
                Exit Function
            Else
                Check = False
            End If
        Next
    End Function
    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        AxMap1.Refresh()
        Que.Clear()
        QueueNumber.Clear()
        IntPnt.Clear()
        Err = False
        Count = 0
        Rpt(0) = Nothing
        Rpt(1) = Nothing
        AxMap1.RemoveLayer(LyrHnd1)
        AxMap1.RemoveLayer(LyrHnd2)
        TotalArea = 0
        AxMap1.ClearLabels(LyrHnd3)
        AxMap1.ClearLabels(LyrHnd2)
        AxMap1.ClearLabels(LyrHnd1)
        TComboBox1.Items.Clear()
    End Sub


    Private Sub LoadShapeFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadShapeFileToolStripMenuItem.Click
        Dim i, k As Integer
        Dim strName(2) As String
        Dim sf As New MapWinGIS.Shapefile
        Dim Shp As New MapWinGIS.Shapefile


        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Shapefiles: (*.shp)|*.shp"
        k = 0

        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            strName = OpenFileDialog1.FileNames
            i = OpenFileDialog1.FileNames.Length

            Select Case i
                Case 1
                    MessageBox.Show("Must be 2 Layers")
                    Exit Sub
                Case 2
                    sf.Open(strName(0))
                    If sf.ShapefileType = MapWinGIS.ShpfileType.SHP_POINT Then
                        shpManhol = strName(0)
                        shpPipe = strName(1)

                    Else
                        shpManhol = strName(1)
                        shpPipe = strName(0)


                    End If
                    Shp.Open(shpManhol)
                    For i = 0 To Shp.NumFields - 1
                        TComboBox1.Items.Add(Shp.Field(i).Name)
                        ToolStripComboBox1.Items.Add(Shp.Field(i).Name)
                    Next
                    Shp.Close()
                    Shp.Open(shpPipe)
                    PipeNo = Shp.NumShapes
                    Shp.Close()
                   
                Case Is > 2
                    MessageBox.Show("Must be 2 Layers")
                    Exit Sub
            End Select


            sf.Close()
            Err = False
            Add_Map()


        End If

    End Sub

    Private Sub Add_BaseMap()
        Dim sfb As New MapWinGIS.Shapefile

        sfb.Open(shpBlock, Me)

        Dim hnd As Integer = AxMap1.AddLayer(sfb, True)
        AxMap1.set_ShapeLayerFillColor(hnd, RGB(255, 255, 0))
        ' sfb.Close()

    End Sub

    Private Sub Add_Map()

        Dim i As Integer
        Dim sf1 As New MapWinGIS.Shapefile
        Dim sf2 As New MapWinGIS.Shapefile
       
        sf1.Open(shpManhol, Me)
        sf2.Open(shpPipe, Me)

        LyrHnd1 = AxMap1.AddLayer(sf1, True)
        LyrHnd2 = AxMap1.AddLayer(sf2, True)
        Dim pt As New MapWinGIS.Point
        Dim hnd As Integer = AxMap1.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
        For i = 0 To sf1.NumShapes - 1
            pt = sf1.QuickPoint(i, 0)
            AxMap1.DrawCircleEx(hnd, pt.x, pt.y, 4, RGB(0, 0, 255), True)
        Next

        AxMap1.Refresh()
        AxMap1.Redraw()
        AxMap1.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
        AxMap1.set_ShapeLayerFillColor(LyrHnd1, RGB(0, 0, 255))
        AxMap1.set_ShapeLayerLineColor(LyrHnd2, RGB(255, 0, 0))
        AxMap1.ZoomToLayer(LyrHnd2)
         
    End Sub



    Public Sub [Error](ByVal KeyOfSender As String, ByVal ErrorMsg As String) Implements MapWinGIS.ICallback.Error

    End Sub

    Public Sub Progress(ByVal KeyOfSender As String, ByVal Percent As Integer, ByVal Message As String) Implements MapWinGIS.ICallback.Progress

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Err = False
        Count = 0
        Man_No = 0.01
        Que = New Queue(Of Object)()
        QueueNumber = New Queue(Of Object)()
        StackNumber = New Stack(Of Object)()
        AreaList = New Queue(Of Object)()
        TextBox1.Text = Settings1.Default.Net
        'TextBox2.Text = Settings1.Default.BaseMap
        'shpBlock = TextBox2.Text
        'If shpBlock <> "" Then Add_BaseMap()
        NetName = Settings1.Default.Net + "."
        AxMap1.CursorMode = MapWinGIS.tkCursorMode.cmNone
        Tol = 5
        TotalArea = 0


       

    End Sub


    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        AxMap1.Width = Me.Width - 10
        AxMap1.Height = Me.Height - Menu1.Height - 5

    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim pt As New MapWinGIS.Point
        Dim k As Integer

        CreatField()
        ReadShp()
        Err = False

        For Each pt In MHList
            k = GetLevel1(pt)

            Select Case k
                Case 1
                     
                    Que.Enqueue(Rpt(0))
                    IntPnt.Add(Rpt(0))

                Case 2

                    Dim ind As Integer = MHList.IndexOf(pt)
                    Dim pt2 As MapWinGIS.Point = MHList.Item(ind)
                    Dim ang1 As Double = AngleF(Rpt(0), pt2, pt)
                    Dim ang2 As Double = AngleF(Rpt(1), pt2, pt)

                    ' If ang1 < 0 Then     ' Pipe in Right Side

                    Que.Enqueue(Rpt(0))
                    IntPnt.Add(Rpt(0))

                    Que.Enqueue(Rpt(1))
                    IntPnt.Add(Rpt(1))
                    '   End If
                    '  If ang2 > 0 Then

                    'Que.Enqueue(Rpt(1))
                    '   IntPnt.Add(Rpt(1))
                    'Que.Enqueue(Rpt(0))
                    ' IntPnt.Add(Rpt(0))

                    'End If

                Case 3
                    Que.Enqueue(Rpt(0))
                    IntPnt.Add(Rpt(0))

                    Que.Enqueue(Rpt(1))
                    IntPnt.Add(Rpt(1))

                    Que.Enqueue(Rpt(2))
                    IntPnt.Add(Rpt(2))


                Case 4
                    Que.Enqueue(Rpt(0))
                    IntPnt.Add(Rpt(0))

                    Que.Enqueue(Rpt(1))
                    IntPnt.Add(Rpt(1))

                    Que.Enqueue(Rpt(2))
                    IntPnt.Add(Rpt(2))
                    Que.Enqueue(Rpt(3))
                    IntPnt.Add(Rpt(3))
            End Select

          

        Next

        Timer1.Enabled = True
        Main_Program()
        AddLabel()
        Timer1.Enabled = False
        ProgressBar1.Value = Total
        IntPnt.Clear()
        TotalArea = 0
        PipeConnection()
        MessageBox.Show("Finshed..")
    End Sub

    Private Sub CalculateArea()
        Dim sf As New MapWinGIS.Shapefile

        Dim pt(2) As MapWinGIS.Point
        Dim i, j As Integer
        Dim st As String
        Dim k, kk As Integer
        Dim fl(2) As Double

        sf.Open(shpManhol)
        j = 0
        kk = 0
        TotalArea = 0
        ReDim fl(sf.NumShapes)
        ReDim pt(sf.NumShapes)
        ProgressBar1.Maximum = sf.NumShapes + 1
        For i = 0 To sf.NumShapes - 1
            st = sf.CellValue(Fd2_Indx, i)
            k = st.IndexOf(".")
            fl(i) = Convert.ToDouble((Mid(st, k + 2, Len(st))))
            ProgressBar1.Value = i
        Next


        For j = 0 To Int(Main_Man) - 1
            For i = 0 To sf.NumShapes - 1
                If Int(fl(i)) = j Then
                    pt(kk) = sf.QuickPoint(i, 0)
                    kk += 1
                End If

            Next
            If kk = 1 Then
                pt(1) = PointInPipe(pt(0))
                Define2Area(pt, kk)
                kk = 0
            Else
                DefineArea(pt, kk - 1)
                kk = 0
            End If
            ProgressBar1.Value = j
        Next

        sf.Close()
    End Sub

    Private Function PointInPipe(ByVal pt As MapWinGIS.Point) As MapWinGIS.Point
        Dim sf As New MapWinGIS.Shapefile
        Dim pt1 As New MapWinGIS.Point
        Dim pt2 As New MapWinGIS.Point
        Dim i As Integer
        sf.Open(shpPipe)

        For i = 0 To sf.NumShapes - 1
            pt1 = sf.QuickPoint(i, 0)
            pt2 = sf.QuickPoint(i, 1)
            If pt.x = pt1.x And pt.y = pt1.y Then
                PointInPipe = pt2

            End If


            If pt.x = pt2.x And pt.y = pt2.y Then
                PointInPipe = pt1
            End If
            ProgressBar1.Value = i
        Next
        sf.Close()
    End Function

    Private Sub Define2Area(ByVal pt() As MapWinGIS.Point, ByVal k As Integer)
        Dim sfb As New MapWinGIS.Shapefile
        Dim Ext As New MapWinGIS.Extents
        Dim i, j, k2, Fid As Integer
        Dim selected() As Integer
        Dim Areaf, aSum, dx, dy As Double
        sfb.Open(shpBlock)
        aSum = 0
        Areaf = 0
        AxMap1.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
        Ext.SetBounds(pt(0).x, pt(0).y, 0, pt(1).x, pt(1).y, 0)
        sfb.SelectShapes(Ext, 10, SelectMode.INTERSECTION, selected)
        k2 = UBound(selected)
        If k2 >= 2 Then k2 = 1
        For j = 0 To k2
            Fid = selected(j)
            Areaf += Area(Fid)
        Next
        aSum += Areaf
        TotalArea += aSum
        Application.DoEvents()
        sfb.Close()
        Dim m As Double
        Dim Id As Integer
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpManhol)
        sf.StartEditingTable()
        m = aSum
        m = m / (k)
        Id = MMID(pt(0))
        sf.EditCellValue(AreaFd2_Indx, Id, m.ToString("0.00"))
        sf.StopEditingTable()
        sf.Close()
    End Sub

    Private Sub DefineArea(ByVal pt() As MapWinGIS.Point, ByVal k As Integer)
        Dim sfb As New MapWinGIS.Shapefile
        Dim Ext As New MapWinGIS.Extents
        Dim i, j, k2, Fid As Integer
        Dim selected() As Integer
        Dim Areaf, aSum As Double

        sfb.Open(shpBlock)
        aSum = 0
        Areaf = 0
        For i = 0 To k - 1
            

            Ext.SetBounds(pt(i).x, pt(i).y, 0, pt(i + 1).x, pt(i + 1).y, 0)

            sfb.SelectShapes(Ext, 10.0, SelectMode.INTERSECTION, selected)
            k2 = UBound(selected)
            If k2 >= 2 Then k2 = 1
            For j = 0 To k2
                Fid = selected(j)
                Areaf += Area(Fid)
            Next
            aSum += Areaf
            TotalArea += aSum
            Application.DoEvents()
        Next
        sfb.Close()

        Dim m As Double
        Dim Id As Integer
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpManhol)
        sf.StartEditingTable()
        m = aSum
        m = m / (k)
        For i = 0 To k
            Id = MMID(pt(i))
            sf.EditCellValue(AreaFd2_Indx, Id, m.ToString("0.000"))
            Application.DoEvents()
        Next
        sf.StopEditingTable()
        sf.Close()
    End Sub


    Private Sub C2Area(ByVal pt1 As MapWinGIS.Point, ByVal pt2 As MapWinGIS.Point)
        Dim sfb As New MapWinGIS.Shapefile
        Dim Ext As New MapWinGIS.Extents
        Dim i, j, k2, Fid As Integer
        Dim selected() As Integer
        sfb.Open(shpBlock)
        
        Ext.SetBounds(pt1.x, pt1.y, 0, pt2.x, pt2.y, 0)

        sfb.SelectShapes(Ext, 10.0, SelectMode.INTERSECTION, selected)
        k2 = UBound(selected)
        If k2 >= 2 Then k2 = 1
        For j = 0 To k2
            Fid = selected(j)
            TotalArea += Area(Fid)
        Next
        Application.DoEvents()

        sfb.Close()
 
    End Sub





    Private Sub AddLabel()
        Dim sf As New MapWinGIS.Shapefile
        Dim i As Integer
        Dim pt As New MapWinGIS.Point
        Dim hnd As Integer = AxMap1.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
        sf.Open(shpManhol)
        For i = 0 To sf.NumShapes - 1
            pt = sf.QuickPoint(i, 0)
            AxMap1.DrawCircleEx(hnd, pt.x, pt.y, 4, RGB(0, 0, 255), True)
            AxMap1.AddLabel(0, sf.CellValue(Fd2_Indx, i), RGB(0, 0, 255), pt.x, pt.y, tkHJustification.hjLeft)
        Next

        sf.Close()
    End Sub


    Private Sub Main_Program()
        Dim pt, P1, P2, P3, P4 As New MapWinGIS.Point
        Dim k, j, i As Integer
        Dim fd As New MapWinGIS.Field
        Dim sf As New MapWinGIS.Shapefile
        Dim sBlock As New MapWinGIS.Shapefile
        AxMap1.ClearDrawings()
        sf.Open(shpManhol)
        sBlock.Open(shpBlock)
        sf.StartEditingTable()
        If CheckBox1.Checked = True Then
            Main_Man = NumericUpDown1.Value
        Else
            Main_Man = 1
        End If



        For i = 0 To Que.Count - 1
            pt = Que.Dequeue
            SetPrior(pt, sf)
            Man_No = 0.01
            Main_Man += 1
        Next


        sf.StopEditingTable()
        sf.Close()
        sBlock.Close()
    End Sub

    Private Sub AddNumber(ByVal pt As MapWinGIS.Point, ByVal sf As MapWinGIS.Shapefile)
        Dim kk As Integer
        Dim st As String
        If Err = False Then
            kk = MMID(pt)
            st = NetName + (Man_No + Main_Man).ToString("0.00")
            sf.EditCellValue(Fd2_Indx, kk, st)
        End If

    End Sub

    Private Function GetNo(ByVal pt As MapWinGIS.Point) As Integer
        Dim sf As New MapWinGIS.Shapefile
        Dim Ext As New MapWinGIS.Extents
        Dim selectedShapes() As Integer
        sf.Open(shpPipe)
        Ext.SetBounds(pt.x, pt.y, 0, pt.x, pt.y, 0)
        Dim f As Boolean = sf.SelectShapes(Ext, 0.1, SelectMode.INTERSECTION, selectedShapes)
        GetNo = UBound(selectedShapes)

        sf.Close()
    End Function

    Private Sub GetPointIntersection(ByVal pt As MapWinGIS.Point)
        Dim sf As New MapWinGIS.Shapefile
        Dim Temp As New MapWinGIS.Point
        Dim Ext As New MapWinGIS.Extents
        Dim P1, P2 As New MapWinGIS.Point
        Dim selectedShapes() As Integer
        Dim j As Integer
        sf.Open(shpPipe)
        Ext.SetBounds(pt.x, pt.y, 0, pt.x, pt.y, 0)
        Dim i As Integer = 0
        Dim f As Boolean = sf.SelectShapes(Ext, 0.1, SelectMode.INTERSECTION, selectedShapes)
        If f = False Then
            MessageBox.Show(f.ToString + "    " + pt.x.ToString + "     " + pt.y.ToString + "    ")

        End If
       
        If f = True Then

            For j = 0 To UBound(selectedShapes)

                Temp = sf.QuickPoint(selectedShapes(j), 0)

                If (Check(Temp) = False) Then
                    P1 = sf.QuickPoint(selectedShapes(j), 0)
                    QueueNumber.Enqueue(P1)
                    IntPnt.Add(P1)

                End If

                Temp = sf.QuickPoint(selectedShapes(j), 1)

                If (Check(Temp) = False) Then
                    P2 = sf.QuickPoint(selectedShapes(j), 1)
                    QueueNumber.Enqueue(P2)
                    IntPnt.Add(P2)

                End If

            Next
        End If

        sf.Close()
    End Sub

    Private Function GetLevel1(ByVal pt As MapWinGIS.Point) As Integer
        Dim pold As New MapWinGIS.Point
        Dim sf As New MapWinGIS.Shapefile

        Dim Ext As New MapWinGIS.Extents
        Dim selectedShapes() As Integer
        Dim j As Integer
        sf.Open(shpPipe)

        Count = 0
        Ext.SetBounds(pt.x, pt.y, 0, pt.x, pt.y, 0)
        Dim f As Boolean = sf.SelectShapes(Ext, 10, SelectMode.INTERSECTION, selectedShapes)
        If f = False Then
            MessageBox.Show(f.ToString + "    " + pt.x.ToString + "     " + pt.y.ToString + "    ")

        End If
        If f = True Then

            For j = 0 To UBound(selectedShapes)

                pold = sf.QuickPoint(selectedShapes(j), 0)
                If (Check(pold) = False) Then
                    IntPnt.Add(pold)
                    Rpt(Count) = pold
                    Count += 1
                End If

                pold = sf.QuickPoint(selectedShapes(j), 1)
                If (Check(pold) = False) Then
                    IntPnt.Add(pold)
                    Rpt(Count) = pold
                    Count += 1
                End If

            Next

        End If
        GetLevel1 = Count
        sf.Close()
    End Function

    Private Function Slope(ByVal pt0 As MapWinGIS.Point, ByVal pt1 As MapWinGIS.Point) As Single

        Dim Dx, Dy As Single
        Dx = Math.Abs(pt0.x - pt1.x)
        Dy = Math.Abs(pt0.y - pt1.y)
        If Dx = 0 Then Dx = 1
        Slope = Dy / Dx

    End Function

    Private Function MMID(ByVal pt As MapWinGIS.Point) As Integer
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpManhol)
        Dim selectedshp() As Integer
        Dim Ext As New MapWinGIS.Extents
        Dim j As Integer
        Ext.SetBounds(pt.x, pt.y, 0, pt.x, pt.y, 0)

        Dim f As Boolean = sf.SelectShapes(Ext, 2.5, SelectMode.INTERSECTION, selectedshp)
        If f = False Then
            MessageBox.Show(f.ToString + "    " + pt.x.ToString + "     " + pt.y.ToString + "    ")
            Application.Exit()
        End If

        MMID = selectedshp(0)
        sf.Close()
    End Function
    Private Function Find3ID(ByVal pt As MapWinGIS.Point, ByVal p1 As MapWinGIS.Point, ByVal p2 As MapWinGIS.Point, ByVal p3 As MapWinGIS.Point) As MapWinGIS.Point
        Dim sf As New MapWinGIS.Shapefile
        Dim pold As New MapWinGIS.Point
        sf.Open(shpPipe)
        Dim selectedshp() As Integer
        Dim Ext As New MapWinGIS.Extents
        Dim j As Integer
        Ext.SetBounds(pt.x, pt.y, 0, pt.x, pt.y, 0)
        Dim f As Boolean = sf.SelectShapes(Ext, 0.5, SelectMode.INTERSECTION, selectedshp)

        arr.Add(pt)
        arr.Add(p1)
        arr.Add(p2)
        arr.Add(p3)
        For j = 0 To UBound(selectedshp)
            pold = sf.QuickPoint(selectedshp(j), 0)
            If CheckList(pold) = False Then
                Find3ID = pold
            End If
            pold = sf.QuickPoint(selectedshp(j), 1)
            If CheckList(pold) = False Then
                Find3ID = pold
            End If
        Next
        sf.Close()
        arr.Clear()
    End Function

    Private Function FindID(ByVal pt As MapWinGIS.Point, ByVal p1 As MapWinGIS.Point, ByVal p2 As MapWinGIS.Point) As MapWinGIS.Point
        Dim sf As New MapWinGIS.Shapefile
        Dim pold As New MapWinGIS.Point
        sf.Open(shpPipe)
        Dim selectedshp() As Integer
        Dim Ext As New MapWinGIS.Extents
        Dim j As Integer
        Ext.SetBounds(pt.x, pt.y, 0, pt.x, pt.y, 0)
        Dim f As Boolean = sf.SelectShapes(Ext, 0.5, SelectMode.INTERSECTION, selectedshp)

        arr.Add(pt)
        arr.Add(p1)
        arr.Add(p2)
        For j = 0 To UBound(selectedshp)
            pold = sf.QuickPoint(selectedshp(j), 0)
            If CheckList(pold) = False Then
                FindID = pold
            End If
            pold = sf.QuickPoint(selectedshp(j), 1)
            If CheckList(pold) = False Then
                FindID = pold
            End If
        Next
        sf.Close()
        arr.Clear()
    End Function

    Private Function CheckList(ByVal pt As MapWinGIS.Point) As Boolean
        Dim p As New MapWinGIS.Point
        For Each p In arr
            If (pt.x = p.x And pt.y = p.y) Then
                CheckList = True
                Exit Function
            Else
                CheckList = False
            End If
        Next

    End Function

    Private Sub TComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TComboBox1.SelectedIndexChanged


        FN = TComboBox1.Items.IndexOf(TComboBox1.Text)





    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Settings1.Default.Net = TextBox1.Text
        Settings1.Default.Save()
        NetName = TextBox1.Text + "."
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        AxMap1.CursorMode = MapWinGIS.tkCursorMode.cmPan
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        AxMap1.CursorMode = MapWinGIS.tkCursorMode.cmZoomIn
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        AxMap1.CursorMode = MapWinGIS.tkCursorMode.cmZoomOut
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        AxMap1.ZoomToMaxVisibleExtents()
    End Sub

    Private Sub SetPrior(ByVal Pt As MapWinGIS.Point, ByVal sf As MapWinGIS.Shapefile)
        Dim P1, P2, P0, P3 As New MapWinGIS.Point
        Dim K As Integer
        Dim Pnt(4) As MapWinGIS.Point
        Dim S0, S1, S2, S3, DEs1, DEs2, DEs3 As Single
        Dim PL As New ArrayList
        Dim i, j As Integer
        Static sK As Integer = 0
        Dim A(4), Temp1 As Double
        Dim ID(4) As Integer
        Dim PP(4) As Integer
        Dim Temp2 As Integer
        Do
            Try

                Application.DoEvents()
                AddNumber(Pt, sf)

                K = GetNo(Pt)
                GetPointIntersection(Pt)

                If K = 1 Then


                    P1 = QueueNumber.Dequeue

                    Pt = P1
                    Man_No += 0.01

                End If

                If K = 2 Then

                    P1 = QueueNumber.Dequeue
                    P2 = QueueNumber.Dequeue
                    P0 = FindID(Pt, P1, P2)

                    If Math.Abs(Slope(P0, Pt) - Slope(Pt, P1)) < Math.Abs(Slope(P0, Pt) - Slope(Pt, P2)) Then
                        Pt = P1
                        Man_No += 0.01
                        PL.Add(P2)

                    Else
                        Pt = P2
                        Man_No += 0.01
                        PL.Add(P1)
                    End If


                End If

                If K = 3 Then
                    P1 = QueueNumber.Dequeue
                    P2 = QueueNumber.Dequeue
                    P3 = QueueNumber.Dequeue

                    P0 = Find3ID(Pt, P1, P2, P3)
                    S0 = Slope(P0, Pt)
                    S1 = Slope(Pt, P1)
                    S2 = Slope(Pt, P2)
                    S3 = Slope(Pt, P3)
                    DEs1 = Math.Abs(S0 - S1)
                    DEs2 = Math.Abs(S0 - S2)
                    DEs3 = Math.Abs(S0 - S3)

                    A(0) = DEs1
                    A(1) = DEs2
                    A(2) = DEs3
                    Pnt(0) = P1
                    Pnt(1) = P2
                    Pnt(2) = P3
                    PP(0) = 0
                    PP(1) = 1
                    PP(2) = 2

                    For i = 0 To 2
                        For j = 0 To 2
                            If A(i) < A(j) Then
                                Temp1 = A(i)
                                A(i) = A(j)
                                A(j) = Temp1

                                Temp2 = PP(i)
                                PP(i) = PP(j)
                                PP(j) = Temp2
                            End If
                        Next
                    Next



                    Pt = Pnt(PP(0))

                    Man_No += 0.01
                    PL.Add(Pnt(PP(1)))
                    PL.Add(Pnt(PP(2)))

                End If

             Catch ex As Exception
                Dim hhh As Integer = AxMap1.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
                AxMap1.AddLabel(0, "Connection Error", RGB(255, 0, 0), Pt.x, Pt.y, tkHJustification.hjRight)
                Debug.Print(Pt.x.ToString + "   " + Pt.y.ToString + " EXCEPTUON  ")
                Err = True
                Exit Sub

            End Try
        Loop Until (K = 0)
        P0 = Nothing
        P1 = Nothing
        P2 = Nothing
        Pt = Nothing
        P3 = Nothing
        Dim Pn As New MapWinGIS.Point
        For Each Pn In PL
            Man_No = 0.01
            Main_Man += 1
            SetPrior(Pn, sf)
        Next
        ProgressBar1.Value = 100

    End Sub


    Private Function Area(ByVal Fid As Integer) As Double
        Dim sf As New MapWinGIS.Shapefile
        Dim utils As New MapWinGIS.Utils()
        Dim shape As New MapWinGIS.Shape()
        Dim a As Double
        Dim i As Integer
        sf.Open(shpBlock)
        shape = sf.Shape(Fid)
        'Get the area of the polygon shape
        a = utils.Area(shape) / 10000.0   'in meter
        Area = a / 2

        sf.Close()
    End Function
     
    Private Sub CreatField()
        Dim fd As MapWinGIS.Field

        Dim sf As New MapWinGIS.Shapefile
        Try
            sf.Open(shpManhol)
            sf.StartEditingTable()
            Fd2_Indx = sf.NumFields
            fd = New MapWinGIS.Field
            fd.Name = "Man_Number"
            fd.Type = MapWinGIS.FieldType.STRING_FIELD
            fd.Width = 15
            sf.EditInsertField(fd, Fd2_Indx)
            sf.StartEditingTable()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        sf.StopEditingTable()
        sf.Close()


        'Dim fd2 As MapWinGIS.Field
        'Dim sfb As New MapWinGIS.Shapefile
        'Try
        '    sfb.Open(shpManhol)
        '    sfb.StartEditingTable()
        '    AreaFd2_Indx = sfb.NumFields
        '    fd2 = New MapWinGIS.Field
        '    fd2.Name = "Area"
        '    fd2.Type = MapWinGIS.FieldType.STRING_FIELD
        '    fd2.Width = 15
        '    sfb.EditInsertField(fd2, AreaFd2_Indx)
        '    sfb.StartEditingTable()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
        'sfb.StopEditingTable()
        'sfb.Close()

    End Sub


    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        If CheckBox1.Checked = True Then Main_Man = NumericUpDown1.Value
    End Sub

   
    '  Convert Shape file to Excel File


    Private Sub PipeConnection()
        Dim sf As New MapWinGIS.Shapefile
        Dim i As Integer
        Dim Pt As New MapWinGIS.Point
        sf.Open(shpManhol)
        Dim Man_No(sf.NumShapes + 1) As Double
        Dim x1(sf.NumShapes + 1) As Double
        Dim y1(sf.NumShapes + 1) As Double
        Dim XCoord(sf.NumShapes + 1) As Double
        Dim YCoord(sf.NumShapes + 1) As Double
        Dim GLevel(sf.NumShapes + 1) As Double
        Dim strVal As String
        Dim ln As Integer
        ReDim Man_Str(sf.NumShapes + 1)
        ReDim Man2_Str(sf.NumShapes + 1)
        ReDim Man(sf.NumShapes + 1)
        ReDim xx(sf.NumShapes + 1)
        ReDim yy(sf.NumShapes + 1)
        Dim Area(sf.NumShapes + 1) As Single
        ProgressBar1.Maximum = sf.NumShapes

        For i = 0 To sf.NumShapes - 1
            strVal = GetStr(sf.CellValue(Fd2_Indx, i))

            Man_Str(i) = strVal
            Man2_Str(i) = strVal

            ln = GetStrlen(strVal)
            If ln = 2 Then strVal = Int(strVal) + GetNewStr(strVal)
            Man(i) = Val(strVal)
            Man_No(i) = Val(strVal)
            Pt = sf.QuickPoint(i, 0)
            XCoord(i) = Pt.x
            YCoord(i) = Pt.y
            xx(i) = XCoord(i)
            yy(i) = YCoord(i)

            GLevel(i) = sf.CellValue(Elev_No, i)
            ' Area(i) = sf.CellValue(AreaFd2_Indx, i)


A10:        ProgressBar1.Value = i
            Application.DoEvents()
        Next

        Sorting(XCoord, YCoord, Man_No, GLevel, sf.NumShapes - 1, Man_Str)


        sf.Close()
    End Sub

    Private Function GetStrlen(ByVal str As String) As Integer
        Dim k As Integer
        Dim st As String
        k = str.IndexOf(".")
        st = Mid(str, k + 2, Len(str))
        GetStrlen = Len(st)
    End Function


    Private Function GetStr(ByVal str As String) As String
        Dim k As Integer

        k = str.IndexOf(".")

        GetStr = Mid(str, k + 2, Len(str))
    End Function
    

    Private Function GetNewStr(ByVal str As String) As String
        Dim k As Integer
        Dim st As String
        k = str.IndexOf(".")
        st = Mid(str, k + 2, Len(str))
        GetNewStr = "0.0" + st

    End Function

    Private Sub Sorting(ByVal XCoord() As Double, ByVal YCoord() As Double, ByVal Man_no() As Double, ByVal Glevel() As Double, ByVal Cnt As Integer, ByVal Man_Str() As String)
        Dim i, j As Integer
        Dim T1 As Double
        Dim ptX As Double
        Dim ptY As Double
        Dim pt0 As New MapWinGIS.Point
        Dim pt1 As New MapWinGIS.Point
        Dim pt2 As New MapWinGIS.Point
        Dim sf As New MapWinGIS.Shapefile
        Dim xd, yd, Distance As Double
        Dim Tstr As String
        Dim ff As Boolean
        sf.Open(shpPipe)
        Dim sXL1 As Microsoft.Office.Interop.Excel.Application
        Dim sWB1 As Microsoft.Office.Interop.Excel.Workbook
        Dim sSheet1 As Microsoft.Office.Interop.Excel.Worksheet
        Dim range1 As Object
        Dim arrayst1(19) As String
        Dim length(Cnt + 1) As Double
        sXL1 = CreateObject("Excel.Application")
        ' Get a new blank workbook.
        sWB1 = sXL1.Workbooks.Add
        sSheet1 = sWB1.ActiveSheet
        range1 = sSheet1.Range("A1", "K1")
        arrayst1(0) = "Area"
        arrayst1(1) = "Length"
        arrayst1(2) = "Man_No1"
        arrayst1(3) = "Man_No2"
        arrayst1(4) = "Ground L"
        arrayst1(5) = "x1"
        arrayst1(6) = "y1"
        arrayst1(7) = "x2"
        arrayst1(8) = "y2"
        arrayst1(9) = "Mh1"
        arrayst1(10) = "Mh2"

        ff = False
        range1.Value2 = arrayst1

        For i = 0 To Cnt
            For j = i + 1 To Cnt
                If Man_no(i) > Man_no(j) Then


                    T1 = Man_no(j)
                    Man_no(j) = Man_no(i)
                    Man_no(i) = T1

                    Tstr = Man_Str(j)
                    Man_Str(j) = Man_Str(i)
                    Man_Str(i) = Tstr



                    T1 = XCoord(j)
                    XCoord(j) = XCoord(i)
                    XCoord(i) = T1

                    T1 = YCoord(j)
                    YCoord(j) = YCoord(i)
                    YCoord(i) = T1


                    T1 = Glevel(j)
                    Glevel(j) = Glevel(i)
                    Glevel(i) = T1

                End If
            Next j
            ProgressBar1.Value = i
            Application.DoEvents()
        Next i




        Dim sv, kk As Integer
        sv = 0
        For kk = 0 To Cnt

            If Math.Abs(Int(Man_no(kk)) - Int(Man_no(kk + 1))) >= 1 Then

                For i = sv To kk
                    For j = i + 1 To kk

                        If Man_no(i) < Man_no(j) Then

                            T1 = Man_no(j)
                            Man_no(j) = Man_no(i)
                            Man_no(i) = T1

                            Tstr = Man_Str(j)
                            Man_Str(j) = Man_Str(i)
                            Man_Str(i) = Tstr

                            T1 = XCoord(j)
                            XCoord(j) = XCoord(i)
                            XCoord(i) = T1

                            T1 = YCoord(j)
                            YCoord(j) = YCoord(i)
                            YCoord(i) = T1

                            T1 = Glevel(j)
                            Glevel(j) = Glevel(i)
                            Glevel(i) = T1


                        End If


                        ProgressBar1.Value = j
                        Application.DoEvents()
                    Next j
                    ProgressBar1.Value = i
                    Application.DoEvents()
                Next i
                sv = kk + 1
            End If

            ProgressBar1.Value = kk
            Application.DoEvents()
        Next

        j = 0


        For i = 0 To Cnt
            range1 = sSheet1.Range("A" & j + 2, "K" & j + 2)


            xd = Math.Abs(XCoord(i) - XCoord(i + 1))
            yd = Math.Abs(YCoord(i) - YCoord(i + 1))
            Distance = Math.Sqrt(xd * xd + yd * yd)
            arrayst1(0) = ((Distance * 50.0) / 10000).ToString("0.000")
            arrayst1(1) = Format(Distance, "0")
            If (Man_no(i) = 0 Or Man_no(i + 1) = 0) Then
                GoTo A300
            End If

            arrayst1(2) = Man_no(i)
            arrayst1(3) = Man_no(i + 1)
            arrayst1(4) = Glevel(i)
            arrayst1(5) = XCoord(i)
            arrayst1(6) = YCoord(i)
            arrayst1(7) = XCoord(i + 1)
            arrayst1(8) = YCoord(i + 1)

            arrayst1(9) = Man_Str(i)
            arrayst1(10) = Man_Str(i + 1)
            pt1.x = XCoord(i)
            pt1.y = YCoord(i)
            pt2.x = XCoord(i + 1)
            pt2.y = YCoord(i + 1)
            ' C2Area(pt1, pt2)
            '(TotalArea / PipeNo).ToString("0.00")
            If Math.Abs(Int(Man_no(i)) - Int(Man_no(i + 1))) >= 1 Then
                ptX = XCoord(i)
                ptY = YCoord(i)
                Get_Intersection(ptX, ptY)
                arrayst1(3) = xxx
                arrayst1(7) = xx2
                arrayst1(8) = yy2

                If xxx = 0 Or xxxStr = "0.0" Then
                    '   arrayst1(10) = "0.0"

                    GoTo A300
                Else
                    arrayst1(10) = xxxStr
                End If
                xd = Math.Abs(ptX - xx2)
                yd = Math.Abs(ptY - yy2)
                Distance = Math.Sqrt(xd * xd + yd * yd)
                arrayst1(1) = Format(Distance, "0")
                arrayst1(0) = ((Distance * 50.0) / 10000).ToString("0.000")
            End If
            range1.Value2 = arrayst1
            j = j + 1

A300:
            '  ProgressBar1.Value = i
            Application.DoEvents()
        Next

        j = 0
        'For i = 0 To Cnt - 1
        '    range1 = sSheet1.Range("A" & j + 2)

        '    arrayst1(0) = ((length(i) * 50) / 10000).ToString
        '    range1.Value2 = arrayst1(0)
        '    j += 1
        'Next

        sXL1.Visible = True
        sXL1.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal

        '  Report(sWB1)

        sXL1 = Nothing
        sWB1 = Nothing
        sSheet1 = Nothing
        range1 = Nothing
    End Sub

    Private Sub Get_Intersection(ByVal ptx As Double, ByVal pty As Double)
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpPipe)
        Dim i, n1, k, Index(10000) As Integer
        Dim Pot(4) As Double
        k = 0
        ProgressBar1.Maximum = sf.NumShapes
        For i = 0 To sf.NumShapes - 1
            Pot = sf.QuickPoints(i, n1)

            If (ptx >= (Pot(0) - Tol) And ptx <= (Pot(0) + Tol)) And (pty >= (Pot(1) - Tol) And pty <= (Pot(1) + Tol)) Then
                Index(k) = PointInPipe(Pot(2), Pot(3))
                If Index(k) = -1 Then GoTo AEnd
                k = k + 1

            End If
            If (ptx >= (Pot(2) - Tol) And ptx <= (Pot(2) + Tol)) And (pty >= (Pot(3) - Tol) And pty <= (Pot(3) + Tol)) Then
                Index(k) = PointInPipe(Pot(0), Pot(1))
                If Index(k) = -1 Then GoTo AEnd
                k = k + 1

            End If
AEnd:       ProgressBar1.Value = i
            Application.DoEvents()
        Next i

        Dim ind, Ct As Integer
        ind = PointInPipe(ptx, pty)
        Ct = k - 1

        Check_Man2(Index, Ct, ind)

        sf.Close()

    End Sub
    Private Function PointInPipe(ByVal x1 As Double, ByVal y1 As Double) As Integer
        Dim i As Integer
        Dim pt(2) As Double
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpManhol)
        For i = 0 To sf.NumShapes - 1
            pt = sf.QuickPoints(i, 1)
            If x1 >= pt(0) - Tol And x1 <= pt(0) + Tol And y1 >= pt(1) - Tol And y1 <= pt(1) + Tol Then
                PointInPipe = i
                Exit For
            Else : PointInPipe = -1
            End If
        Next
        sf.Close()

    End Function



    Private Sub Check_Man2(ByVal inx() As Integer, ByVal k As Integer, ByVal ind As Integer)
        Dim i As Integer
        Dim T1, mMin, Y2 As Double
        Dim Tstr, T5 As String
        Y2 = 10000.0

        For i = 0 To k
            T1 = Man(inx(i))
            Tstr = Man2_Str(inx(i))
            If T1 < Y2 Then
                Y2 = T1
                mMin = Y2
                xx2 = xx(inx(i))
                yy2 = yy(inx(i))
                T5 = Man2_Str(inx(i))

            End If
        Next i
        F = False
        If Man(ind) > mMin Then
            xxx = mMin
            xxxStr = T5
        Else
            xxx = 0

            xxxStr = "0.0"
        End If

    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "Shapefiles: (*.shp)|*.shp"




    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Static ii As Integer = 0
        ProgressBar1.Value = ii
        ii += 1
        ProgressBar1.Maximum = 100 + ii
        Total = 100 + ii
    End Sub

    
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        AxMap1.Refresh()
        Que.Clear()
        QueueNumber.Clear()
        IntPnt.Clear()
        Count = 0
        Rpt(0) = Nothing
        Rpt(1) = Nothing
        AxMap1.RemoveAllLayers()
        TComboBox1.Items.Clear()


        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "Shapefiles: (*.shp)|*.shp"
        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            shpBlock = OpenFileDialog1.FileName
            Settings1.Default.BaseMap = shpBlock
            Settings1.Default.Save()
            'TextBox2.Text = shpBlock
        End If
        If shpBlock <> "" Then Add_BaseMap()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        AxMap1.CursorMode = tkCursorMode.cmSelection

       



    End Sub

    Private Sub AxMap1_SelectBoxFinal(ByVal sender As System.Object, ByVal e As AxMapWinGIS._DMapEvents_SelectBoxFinalEvent) Handles AxMap1.SelectBoxFinal
        Dim sf As New MapWinGIS.Shapefile
        Dim myExtents As New MapWinGIS.Extents()
        Dim selectedShapes() As Integer
        Dim i As Integer, hndl As Integer
        Dim pxMin As Double, pxMax As Double, pyMin As Double, pyMax As Double, pzMin As Double, pzMax As Double
        Dim pt As New MapWinGIS.Point

        If AxMap1.CursorMode = MapWinGIS.tkCursorMode.cmSelection Then

            hndl = sf.Open(shpManhol)
            'Convert the boundaries of the selection box from pixel units to projected map coordinates
            AxMap1.PixelToProj(e.left, e.bottom, pxMin, pyMin)
            AxMap1.PixelToProj(e.right, e.top, pxMax, pyMax)
            'Set the extents object to be used to find shapes that have been selected in the shapefile
            myExtents.SetBounds(pxMin, pyMin, 0, pxMax, pyMax, 0)
            'Check if there are any shapes with in the shapefile that intersect with the selection box
            If sf.SelectShapes(myExtents, 0, MapWinGIS.SelectMode.INTERSECTION, selectedShapes) Then
                If UBound(selectedShapes) > 0 Then
                    MessageBox.Show("Slelect 1 Feature")
                    GoTo A100
                End If

                Dim kk As Integer = AxMap1.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
                ' sf.StartEditingTable()
                pt = sf.QuickPoint(selectedShapes(0), 0)
                AxMap1.DrawCircleEx(LyrHnd2, pt.x, pt.y, 4, RGB(255, 0, 0), True)
                '  sf.EditCellValue(Fd2_Indx, selectedShapes(0), "111")
                '  sf.StopEditingTable()
                sf.Close()
            End If

        End If
A100:
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Elev_No = ToolStripComboBox1.Items.IndexOf(ToolStripComboBox1.Text)
        Button1.Enabled = True

    End Sub
End Class


