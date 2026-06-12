Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Điều phối thứ tự nhập liệu cho toàn bộ mẫu Ibaraki Beta và gọi các helper ghi Excel tương ứng.
''' </summary>
Friend Module Service
    ''' <summary>
    ''' Chạy toàn bộ workflow nhập liệu trọng lượng/vật tư Ibaraki Beta trên workbook đang mở.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub WtIbarakiBeta(xlApp As Application)
        ' Phí vận chuyển.
        Fare(xlApp, HdrYNQ(vbTab & vbTab & "運賃 (2トン車): "))
        ' Unit GL-150.
        HdrWrng(vbTab & vbTab & "外周/内周GL-150" & vbCrLf)
        Unit150(xlApp)
        ' Unit GL-300.
        Unit300(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-300: "))
        ' Unit GL-300/+30.
        Unit300Cut(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-300/+30: "))
        ' Unit GL-400.
        Unit400(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-400: "))
        ' Unit GL-400/+30.
        Unit400Cut(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-400/+30: "))
        ' Cửa vào/cửa sau.
        EntrBackDoor(xlApp, HdrYNQ(vbTab & vbTab & "玄関・勝手口: "))
        ' Garage GL-300.
        Unit300Gar(xlApp, HdrYNQ(vbTab & vbTab & "ガレージ外周GL-300: "))
        ' Slab unit.
        HdrWrng(vbTab & vbTab & "スラブユニット" & vbCrLf)
        SlabUnit(xlApp)
        ' Bình nước nóng điện.
        ElecWtrHtr(xlApp, HdrDInp(vbTab & vbTab & "電気温水器: "))
        ' Joint góc.
        HdrWrng(vbTab & vbTab & "コーナー" & vbCrLf)
        JtCor(xlApp)
        ' Joint thẳng.
        HdrWrng(vbTab & vbTab & "ストレート" & vbCrLf)
        JtStr(xlApp)
        ' Cap tire.
        CapTire(xlApp, HdrYNQ(vbTab & vbTab & "キャップタイヤ (320): "))
        ' Đầu biên.
        PubDVal(xlApp, "BA180", HdrDInp(vbTab & vbTab & "端部(700×350): "))
        ' Long corner.
        LongCor(xlApp, HdrYNQ(vbTab & vbTab & "ロングコーナー (D16): "))
        ' Crank.
        Crank(xlApp, HdrYNQ(vbTab & vbTab & "クランク: "))
        ' Island D16.
        Island(xlApp, HdrYNQ(vbTab & vbTab & "島 (D16): "))
        ' Thép thẳng D16.
        Straight(xlApp, HdrYNQ(vbTab & vbTab & "ストレート (D16): "))
        ' Corner 3D.
        Corner3d(xlApp, HdrYNQ(vbTab & vbTab & "コーナー3 (D16): "))
        ' Crank 3D.
        Crank3d(xlApp, HdrYNQ(vbTab & vbTab & "クランク3 (D16): "))
        ' U-type 3D.
        UType3d(xlApp, HdrYNQ(vbTab & vbTab & "コ型3 (D16): "))
        ' M-type.
        PubDModVal(xlApp, "195", "350×460×460×350", 2.7, HdrDInp(vbTab & vbTab & "Ｍ型 (D16[350×460×460×350]): "))
        ' Hook D10.
        HdrWrng(vbTab & vbTab & "フック (D10)" & vbCrLf)
        Hook(xlApp)
        ' Gia cường chính.
        MainReinf(xlApp, HdrYNQ(vbTab & vbTab & "主筋補強 (D10): "))
        ' Slab uốn.
        SlabBndg(xlApp, HdrYNQ(vbTab & vbTab & "スラブ曲 (D13): "))
        ' Slab thẳng.
        HdrWrng(vbTab & vbTab & "スラブ直 (D13)" & vbCrLf)
        SlabStr(xlApp)
        ' Slab gia cường uốn.
        SlabReinfBndg(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強曲 (D10): "))
        ' Slab gia cường thẳng.
        SlabReinfStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強直 (D10): "))
        ' Sleeve.
        Sleeve(xlApp, HdrDInp(vbTab & vbTab & "スリーブ: "))
        ' Danh sách phụ tư.
        HdrWrng(vbTab & vbTab & "副資材リスト" & vbCrLf)
        Parts(xlApp)
    End Sub
End Module
