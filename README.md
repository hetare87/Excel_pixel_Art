# Excel_pixel_Art
Excelでドット絵が描けます。

```VBA

Option Explicit

'========================================================
' GDI+ で画像ファイルからピクセル取得してセルに描画
' jpg/jpeg/png 対応・2回目以降はクリアして再描画
'使い方
'Dotify_Image_To_Cellsを実行してください。
'ダイアログボックスから任意の画像を選択して実行してください。
'環境によって動かない、動作が遅いなどあればすいません。
'========================================================

#If VBA7 Then
    ' 64bit / 32bit (VBA7) 共通
    Private Type GdiplusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As LongPtr
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type

    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As LongPtr) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)

    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
    Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
    Private Declare PtrSafe Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal x As Long, ByVal y As Long, ByRef argb As Long) As Long
#Else
    ' 古いExcel (VBA6以前)
    Private Type GdiplusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type

    Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

    Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, ByRef bitmap As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
    Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, ByRef width As Long) As Long
    Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, ByRef height As Long) As Long
    Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, ByRef argb As Long) As Long
#End If

'=========================
' メイン処理
'=========================
Public Sub Dotify_Image_To_Cells()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    '==== 調整用定数（PCが重い場合は MAX_CELLS を減らしてください）====
    Const START_ROW As Long = 1
    Const START_COL As Long = 1

    ' 出力解像度の設定
    ' OUTPUT_WIDTH: 出力の横ピクセル数（MAX_CELLS_Wより大きく設定可能）
    ' OUTPUT_HEIGHT: 出力の縦ピクセル数（OUTPUT_WIDTHに応じて自動計算されるが、直接指定も可）
    Const OUTPUT_WIDTH As Long = 200  ' 好みに応じて増減してください（例: 200?500 等）
    Const OUTPUT_HEIGHT As Long = 0   ' 0 の場合は自動計算（縦は比率から算出）

    ' 画像上限の設定（実際の出力セル数の最大値）
    Const MAX_CELLS_W As Long = 400   ' 横幅のセル数上限
    Const MAX_CELLS_H As Long = 400   ' 縦幅のセル数上限
    ' セルの見た目サイズ
    ' 正方形に近づけるための係数。環境に合わせて微調整してください。
    ' 例:
    Const CELL_W_SIZE As Double = 0.9  ' 列幅（調整で正方形に見せる）
    Const CELL_H_SIZE As Double = 14   ' 行の高さ
    Const PROGRESS_EVERY_ROWS As Long = 5

    '==== ファイル選択 ====
    Dim filePath As Variant ' Variantにしてキャンセル判定しやすくする
    filePath = Application.GetOpenFilename( _
        "Image Files (*.jpg;*.jpeg;*.png;*.bmp),*.jpg;*.jpeg;*.png;*.bmp", _
        , "画像を選択してください")
    
    If filePath = False Then Exit Sub

    Dim oldStatus As Variant
    oldStatus = Application.StatusBar

    '==== 高速化設定 ====
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "初期化中..."

    On Error GoTo EH

    '==== シートのクリア ====
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone
    ' 全体の列幅・行高さを一旦リセット（あるいはデフォルトに）
    ws.Cells.ColumnWidth = 8.38
    ws.Cells.RowHeight = 13.5

    '==== GDI+ 起動 ====
#If VBA7 Then
    Dim token As LongPtr
    Dim bmp As LongPtr
#Else
    Dim token As Long
    Dim bmp As Long
#End If

    token = GDIPlus_Startup()

    '==== 画像読み込み ====
#If VBA7 Then
    If GdipCreateBitmapFromFile(StrPtr(filePath), bmp) <> 0 Or bmp = 0 Then
        Err.Raise vbObjectError + 1000, , "画像を読み込めません（GDI+）"
    End If
#Else
    If GdipCreateBitmapFromFile(StrPtr(filePath), bmp) <> 0 Or bmp = 0 Then
        Err.Raise vbObjectError + 1000, , "画像を読み込めません（GDI+）"
    End If
#End If

    Dim imgW As Long, imgH As Long
    Call GdipGetImageWidth(bmp, imgW)
    Call GdipGetImageHeight(bmp, imgH)

    If imgW = 0 Or imgH = 0 Then Err.Raise vbObjectError + 1001, , "画像サイズ取得失敗"

    '==== 出力サイズ計算（アスペクト比維持）====

    Dim outW As Long, outH As Long

    ' OUTPUT_WIDTH が設定されている場合、それを使用
    outW = OUTPUT_WIDTH

    If OUTPUT_HEIGHT > 0 Then
        outH = OUTPUT_HEIGHT
    Else
        ' 幅基準で計算
        outH = CLng((CDbl(imgH) / CDbl(imgW)) * outW)
    End If

    ' 高さが上限を超えたら高さ基準で再計算
    If outW > MAX_CELLS_W Then
        outW = MAX_CELLS_W
        If OUTPUT_HEIGHT = 0 Then
            outH = CLng((CDbl(imgH) / CDbl(imgW)) * outW)
        End If
    End If
    If outH > MAX_CELLS_H Then
        outH = MAX_CELLS_H
        outW = CLng((CDbl(imgW) / CDbl(imgH)) * outH)
    End If

    If outW < 1 Then outW = 1
    If outH < 1 Then outH = 1

    Dim stepX As Double, stepY As Double
    stepX = CDbl(imgW) / CDbl(outW)
    stepY = CDbl(imgH) / CDbl(outH)

    '==== セルサイズ調整（描画範囲のみ）====
    ' 正方形に見せるためのセルサイズ設定
    Dim r As Long, c As Long
    With ws.Range(ws.Cells(START_ROW, START_COL), ws.Cells(START_ROW + outH - 1, START_COL + outW - 1))
        .ColumnWidth = CELL_W_SIZE
        .RowHeight = CELL_H_SIZE
    End With

    '==== ピクセル取得＆描画 ====
    Dim rr As Long, cc As Long
    Dim srcX As Long, srcY As Long
    Dim argb As Long, excelColor As Long

    For rr = 0 To outH - 1
        ' 進捗表示
        If (rr Mod PROGRESS_EVERY_ROWS) = 0 Then
            Application.StatusBar = "描画中... " & Format((rr + 1) / outH, "0%")
            DoEvents
        End If

        srcY = Fix(rr * stepY)
        If srcY >= imgH Then srcY = imgH - 1

        For cc = 0 To outW - 1
            srcX = Fix(cc * stepX)
            If srcX >= imgW Then srcX = imgW - 1

            ' ピクセル色取得
            Call GdipBitmapGetPixel(bmp, srcX, srcY, argb)
            
            ' ARGBからExcel用RGBへ変換
            excelColor = ARGB_To_ExcelRGB(argb)

            ' セルに着色
            ws.Cells(START_ROW + rr, START_COL + cc).Interior.Color = excelColor
        Next
    Next
    
    ' ズームを引いて全体を見やすくする
    ActiveWindow.Zoom = 50

    '==== 後始末 ====
    Call GdipDisposeImage(bmp)
    Call GdiplusShutdown(token)

    Application.StatusBar = oldStatus
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "完了しました！" & vbCrLf & "サイズ: " & outW & " x " & outH, vbInformation
    Exit Sub

EH:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    On Error Resume Next
    If bmp <> 0 Then Call GdipDisposeImage(bmp)
    If token <> 0 Then Call GdiplusShutdown(token)
    
    Application.StatusBar = oldStatus
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'=========================
' GDI+ 起動関数
'=========================
#If VBA7 Then
Private Function GDIPlus_Startup() As LongPtr
#Else
Private Function GDIPlus_Startup() As Long
#End If
    Dim si As GdiplusStartupInput
    si.GdiplusVersion = 1
    
#If VBA7 Then
    Dim token As LongPtr
#Else
    Dim token As Long
#End If
    
    ' inputbuf, outputbuf の引数定義を修正済み
    If GdiplusStartup(token, si, 0) <> 0 Then
        Err.Raise vbObjectError + 2000, , "GDI+ Startup Failed"
    End If
    GDIPlus_Startup = token
End Function

'=========================
' ARGB(0xAARRGGBB) → Excel RGB
'=========================
Private Function ARGB_To_ExcelRGB(ByVal argb As Long) As Long
    ' GDI+は ARGB (Alpha, Red, Green, Blue)
    ' Excelは BGR (Blue, Green, Red) ※リトルエンディアンのRGB関数結果
    
    Dim r As Long, g As Long, b As Long
    
    ' Long型は符号付きなので、&HFF...でのマスク処理に注意
    ' Blue
    b = (argb And &HFF&)
    
    ' Green
    g = (argb And &HFF00&) \ &H100&
    
    ' Red
    r = (argb And &HFF0000) \ &H10000
    
    ' 負の数対策（VBAのLong型の仕様回避）
    If r < 0 Then r = r + 256
    If g < 0 Then g = g + 256
    If b < 0 Then b = b + 256
    
    ARGB_To_ExcelRGB = RGB(r, g, b)
End Function
```

