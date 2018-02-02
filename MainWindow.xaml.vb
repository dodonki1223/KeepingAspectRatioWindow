Imports System.Windows.Interop
Imports System.Runtime.InteropServices

Class MainWindow

#Region "定数"

    ''' <summary>WM_SIZING Message定数群</summary>
    ''' <remarks></remarks>
    Private Class _cWM_SIZING

        ''' <summary>Message番号</summary>
        ''' <remarks>
        '''   サイズが変更された後にウィンドウに送信されます。
        '''   ウィンドウは、このメッセージをWindowProc関数を通じて受信します。
        ''' </remarks>
        Public Const Message As Integer = &H214

        ''' <summary>サイズ変更後の位置を取得するパラメータ</summary>
        ''' <remarks></remarks>
        Public Enum wParam

            ''' <summary>左端</summary>
            WMSZ_LEFT = 1

            ''' <summary>右端</summary>
            WMSZ_RIGHT = 2

            ''' <summary>上端</summary>
            WMSZ_TOP = 3

            ''' <summary>左上隅</summary>
            WMSZ_TOPLEFT = 4

            ''' <summary>右上隅</summary>
            WMSZ_TOPRIGHT = 5

            ''' <summary>下端</summary>
            WMSZ_BOTTOM = 6

            ''' <summary>左下隅</summary>
            WMSZ_BOTTOMLEFT = 7

            ''' <summary>右下隅</summary>
            WMSZ_BOTTOMRIGHT = 8

        End Enum

    End Class

    ''' <summary>ウィンドウの幅・高さの修正率</summary>
    ''' <remarks></remarks>
    Private Const _cFixRate As Double = 525 / 350

#End Region

#Region "構造体"

    ''' <summary>RECT構造体</summary>
    ''' <remarks>
    '''   StructLayoutについて
    '''     名前空間: System.Runtime.InteropServices
    '''     アセンブリ: mscorlib
    '''     継承・インターフェイス: Attribute, _Attribute
    '''     StructLayout属性は、メモリ上でのフィールド(メンバ変数)の配置方法を指定するための属性です
    '''   LayoutKind.Sequentialについて
    '''     ランタイムによる自動的な並べ替えを行わず、コード上で記述されている順序のままフィールドを配置する
    '''   RECT構造体
    '''     四角形の左上隅および右下隅の座標を定義します
    ''' </remarks>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT

        ''' <summary>left</summary>
        ''' <remarks>四角形の左上隅のＸ座標を指定します</remarks>
        Public left As Integer

        ''' <summary>top</summary>
        ''' <remarks>四角形の左上隅のＹ座標を指定します</remarks>
        Public top As Integer

        ''' <summary>right</summary>
        ''' <remarks>四角形の右下隅のＸ座標を指定します</remarks>
        Public right As Integer

        ''' <summary>bottom</summary>
        ''' <remarks>四角形の右下隅のＹ座標を指定します</remarks>
        Public bottom As Integer

    End Structure

#End Region

#Region "イベント"

    ''' <summary>SourceInitializedイベントを発生</summary>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>ウインドウの初期化中に呼び出されます</remarks>
    Protected Overrides Sub OnSourceInitialized(ByVal e As EventArgs)

        '基底クラスのSourceInitializedイベントを発生させる
        MyBase.OnSourceInitialized(e)

        'WPFコンテンツを格納するWin32のウィンドウを取得する
        Dim mHwndSource As HwndSource = CType(HwndSource.FromVisual(Me), HwndSource)

        'ウィンドウメッセージを受信するイベントハンドラーを追加
        mHwndSource.AddHook(AddressOf WndHookProc)

    End Sub

    ''' <summary>ウィンドウプロシージャをフック</summary>
    ''' <param name="hwnd">ウィンドウのハンドル</param>
    ''' <param name="msg">メッセージの識別子</param>
    ''' <param name="wParam">メッセージの最初のパラメータ</param>
    ''' <param name="lParam">メッセージの２番目のパラメータ</param>
    ''' <param name="handled">ハンドルフラグ</param>
    ''' <returns>0 に初期化されたポインターまたはハンドル</returns>
    ''' <remarks>
    '''   ウインドウプロシージャ
    '''     メッセージを処理する専用のルーチン
    '''   Hook（フック）
    '''     独自の処理を割り込ませるための仕組み
    ''' </remarks>
    Private Function WndHookProc(ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByRef handled As Boolean) As IntPtr

        'メッセージがウインドウの移動中の時
        If msg = _cWM_SIZING.Message Then

            'アンマネージメモリのRECT構造体をマネージオブジェクト（RECT構造体）にデータをマーシャリングする
            '※ウィンドウプロシージャに渡ってきた「lParam」を.NET側で使えるようにデータを変換する。Marshalingは「整列」という意味の英単語
            Dim mRect As RECT = Marshal.PtrToStructure(lParam, GetType(RECT))

            'ウィンドウの幅と高さを求める
            Dim mWindowWidth As Integer = mRect.right - mRect.left
            Dim mWindowHeight As Integer = mRect.bottom - mRect.top

            'ウィンドウの幅と高さの増減値を取得  
            ' ウィンドウ幅  の増減値：「(ウィンドウ高さ * 修正率) - ウィンドウ幅  」
            ' ウィンドウ高さの増減値：「(ウィンドウ幅   * 修正率) - ウィンドウ高さ」
            Dim mChangeWidth As Integer = Math.Round((mWindowHeight * _cFixRate)) - mWindowWidth
            Dim mChangeHeight As Integer = Math.Round((mWindowWidth / _cFixRate)) - mWindowHeight

            Select Case wParam.ToInt32()

                Case _cWM_SIZING.wParam.WMSZ_LEFT, _cWM_SIZING.wParam.WMSZ_RIGHT

                    '「左端」と「右端」の時は、ウインドウ幅の増減値を右下隅のＹ座標に設定
                    mRect.bottom = mRect.bottom + mChangeHeight

                Case _cWM_SIZING.wParam.WMSZ_TOP, _cWM_SIZING.wParam.WMSZ_BOTTOM

                    '「上端」と「下端」の時は、ウインドウ高さの増減値を右下隅のＸ座標に設定
                    mRect.right = mRect.right + mChangeWidth

                Case _cWM_SIZING.wParam.WMSZ_TOPLEFT

                    'ウィンドウ幅の増減値が０より大きい時
                    If (mChangeWidth > 0) Then

                        'ウィンドウの左位置を再設定「ウィンドウの左位置 - ウィンドウ幅の増減値」
                        mRect.left = mRect.left - mChangeWidth

                    Else

                        'ウィンドウの上位置を再設定「ウィンドウの上位置 - ウィンドウ高さの増減値」
                        mRect.top = mRect.top - mChangeHeight

                    End If

                Case _cWM_SIZING.wParam.WMSZ_TOPRIGHT

                    'ウィンドウ幅の増減値が０より大きい時
                    If (mChangeWidth > 0) Then

                        'ウィンドウの右位置を再設定「ウィンドウの右位置 + ウィンドウ幅の増減値」
                        mRect.right = mRect.right + mChangeWidth

                    Else

                        'ウィンドウの上位置を再設定「ウィンドウの上位置 - ウィンドウ高さの増減値」
                        mRect.top = mRect.top - mChangeHeight

                    End If

                Case _cWM_SIZING.wParam.WMSZ_BOTTOMLEFT

                    'ウィンドウ幅の増減値が０より大きい時
                    If (mChangeWidth > 0) Then

                        'ウィンドウの左位置を再設定「ウィンドウの左位置 - ウィンドウ幅の増減値」
                        mRect.left = mRect.left - mChangeWidth

                    Else

                        'ウィンドウの下位置を再設定「ウィンドウの下位置 + ウィンドウ高さの増減値」
                        mRect.bottom = mRect.bottom + mChangeHeight

                    End If

                Case _cWM_SIZING.wParam.WMSZ_BOTTOMRIGHT

                    'ウィンドウ幅の増減値が０より大きい時
                    If (mChangeWidth > 0) Then

                        'ウィンドウの右位置を再設定「ウィンドウの右位置 + ウィンドウ幅の増減値」
                        mRect.right = mRect.right + mChangeWidth

                    Else

                        'ウィンドウの下位置を再設定「ウィンドウの下位置 + ウィンドウ高さの増減値」
                        mRect.bottom = mRect.bottom + mChangeHeight

                    End If

            End Select

            'マネージオブジェクト（RECT構造体）をアンマネージメモリブロックにデータをマーシャリングする
            '※この処理で変更したRECT構造体の値を
            Marshal.StructureToPtr(mRect, lParam, False)

        End If

        Return IntPtr.Zero

    End Function

#End Region

End Class
