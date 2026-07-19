Attribute VB_Name = "mQRCode"
'==============================================================================
' mQRCode  ---  純VBA QRコード生成エンジン（外部ライブラリ・API 不要）
'------------------------------------------------------------------------------
'  ・バイトモード（UTF-8）で任意の文字列を QR コードに符号化します。
'    英数字・記号はもちろん、日本語（漢字・かな）もそのまま扱えます。
'  ・型番(バージョン) 1〜40、誤り訂正レベル L / M / Q / H に対応。
'  ・Reed-Solomon 誤り訂正、ブロック分割・インターリーブ、マスク自動選択を
'    ISO/IEC 18004 に沿って実装しています。
'
'  公開関数はただ 1 つ:
'      QR_Generate(text, [eccLevel]) -> Variant(2次元Boolean配列)
'          戻り値 m は m(x, y) 形式。True=黒モジュール、False=白。
'          x=列, y=行, 範囲は 0 〜 size-1 (size = バージョン*4+17)。
'
'  ※ 実装は Nayuki 方式のアルゴリズムを踏襲し、Python で segno / zbar と
'    突き合わせて全型番・全レベルでデコード一致を確認済みです。
'==============================================================================
Option Explicit
Option Base 0

Private mTablesReady As Boolean
Private ECCB() As Long          ' ECCB(ecl, ver) : 1ブロックあたりの誤り訂正符号語数
Private NEB()  As Long          ' NEB(ecl, ver)  : 誤り訂正ブロック数

'============================ 公開 API =========================================

' text     : QR 化する文字列
' eccLevel : "L" / "M" / "Q" / "H"（既定 "M"）。誤り訂正が強いほど汚れに強いが容量減。
Public Function QR_Generate(ByVal text As String, Optional ByVal eccLevel As String = "M") As Variant
    EnsureTables
    Dim ecl As Long: ecl = ParseEcl(eccLevel)

    Dim bytes() As Long, n As Long
    n = StringToUTF8(text, bytes)

    ' --- 収まる最小バージョンを選ぶ ---
    Dim ver As Long: ver = 0
    Dim v As Long
    For v = 1 To 40
        Dim capBits As Long: capBits = GetNumDataCodewords(v, ecl) * 8
        Dim usedBits As Long: usedBits = 4 + ByteModeCharCountBits(v) + 8 * n
        If usedBits <= capBits Then ver = v: Exit For
    Next v
    If ver = 0 Then
        Err.Raise vbObjectError + 513, "QR_Generate", _
            "データが長すぎて QR コードに収まりません（UTF-8 " & n & " バイト）。" & _
            "内容を短くするか、誤り訂正レベルを下げてください。"
    End If

    Dim dataCw() As Long: dataCw = BuildDataCodewords(bytes, n, ver, ecl)
    Dim allCw() As Long:  allCw = AddEccInterleave(dataCw, ver, ecl)
    QR_Generate = BuildMatrix(allCw, ver, ecl)
End Function

'============================ 文字列 → UTF-8 ==================================

Private Function StringToUTF8(ByVal s As String, ByRef outBytes() As Long) As Long
    Dim tmp() As Long
    ReDim tmp(0 To Len(s) * 4 + 4)
    Dim cnt As Long: cnt = 0
    Dim i As Long, L As Long: L = Len(s)
    i = 1
    Do While i <= L
        Dim code As Long
        code = AscW(Mid$(s, i, 1)) And &HFFFF&           ' 0..65535 に正規化
        ' サロゲートペア（絵文字・拡張漢字など）を結合
        If code >= &HD800& And code <= &HDBFF& And i < L Then
            Dim lo As Long: lo = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00& And lo <= &HDFFF& Then
                code = &H10000 + ((code - &HD800&) * &H400&) + (lo - &HDC00&)
                i = i + 1
            End If
        End If
        i = i + 1
        If code < &H80& Then
            tmp(cnt) = code: cnt = cnt + 1
        ElseIf code < &H800& Then
            tmp(cnt) = &HC0& Or (code \ &H40&): cnt = cnt + 1
            tmp(cnt) = &H80& Or (code And &H3F&): cnt = cnt + 1
        ElseIf code < &H10000 Then
            tmp(cnt) = &HE0& Or (code \ &H1000&): cnt = cnt + 1
            tmp(cnt) = &H80& Or ((code \ &H40&) And &H3F&): cnt = cnt + 1
            tmp(cnt) = &H80& Or (code And &H3F&): cnt = cnt + 1
        Else
            tmp(cnt) = &HF0& Or (code \ &H40000): cnt = cnt + 1
            tmp(cnt) = &H80& Or ((code \ &H1000&) And &H3F&): cnt = cnt + 1
            tmp(cnt) = &H80& Or ((code \ &H40&) And &H3F&): cnt = cnt + 1
            tmp(cnt) = &H80& Or (code And &H3F&): cnt = cnt + 1
        End If
    Loop
    If cnt = 0 Then
        ReDim outBytes(0 To 0)
    Else
        ReDim outBytes(0 To cnt - 1)
        For i = 0 To cnt - 1: outBytes(i) = tmp(i): Next i
    End If
    StringToUTF8 = cnt
End Function

'============================ データ符号語生成 ================================

Private Function BuildDataCodewords(bytes() As Long, ByVal n As Long, ByVal ver As Long, ByVal ecl As Long) As Long()
    Dim ndc As Long:  ndc = GetNumDataCodewords(ver, ecl)
    Dim cap As Long:  cap = ndc * 8
    Dim bits() As Long: ReDim bits(0 To cap - 1)
    Dim bl As Long: bl = 0

    AppendBits bits, bl, 4, 4                                 ' モード指示子: バイト = 0100
    AppendBits bits, bl, n, ByteModeCharCountBits(ver)        ' 文字数指示子
    Dim i As Long
    For i = 0 To n - 1
        AppendBits bits, bl, bytes(i), 8
    Next i

    Dim term As Long: term = cap - bl: If term > 4 Then term = 4
    AppendBits bits, bl, 0, term                              ' 終端パターン
    AppendBits bits, bl, 0, (8 - (bl Mod 8)) Mod 8            ' バイト境界まで 0 埋め

    Dim padByte As Long: padByte = &HEC&
    Do While bl < cap                                         ' 埋め草符号語 EC/11 交互
        AppendBits bits, bl, padByte, 8
        If padByte = &HEC& Then padByte = &H11& Else padByte = &HEC&
    Loop

    Dim cw() As Long: ReDim cw(0 To ndc - 1)
    Dim b As Long
    For i = 0 To ndc - 1
        Dim vv As Long: vv = 0
        For b = 0 To 7
            vv = vv * 2 + bits(i * 8 + b)
        Next b
        cw(i) = vv
    Next i
    BuildDataCodewords = cw
End Function

Private Sub AppendBits(bits() As Long, ByRef bl As Long, ByVal val As Long, ByVal numBits As Long)
    Dim i As Long
    For i = numBits - 1 To 0 Step -1
        bits(bl) = (val \ CLng(2 ^ i)) And 1&
        bl = bl + 1
    Next i
End Sub

'============================ Reed-Solomon (GF(256)) ==========================

Private Function GFMul(ByVal x As Long, ByVal y As Long) As Long
    Dim z As Long, i As Long
    z = 0
    For i = 7 To 0 Step -1
        z = z * 2
        If (z And &H100&) <> 0 Then z = z Xor &H11D&
        If ((y \ CLng(2 ^ i)) And 1&) <> 0 Then z = z Xor x
    Next i
    GFMul = z And &HFF&
End Function

Private Function RSDivisor(ByVal degree As Long) As Long()
    Dim result() As Long: ReDim result(0 To degree - 1)
    result(degree - 1) = 1
    Dim root As Long: root = 1
    Dim i As Long, j As Long
    For i = 0 To degree - 1
        For j = 0 To degree - 1
            result(j) = GFMul(result(j), root)
            If j + 1 < degree Then result(j) = result(j) Xor result(j + 1)
        Next j
        root = GFMul(root, 2)
    Next i
    RSDivisor = result
End Function

Private Function RSRemainder(data() As Long, ByVal dataLen As Long, divisor() As Long, ByVal degree As Long) As Long()
    Dim result() As Long: ReDim result(0 To degree - 1)   ' 0 で初期化
    Dim i As Long, k As Long
    For i = 0 To dataLen - 1
        Dim factor As Long: factor = (data(i) Xor result(0)) And &HFF&
        For k = 0 To degree - 2
            result(k) = result(k + 1)
        Next k
        result(degree - 1) = 0
        For k = 0 To degree - 1
            result(k) = result(k) Xor GFMul(divisor(k), factor)
        Next k
    Next i
    RSRemainder = result
End Function

Private Function AddEccInterleave(dataCw() As Long, ByVal ver As Long, ByVal ecl As Long) As Long()
    Dim nb As Long:   nb = NEB(ecl, ver)
    Dim be As Long:   be = ECCB(ecl, ver)
    Dim raw As Long:  raw = GetNumRawDataModules(ver) \ 8
    Dim nShort As Long: nShort = nb - (raw Mod nb)
    Dim sLen As Long:   sLen = raw \ nb
    Dim divisor() As Long: divisor = RSDivisor(be)

    Dim blocks() As Long: ReDim blocks(0 To nb - 1, 0 To sLen)   ' 各ブロック長 = sLen+1
    Dim i As Long, j As Long, k As Long
    k = 0
    For j = 0 To nb - 1
        Dim dl As Long: dl = sLen - be
        If j >= nShort Then dl = dl + 1
        Dim datArr() As Long: ReDim datArr(0 To dl - 1)
        For i = 0 To dl - 1
            datArr(i) = dataCw(k + i)
            blocks(j, i) = datArr(i)
        Next i
        k = k + dl
        Dim ecc() As Long: ecc = RSRemainder(datArr, dl, divisor, be)
        Dim es As Long: es = (sLen + 1) - be
        For i = 0 To be - 1
            blocks(j, es + i) = ecc(i)
        Next i
    Next j

    Dim result() As Long: ReDim result(0 To raw - 1)
    k = 0
    For i = 0 To sLen
        For j = 0 To nb - 1
            If Not (i = (sLen - be) And j < nShort) Then    ' 短ブロックの埋め位置を飛ばす
                result(k) = blocks(j, i)
                k = k + 1
            End If
        Next j
    Next i
    AddEccInterleave = result
End Function

'============================ 行列（モジュール）構築 =========================

Private Function BuildMatrix(allCw() As Long, ByVal ver As Long, ByVal ecl As Long) As Variant
    Dim size As Long: size = ver * 4 + 17
    Dim mods() As Boolean, isFunc() As Boolean
    ReDim mods(0 To size - 1, 0 To size - 1)
    ReDim isFunc(0 To size - 1, 0 To size - 1)

    DrawFunctionPatterns ver, ecl, mods, isFunc, size
    DrawCodewords allCw, mods, isFunc, size

    ' --- マスク自動選択（ペナルティ最小のものを採用） ---
    Dim msk As Long, bestMask As Long, minP As Long, p As Long
    minP = &H7FFFFFFF: bestMask = 0
    For msk = 0 To 7
        DrawFormatBits ecl, msk, mods, isFunc, size
        ApplyMask msk, mods, isFunc, size
        p = PenaltyScore(mods, size)
        If p < minP Then minP = p: bestMask = msk
        ApplyMask msk, mods, isFunc, size          ' XOR をもう一度掛けて元に戻す
    Next msk
    DrawFormatBits ecl, bestMask, mods, isFunc, size
    ApplyMask bestMask, mods, isFunc, size

    BuildMatrix = mods
End Function

Private Sub DrawFunctionPatterns(ByVal ver As Long, ByVal ecl As Long, mods() As Boolean, isFunc() As Boolean, ByVal size As Long)
    Dim i As Long
    For i = 0 To size - 1                                     ' タイミングパターン
        SetFunc mods, isFunc, 6, i, (i Mod 2 = 0)
        SetFunc mods, isFunc, i, 6, (i Mod 2 = 0)
    Next i
    DrawFinder mods, isFunc, size, 3, 3                       ' 位置検出（ファインダ）
    DrawFinder mods, isFunc, size, size - 4, 3
    DrawFinder mods, isFunc, size, 3, size - 4

    Dim pos() As Long, np As Long
    np = AlignPositions(ver, pos)                             ' 位置合わせ（アライメント）
    Dim a As Long, b As Long
    For a = 0 To np - 1
        For b = 0 To np - 1
            If Not ((a = 0 And b = 0) Or (a = 0 And b = np - 1) Or (a = np - 1 And b = 0)) Then
                DrawAlign mods, isFunc, pos(a), pos(b)
            End If
        Next b
    Next a

    DrawFormatBits ecl, 0, mods, isFunc, size                ' 仮のフォーマット情報
    DrawVersion ver, mods, isFunc, size                      ' 型番情報（ver>=7）
End Sub

Private Sub DrawFinder(mods() As Boolean, isFunc() As Boolean, ByVal size As Long, ByVal cx As Long, ByVal cy As Long)
    Dim dx As Long, dy As Long, dist As Long, xx As Long, yy As Long
    For dy = -4 To 4
        For dx = -4 To 4
            dist = MaxL(Abs(dx), Abs(dy))
            xx = cx + dx: yy = cy + dy
            If xx >= 0 And xx < size And yy >= 0 And yy < size Then
                SetFunc mods, isFunc, xx, yy, (dist <> 2 And dist <> 4)
            End If
        Next dx
    Next dy
End Sub

Private Sub DrawAlign(mods() As Boolean, isFunc() As Boolean, ByVal cx As Long, ByVal cy As Long)
    Dim dx As Long, dy As Long
    For dy = -2 To 2
        For dx = -2 To 2
            SetFunc mods, isFunc, cx + dx, cy + dy, (MaxL(Abs(dx), Abs(dy)) <> 1)
        Next dx
    Next dy
End Sub

Private Function AlignPositions(ByVal ver As Long, ByRef pos() As Long) As Long
    If ver = 1 Then AlignPositions = 0: Exit Function
    Dim na As Long: na = ver \ 7 + 2
    Dim st As Long
    If ver = 32 Then
        st = 26
    Else
        st = ((ver * 4 + na * 2 + 1) \ (na * 2 - 2)) * 2
    End If
    ReDim pos(0 To na - 1)
    pos(0) = 6
    Dim p As Long: p = ver * 4 + 10
    Dim i As Long
    For i = na - 1 To 1 Step -1
        pos(i) = p: p = p - st
    Next i
    AlignPositions = na
End Function

Private Sub DrawVersion(ByVal ver As Long, mods() As Boolean, isFunc() As Boolean, ByVal size As Long)
    If ver < 7 Then Exit Sub
    Dim bch As Long: bch = ver
    Dim i As Long
    For i = 0 To 11
        Dim hi As Long: hi = (bch \ &H800&) And 1&           ' bit 11
        bch = (bch * 2) Xor (hi * &H1F25&)
    Next i
    Dim bits As Long: bits = (ver * &H1000&) Or bch          ' 18bit
    For i = 0 To 17
        Dim bit As Boolean: bit = GetBit(bits, i)
        Dim aa As Long: aa = size - 11 + (i Mod 3)
        Dim bb As Long: bb = i \ 3
        SetFunc mods, isFunc, aa, bb, bit
        SetFunc mods, isFunc, bb, aa, bit
    Next i
End Sub

Private Sub DrawFormatBits(ByVal ecl As Long, ByVal msk As Long, mods() As Boolean, isFunc() As Boolean, ByVal size As Long)
    Dim fb As Long
    Select Case ecl                                          ' フォーマットビット: L=1,M=0,Q=3,H=2
        Case 0: fb = 1
        Case 1: fb = 0
        Case 2: fb = 3
        Case 3: fb = 2
    End Select
    Dim data As Long: data = fb * 8 + msk
    Dim bch As Long: bch = data
    Dim i As Long
    For i = 0 To 9
        Dim hi As Long: hi = (bch \ &H200&) And 1&           ' bit 9
        bch = (bch * 2) Xor (hi * &H537&)
    Next i
    Dim bits As Long: bits = ((data * &H400&) Or bch) Xor &H5412&   ' 15bit

    For i = 0 To 5
        SetFunc mods, isFunc, 8, i, GetBit(bits, i)
    Next i
    SetFunc mods, isFunc, 8, 7, GetBit(bits, 6)
    SetFunc mods, isFunc, 8, 8, GetBit(bits, 7)
    SetFunc mods, isFunc, 7, 8, GetBit(bits, 8)
    For i = 9 To 14
        SetFunc mods, isFunc, 14 - i, 8, GetBit(bits, i)
    Next i
    For i = 0 To 7
        SetFunc mods, isFunc, size - 1 - i, 8, GetBit(bits, i)
    Next i
    For i = 8 To 14
        SetFunc mods, isFunc, 8, size - 15 + i, GetBit(bits, i)
    Next i
    SetFunc mods, isFunc, 8, size - 8, True                  ' 常時黒モジュール
End Sub

Private Sub DrawCodewords(data() As Long, mods() As Boolean, isFunc() As Boolean, ByVal size As Long)
    Dim i As Long: i = 0
    Dim total As Long: total = (UBound(data) + 1) * 8
    Dim right As Long
    For right = size - 1 To 1 Step -2
        If right = 6 Then right = 5                           ' 縦タイミング列を跨ぐ
        Dim vert As Long
        For vert = 0 To size - 1
            Dim j As Long
            For j = 0 To 1
                Dim x As Long: x = right - j
                Dim upward As Boolean: upward = (((right + 1) And 2) = 0)
                Dim y As Long
                If upward Then y = size - 1 - vert Else y = vert
                If (Not isFunc(x, y)) And i < total Then
                    mods(x, y) = GetBit(data(i \ 8), 7 - (i And 7))
                    i = i + 1
                End If
            Next j
        Next vert
    Next right
End Sub

Private Sub ApplyMask(ByVal msk As Long, mods() As Boolean, isFunc() As Boolean, ByVal size As Long)
    Dim x As Long, y As Long
    For y = 0 To size - 1
        For x = 0 To size - 1
            If Not isFunc(x, y) Then
                Dim inv As Boolean
                Select Case msk
                    Case 0: inv = ((x + y) Mod 2 = 0)
                    Case 1: inv = (y Mod 2 = 0)
                    Case 2: inv = (x Mod 3 = 0)
                    Case 3: inv = ((x + y) Mod 3 = 0)
                    Case 4: inv = (((x \ 3) + (y \ 2)) Mod 2 = 0)
                    Case 5: inv = ((((x * y) Mod 2)) + ((x * y) Mod 3) = 0)
                    Case 6: inv = ((((x * y) Mod 2) + ((x * y) Mod 3)) Mod 2 = 0)
                    Case 7: inv = ((((x + y) Mod 2) + ((x * y) Mod 3)) Mod 2 = 0)
                End Select
                If inv Then mods(x, y) = Not mods(x, y)
            End If
        Next x
    Next y
End Sub

'============================ マスク評価（ペナルティ） =======================

Private Function PenaltyScore(mods() As Boolean, ByVal size As Long) As Long
    Const N1 As Long = 3, N2 As Long = 3, N3 As Long = 40, N4 As Long = 10
    Dim result As Long: result = 0
    Dim x As Long, y As Long, k As Long
    Dim rh(0 To 6) As Long

    ' 規則1(横)+規則3(横)
    For y = 0 To size - 1
        Dim rc As Boolean: rc = False
        Dim rn As Long: rn = 0
        For k = 0 To 6: rh(k) = 0: Next k
        For x = 0 To size - 1
            If mods(x, y) = rc Then
                rn = rn + 1
                If rn >= 5 Then result = result + IIf(rn = 5, N1, 1)
            Else
                FinderAddHistory rn, rh, size
                If Not rc Then result = result + FinderCount(rh) * N3
                rc = mods(x, y): rn = 1
            End If
        Next x
        result = result + FinderTerminate(rc, rn, rh, size) * N3
    Next y

    ' 規則1(縦)+規則3(縦)
    For x = 0 To size - 1
        Dim rc2 As Boolean: rc2 = False
        Dim rn2 As Long: rn2 = 0
        For k = 0 To 6: rh(k) = 0: Next k
        For y = 0 To size - 1
            If mods(x, y) = rc2 Then
                rn2 = rn2 + 1
                If rn2 >= 5 Then result = result + IIf(rn2 = 5, N1, 1)
            Else
                FinderAddHistory rn2, rh, size
                If Not rc2 Then result = result + FinderCount(rh) * N3
                rc2 = mods(x, y): rn2 = 1
            End If
        Next y
        result = result + FinderTerminate(rc2, rn2, rh, size) * N3
    Next x

    ' 規則2: 同色 2x2 ブロック
    For y = 0 To size - 2
        For x = 0 To size - 2
            Dim c As Boolean: c = mods(x, y)
            If c = mods(x + 1, y) And c = mods(x, y + 1) And c = mods(x + 1, y + 1) Then result = result + N2
        Next x
    Next y

    ' 規則4: 黒モジュール比率
    Dim dark As Long: dark = 0
    For y = 0 To size - 1
        For x = 0 To size - 1
            If mods(x, y) Then dark = dark + 1
        Next x
    Next y
    Dim total As Long: total = size * size
    Dim kk As Long: kk = (Abs(dark * 20 - total * 10) + total - 1) \ total - 1
    result = result + kk * N4

    PenaltyScore = result
End Function

Private Sub FinderAddHistory(ByVal curLen As Long, rh() As Long, ByVal size As Long)
    If rh(0) = 0 Then curLen = curLen + size
    Dim k As Long
    For k = 6 To 1 Step -1
        rh(k) = rh(k - 1)
    Next k
    rh(0) = curLen
End Sub

Private Function FinderCount(rh() As Long) As Long
    Dim n As Long: n = rh(1)
    Dim core As Boolean
    core = (n > 0) And (rh(2) = n) And (rh(3) = n * 3) And (rh(4) = n) And (rh(5) = n)
    Dim c As Long: c = 0
    If core And (rh(0) >= n * 4) And (rh(6) >= n) Then c = c + 1
    If core And (rh(6) >= n * 4) And (rh(0) >= n) Then c = c + 1
    FinderCount = c
End Function

Private Function FinderTerminate(ByVal curColor As Boolean, ByVal curLen As Long, rh() As Long, ByVal size As Long) As Long
    If curColor Then
        FinderAddHistory curLen, rh, size
        curLen = 0
    End If
    curLen = curLen + size
    FinderAddHistory curLen, rh, size
    FinderTerminate = FinderCount(rh)
End Function

'============================ 容量テーブル ====================================

Private Function GetNumRawDataModules(ByVal ver As Long) As Long
    Dim size As Long: size = ver * 4 + 17
    Dim result As Long: result = size * size
    result = result - 8 * 8 * 3
    result = result - (15 * 2 + 1)
    result = result - (size - 16) * 2
    If ver >= 2 Then
        Dim na As Long: na = ver \ 7 + 2
        result = result - (na - 1) * (na - 1) * 25
        result = result - (na - 2) * 2 * 20
        If ver >= 7 Then result = result - 36
    End If
    GetNumRawDataModules = result
End Function

Private Function GetNumDataCodewords(ByVal ver As Long, ByVal ecl As Long) As Long
    GetNumDataCodewords = GetNumRawDataModules(ver) \ 8 - NEB(ecl, ver) * ECCB(ecl, ver)
End Function

Private Function ByteModeCharCountBits(ByVal ver As Long) As Long
    If ver <= 9 Then ByteModeCharCountBits = 8 Else ByteModeCharCountBits = 16
End Function

Private Sub EnsureTables()
    If mTablesReady Then Exit Sub
    ReDim ECCB(0 To 3, 0 To 40)
    ReDim NEB(0 To 3, 0 To 40)
    Dim a As Variant, e As Long, vv As Long

    ' -- ECCB : 1ブロックあたり誤り訂正符号語数（index 0 は番兵 -1）--
    a = Array(-1, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30)
    For vv = 0 To 40: ECCB(0, vv) = a(vv): Next vv
    a = Array(-1, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28)
    For vv = 0 To 40: ECCB(1, vv) = a(vv): Next vv
    a = Array(-1, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30)
    For vv = 0 To 40: ECCB(2, vv) = a(vv): Next vv
    a = Array(-1, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30)
    For vv = 0 To 40: ECCB(3, vv) = a(vv): Next vv

    ' -- NEB : 誤り訂正ブロック数（index 0 は番兵 -1）--
    a = Array(-1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25)
    For vv = 0 To 40: NEB(0, vv) = a(vv): Next vv
    a = Array(-1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49)
    For vv = 0 To 40: NEB(1, vv) = a(vv): Next vv
    a = Array(-1, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68)
    For vv = 0 To 40: NEB(2, vv) = a(vv): Next vv
    a = Array(-1, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81)
    For vv = 0 To 40: NEB(3, vv) = a(vv): Next vv

    mTablesReady = True
End Sub

'============================ 小物ヘルパ ======================================

Private Function ParseEcl(ByVal s As String) As Long
    Dim c As String: c = UCase$(Left$(Trim$(s) & "M", 1))
    Select Case c
        Case "L": ParseEcl = 0
        Case "Q": ParseEcl = 2
        Case "H": ParseEcl = 3
        Case Else: ParseEcl = 1
    End Select
End Function

Private Function GetBit(ByVal x As Long, ByVal i As Long) As Boolean
    GetBit = ((x \ CLng(2 ^ i)) And 1&) <> 0
End Function

Private Sub SetFunc(mods() As Boolean, isFunc() As Boolean, ByVal x As Long, ByVal y As Long, ByVal dark As Boolean)
    mods(x, y) = dark
    isFunc(x, y) = True
End Sub

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function
