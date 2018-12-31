'*******************************************************************************
'	動作説明	：	MCプロトコル
'	動作説明	：	通信制御のサンプルコードです。
'
'	注意		：	本ソフトはフリーソフトです。個人／団体／社内利用を問わず、
'					ご自由にお使い下さい。
'					なお，著作権は作者である"杉本卓也(杉本電気システム)"が
'					保有しています。
'					このソフトウェアを使用したことによって生じたすべての障害・損害・不具
'					合等に関しては、私と私の関係者および私の所属するいかなる団体・組織とも、
'					一切の責任を負いません。各自の責任においてご使用ください。
'
'
'	Copyright(c) 2018 T.SUGIMOTO(SUGIMOTO ELECTRIC SYSTEM SOLUTION). All rights reserved.
'*******************************************************************************
'   修正履歴
'    2018/12/31 T.Sugimoto 新規作成
'    
'*******************************************************************************
Imports System.Collections.Generic
Imports System.Collections
Imports PlcMelsecManager.PlcSystem.Data
Imports System.Text

Namespace PlcSystem.Reader
    ''' <summary>
    ''' PLCレジスタデータ
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PLCDataReader
        Const E71msec As Integer = 250
        Dim mObjSck As System.Net.Sockets.TcpClient
        Dim mObjStm As System.Net.Sockets.NetworkStream
        Dim mDeviceNameArray() As String = {"D300", "D312", "D320"}

        ''' <summary>
        ''' ソケットオープン
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Open()
            mObjSck = New System.Net.Sockets.TcpClient
            mObjSck.Connect("192.168.1.75", 8192)
            ' ソケットストリーム取得
            mObjStm = mObjSck.GetStream()
        End Sub

        ''' <summary>
        ''' 初期化
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Init()
            Dim dat() As Byte
            dat = MakeDiscreteData()
            SendData(dat)

        End Sub

        ''' <summary>
        ''' PLCデータ取得（連続領域）
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Read() As PLCData
            Dim plcData As New PLCData

            Dim dat() As Byte
            Dim 応答データ As String = ""

            'dat = MakeContinueData()
            dat = MakeMonitorData()

            応答データ = SendData(dat)
            MessageBox.Show(ConvertReceiveDeviceData(応答データ))

            Return plcData

        End Function

        ''' <summary>
        ''' ソケットクローズ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Close()
            mObjStm.Close()
            mObjSck.Close()
        End Sub

        ''' <summary>
        ''' PLC送信データ生成（連続領域）
        ''' </summary>
        ''' <remarks></remarks>
        Function MakeContinueData() As Byte()
            ' Dレジへのワード単位読み込み
            ' Dim dat As Byte() = {&H50, &H0, &H0, &HFF, &HFF, &H3, &H0, &HC, &H0, &H10, &H0, &H1, &H4, &H0, &H0, &H2C, &H1, &H0, &HA8, &H3, &H0}
            'Dレジへのワード単位書き込み。書き込みのときは書き込みデータにおうじて電文命令長が異なるので、「要求データ長」をちゃんと変更すること。
            '     Dim dat As Byte() = {&H50, &H0, &H0, &HFF, &HFF, &H3, &H0, &HE, &H0, &H10, &H0, &H1, &H14, &H0, &H0, &H2C, &H1, &H0, &HA8, &H1, &H0, &HFF, &HFF}

            Dim 送信データ() As Byte = {&HFA, &HFF}

            Dim dat(送信データ.Length + 20) As Byte
            dat(0) = &H50 : dat(1) = &H0 'サブヘッダであり固定
            dat(2) = &H0 : dat(3) = &HFF 'ネットワーク番号(PCとE71との通信であればH0, FF . CC-Linkなどを介する場合はマニュアルp74を参考に)
            dat(4) = &HFF : dat(5) = &H3 : dat(6) = &H0 '二重化CPUでないならFF, 03, 00. その他ならマニュアルp102を参考に

            'dat(7) = 12 + 送信データ.Length : dat(8) = &H0 '送信データ長
            dat(7) = &HC : dat(8) = &H0 '送信データ長

            dat(9) = T250msec2HEX(E71msec, 0) : dat(10) = T250msec2HEX(E71msec, 1)   'CPU監視タイマ

            ''dat(11) = &H1 : dat(12) = &H14         'コマンド
            dat(11) = &H1 : dat(12) = &H4         'コマンド
            dat(13) = &H0 : dat(14) = &H0  'サブコマンド
            dat(15) = &H2C : dat(16) = &H1 : dat(17) = &H0   '先頭デバイス(16進数6桁 D300
            dat(18) = &HA8
            dat(19) = &H1
            dat(20) = &H0
            Array.Copy(送信データ, 0, dat, 21, 送信データ.Length)

            Return dat

        End Function


        ''' <summary>
        ''' PLC送信データ生成（離散領域）
        ''' </summary>
        ''' <remarks></remarks>
        Function MakeDiscreteData() As Byte()
            Dim 送信データ() As Byte = {&HFA, &HFF}
            Dim deviceCodeArray() As Byte

            Dim dat(送信データ.Length + 28) As Byte
            dat(0) = &H50 : dat(1) = &H0 'サブヘッダであり固定
            dat(2) = &H0 : dat(3) = &HFF 'ネットワーク番号(PCとE71との通信であればH0, FF . CC-Linkなどを介する場合はマニュアルp74を参考に)
            dat(4) = &HFF : dat(5) = &H3 : dat(6) = &H0 '二重化CPUでないならFF, 03, 00. その他ならマニュアルp102を参考に

            'dat(7) = 20 + 送信データ.Length : dat(8) = &H0 '送信データ長
            dat(7) = &H14 : dat(8) = &H0 '送信データ長

            dat(9) = T250msec2HEX(E71msec, 0) : dat(10) = T250msec2HEX(E71msec, 1)   'CPU監視タイマ

            dat(11) = &H1 : dat(12) = &H8  'コマンド
            dat(13) = &H0 : dat(14) = &H0  'サブコマンド
            dat(15) = &H3  '16bit データ点数
            dat(16) = &H0  '32bit データ点数

            'メモリ設定
            For i = 0 To mDeviceNameArray.Length - 1
                deviceCodeArray = ConvertSendDeviceCode(mDeviceNameArray(i))
                Array.Copy(deviceCodeArray, 0, dat, 17 + 4 * i, 4)
            Next

            Array.Copy(送信データ, 0, dat, 29, 送信データ.Length)

            Return dat

        End Function

        ''' <summary>
        ''' PLC送信データ生成（モニタ要求）
        ''' </summary>
        ''' <remarks></remarks>
        Function MakeMonitorData() As Byte()
            Dim 送信データ() As Byte = {&HFA, &HFF}

            Dim dat(送信データ.Length + 14) As Byte
            dat(0) = &H50 : dat(1) = &H0 'サブヘッダであり固定
            dat(2) = &H0 : dat(3) = &HFF 'ネットワーク番号(PCとE71との通信であればH0, FF . CC-Linkなどを介する場合はマニュアルp74を参考に)
            dat(4) = &HFF : dat(5) = &H3 : dat(6) = &H0 '二重化CPUでないならFF, 03, 00. その他ならマニュアルp102を参考に

            'dat(7) = 6 + 送信データ.Length : dat(8) = &H0 '送信データ長
            dat(7) = &H6 : dat(8) = &H0 '送信データ長

            dat(9) = T250msec2HEX(E71msec, 0) : dat(10) = T250msec2HEX(E71msec, 1)   'CPU監視タイマ

            dat(11) = &H2 : dat(12) = &H8  'コマンド
            dat(13) = &H0 : dat(14) = &H0  'サブコマンド

            Array.Copy(送信データ, 0, dat, 15, 送信データ.Length)

            Return dat

        End Function

        ''' <summary>
        ''' ソケット送信・受信
        ''' </summary>
        ''' <remarks></remarks>
        Function SendData(ByVal sendDat() As Byte) As String
            Dim 応答データ As String = ""
            Dim data As Byte()

            ' ソケット送信
            mObjStm.Write(sendDat, 0, sendDat.GetLength(0))
            System.Threading.Thread.Sleep(250)

            ' ソケット受信
            If mObjSck.Available > 0 Then
                data = New Byte(mObjSck.Available - 1) {}
                mObjStm.Read(data, 0, data.GetLength(0))


                For i = 0 To data.Length - 1
                    応答データ = 応答データ & Convert.ToString(data(i), 16).PadLeft(2, "0")
                Next
            End If

            Return 応答データ
        End Function

        ''' <summary>
        ''' 時間 HEX変換
        ''' </summary>
        ''' <remarks></remarks>
        Function T250msec2HEX(ByVal msec As Integer, ByVal i As Integer) As Integer 'msecをHEXに変換する
            Dim hexi As Integer
            If msec > 38796720 Then : msec = 38796720 : End If : If msec < 0 Then : msec = 0 : End If
            If i = 0 Then : hexi = CInt("&H" & CInt(msec / 250).ToString("x4").Substring(0, 2)) : End If
            If i = 1 Then : hexi = CInt("&H" & CInt(msec / 250).ToString("x4").Substring(2, 2)) : End If
            If i = 2 Then : hexi = CInt(msec / 250) : End If
            Return hexi
        End Function

        ''' <summary>
        ''' PLC デバイスコード変換
        ''' </summary>
        ''' <remarks></remarks>
        Function ConvertSendDeviceCode(ByVal deviceName As String) As Byte()
            Dim strDeviceType As String      'デバイス種別 ex "D","M" etc
            Dim strDeviceNumber As String    'デバイス番号
            Dim deviceData(3) As Byte

            strDeviceType = deviceName.Substring(0, 1)
            strDeviceNumber = deviceName.Substring(1)

            Dim iDeviceNumber As Integer = Integer.Parse(strDeviceNumber)

            '16進文字列に変換
            Dim hexText As String = iDeviceNumber.ToString("x6")
            Dim hexChar As String = ""

            '16進文字列をbyteに変換
            Dim decData(2) As Byte
            For i = 0 To 2
                hexChar = hexText.Substring(i * 2, 2)
                decData(i) = Convert.ToByte(Convert.ToInt32(hexChar, 16))
            Next

            'リトルエンディアンに変換
            Array.Reverse(decData)
            Array.Copy(decData, 0, deviceData, 0, 3)

            Select Case strDeviceType
                Case "D"
                    deviceData(3) = &HA8
                Case "M"
                    deviceData(3) = &H90
                Case Else
                    deviceData(3) = &HA8
            End Select

            Return deviceData
        End Function


        ''' <summary>
        ''' PLC 受信データ変換
        ''' </summary>
        ''' <remarks></remarks>
        Function ConvertReceiveDeviceData(ByVal receiveDevice As String) As String
            '16進文字列に変換
            Dim hexText As String = receiveDevice.Substring(22)
            Dim hexChar As String = ""
            Dim receiveDeviceData As String = ""

            '16進文字列をbyteに変換
            Dim decData(1) As Byte
            For i = 0 To (hexText.Length / 4 - 1)
                For j = 0 To 1
                    hexChar = hexText.Substring(i * 4 + j * 2, 2)
                    decData(j) = Convert.ToByte(Convert.ToInt32(hexChar, 16))
                Next
                receiveDeviceData = receiveDeviceData & BitConverter.ToInt16(decData, 0) & ","
            Next

            Return receiveDeviceData
        End Function
    End Class
End Namespace

