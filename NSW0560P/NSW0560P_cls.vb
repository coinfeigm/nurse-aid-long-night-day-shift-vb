Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports Microsoft.Office.Interop
<System.Runtime.InteropServices.ProgId("clsNSW0560P_NET.clsNSW0560P")> Public Class clsNSW0560P
    Inherits General.PrintBase

    '/**********************************************************************/
    '/
    '/    ｼｽﾃﾑ名称：看護支援システム
    '/ ﾌﾟﾛｸﾞﾗﾑ名称：長夜勤基準(長夜勤-長日勤)
    '/        ＩＤ: NSW0560P
    '/        概要: 出力「長い夜のベンチマーク（長い夜仕事 - ロングデイサービス）」
    '/
    '/      作成者: Darren   CREATE 2018/09/17   【P-09480】 
    '/
    '/**********************************************************************/

    Private m_Excel As Excel.Application
    Private m_Sheet As Excel.Worksheet

    'フォームの条件
    Private m_FromYMD As Object  '出力期間 From [ﾌｫｰﾑ・帳票間共通]
    Private m_ToYMD As Object    '出力期間 To [ﾌｫｰﾑ・帳票間共通]
    Private m_BaseYMD As Object  '基準日 [ﾌｫｰﾑ・帳票間共通]
    Private m_FromLng As Integer 'ﾃﾞｰﾀ取得 開始日
    Private m_ToLng As Integer   'ﾃﾞｰﾀ取得 終了日
    Private m_BaseLng As Integer 'ﾃﾞｰﾀ取得 基準日(代休の有効期限チェックに使用)
    Private m_SaveSortDefault As String  '選択表示順保持変数
    Private m_KinmuDeptCD As String      '選択勤務部署コード
    Private m_KinmuDeptNM As String      '選択勤務部署名称
    Private m_AuthRangeKbn As String     '参照権限区分

    'データ開始
    Private Const DATA_START As Short = 5

    'アイテム設定
    Private m_ChiefCD As String
    Private m_NagaNikkinKinmuCD As String
    Private m_NagaYakinKinmuCD As String

    'チーフ
    Private Structure Chief_Type
        Dim Name As String
        Dim PostCD As String
    End Structure
    Private m_ChiefData() As Chief_Type
    Private m_ChiefName As String

    Private Structure IdoHistory_Type
        Dim StartDate As Integer        '開始日
        Dim EndDate As Integer          '終了日
        Dim strCD As String             'コード
        Dim strNm As String             '氏名
    End Structure
    '出力データ
    Private Structure Output_Type
        Dim StaffMngId As String
        Dim Name As String
        Dim PostCD As String
        Dim IdoHistory() As IdoHistory_Type
        Dim SaiyoHistory() As IdoHistory_Type
        Dim SecName As String
        Dim KangoTaniCD As String
        Dim KinmuCumulative As Dictionary(Of Integer, Nullable(Of Integer))
        Dim KinmuRemaining As Integer
        Dim KinmuGrandTotal As Integer
    End Structure
    Private m_StaffData() As Output_Type

    Public Overrides Sub mPrintOut()
        Try
            General.g_ErrorProc = "clsNSW0560P mPrintOut"

            Dim w_strMsg() As String
            Dim w_strSelKinmuDept As Object
            Dim w_strSelKinmuDeptNM As Object

            If m_objComPrint Is Nothing Then
                m_objComPrint = New NsAid_NSC0020C.clsNSC0020C
                With m_objComPrint
                    .pfrmCondition = m_ProcessForm
                    .pListID = m_ListID
                    .pHospitalCd = m_HospitalCD
                End With
            End If

            If m_ResorceCD = "" Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "リソースコード"
                Call General.paMsgDsp("NS0156", w_strMsg)
                Exit Sub
            End If

            m_AuthRangeKbn = ""
            With General.g_objSecurity
                .pUserKinmuDeptCD = General.g_strUserKinmuDeptCD

                If General.paCheckSelStaffAccount(m_ResorceCD, Trim(m_UserMngID), Trim(General.g_strUserKinmuDeptCD)) Then
                    m_AuthRangeKbn = .fR_RefAuthRangeKbn
                End If

                If m_AuthRangeKbn = "" Then
                    ReDim w_strMsg(2)
                    w_strMsg(1) = "帳票"
                    w_strMsg(2) = "出力"
                    Call General.paMsgDsp("NS0093", w_strMsg)
                    Exit Sub
                End If
            End With

            With m_objComPrint
                .pStopFlg = False

                Call mSetCondition()

                If .mfncStopPrint Then Exit Sub

                Call GetItemValue()

                If m_KinmuDeptCD <> "" Then
                    .mPrintForm_Show()
                    System.Windows.Forms.Application.DoEvents()
                    Call .mPrintForm_Caption("開始プロセス...")

                    w_strSelKinmuDept = General.paSplit(m_KinmuDeptCD, ",")
                    w_strSelKinmuDeptNM = General.paSplit(m_KinmuDeptNM, ",")

                    For w_i = 0 To UBound(w_strSelKinmuDept)
                        m_KangoTaniCD = w_strSelKinmuDept(w_i)
                        m_KangoTaniName = w_strSelKinmuDeptNM(w_i)

                        If m_PreviewFlg > General.G_INTPRINTOUT Then
                            .pEndStatus = General.G_BUTTON_ENDSTATUS_ENUM.PREVIEW_BUTTON
                        Else
                            .pEndStatus = General.G_BUTTON_ENDSTATUS_ENUM.PRINT_BUTTON
                        End If

                        If m_KangoTaniCD <> "" Then
                            .pStopFlg = False
                            If GetStaff() Then
                                If Not .mfncStopPrint Then Call GetKinmuRemaining()
                                If Not .mfncStopPrint Then Call GetKinmuCumulative()
                                If Not .mfncStopPrint Then Call PrintControl()
                                If .mfncStopPrint Then Exit Sub
                            End If
                        End If

                        Erase m_StaffData
                    Next w_i
                End If
            End With

            Erase m_StaffData

        Catch ex As Exception
            m_objComPrint.mExcelClose()
            m_Sheet = Nothing
            m_Excel = Nothing

            Call General.paDllTrpMsg(Convert.ToString(Err.Number), General.g_ErrorProc)
        End Try
    End Sub

    Public Overrides Sub PrintControl()
        Try
            General.g_ErrorProc = "clsNSW0560P PrintControl"

            m_objComPrint.mGetExcelAppli(m_ListID)
            m_Excel = m_objComPrint.pExcelAppli
            m_Sheet = m_objComPrint.pWorksheet

            Call m_objComPrint.mPrintForm_Caption("データの準備...")
            Call m_objComPrint.mPrintForm_Disp(0, UBound(m_StaffData))

            Call PrintHeader()

            Call PrintDetail()

            'A4
            m_objComPrint.pZoom = 70
            m_objComPrint.pPaperSize = 9
            m_objComPrint.pPrinterID = General.paGetItemValue(General.G_STRMAINKEY11, General.G_STRSUBKEY12, General.G_STRPRIKEY1, "", m_HospitalCD)

            If Not m_objComPrint.mfncStopPrint Then
                m_objComPrint.mOutputControl()
                If m_PreviewFlg = General.G_INTPREVIEWSAVE Then
                    If m_objComPrint.mfncExcelSaveInit("", True) = True Then
                        Call m_objComPrint.mOutputControl(m_ListID)
                    End If
                End If
            End If

            m_objComPrint.mExcelClose()
            m_Sheet = Nothing
            m_Excel = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintHeader()
        Try
            General.g_ErrorProc = "clsNSW0560P PrintHeader"

            With m_Sheet
                .Range("A1").Value = .Range("A1").Value & m_KangoTaniName
                .Range("A2").Value = .Range("A2").Value & m_ChiefName
                .Range("A3").Value = m_ListName
                .Range("O1").Value = .Range("O1").Value & Format(m_FromYMD, "yyyy年度").ToString
                .Range("O2").Value = .Range("O2").Value & Format(m_BaseYMD, "yyyy/MM/dd").ToString
            End With

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintDetail()
        Try
            General.g_ErrorProc = "clsNSW0560P PrintDetail"

            Dim w_Int As Integer
            Dim w_strNextRow As String
            Dim w_RangeCopy As Excel.Range

            w_RangeCopy = m_Sheet.Range("A5:Q5")

            For w_Int = 1 To UBound(m_StaffData)
                If m_objComPrint.mfncStopPrint Then Exit Sub
                w_strNextRow = (w_Int + DATA_START - 1).ToString

                m_Sheet.Range("A" & w_strNextRow).Value = w_Int
                m_Sheet.Range("B" & w_strNextRow).Value = m_StaffData(w_Int).Name
                m_Sheet.Range("C" & w_strNextRow).Value = m_StaffData(w_Int).SecName
                m_Sheet.Range("D" & w_strNextRow).Value = m_StaffData(w_Int).KinmuRemaining.ToString
                m_Sheet.Range("E" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(4).ToString
                m_Sheet.Range("F" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(5).ToString
                m_Sheet.Range("G" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(6).ToString
                m_Sheet.Range("H" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(7).ToString
                m_Sheet.Range("I" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(8).ToString
                m_Sheet.Range("J" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(9).ToString
                m_Sheet.Range("K" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(10).ToString
                m_Sheet.Range("L" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(11).ToString
                m_Sheet.Range("M" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(12).ToString
                m_Sheet.Range("N" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(1).ToString
                m_Sheet.Range("O" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(2).ToString
                m_Sheet.Range("P" & w_strNextRow).Value = m_StaffData(w_Int).KinmuCumulative(3).ToString
                m_Sheet.Range("Q" & w_strNextRow).Value = m_StaffData(w_Int).KinmuGrandTotal.ToString

                If m_StaffData.Length - 1 >= w_Int + 1 Then
                    w_strNextRow = (w_Int + DATA_START).ToString
                    w_RangeCopy.Copy(m_Sheet.Range("A" & w_strNextRow & ":Q" & w_strNextRow))

                    If w_Int = 32 Then
                        m_Sheet.HPageBreaks.Add(m_Sheet.Range("A" & w_strNextRow))
                    End If
                End If

                Call m_objComPrint.mPrintForm_Disp(w_Int, UBound(m_StaffData))
            Next w_Int

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub mSetCondition()
        Try
            General.g_ErrorProc = "clsNSW0560P mSetCondition"

            Dim w_Form As PRINTFORM.frmNSZ0010H
            Dim w_Array As Object
            Dim w_SortArray As Object
            Dim w_KinmuDept As Object

            If m_FromYMD = Nothing Then
                m_FromYMD = fncNendo(General.DateUtil.paGetDateIntegerFromDate(Now)) & "/04/01"
                m_BaseYMD = General.DateUtil.paGetDateStringFromDate(Now, General.G_DATE_ENUM.yyyy_MM_dd)
            End If

            w_Form = New PRINTFORM.frmNSZ0010H

            With w_Form
                w_Array = New Object() {m_ListID, m_ListName}
                w_Form.pLabelLet = w_Array
                w_Form.pNendoLet = m_FromYMD
                w_Form.pBaseDateLet = m_BaseYMD
                w_Form.pResorceCDLet = m_ResorceCD
                w_Form.pSelRangeKBNLet = m_SelectRangeKBN
                w_Form.pAuthRangeKbnLet = m_AuthRangeKbn

                If m_SortChange = "1" Then
                    If m_SaveSortDefault <> "" Then
                        w_SortArray = New Object() {m_SaveSortDefault, m_SortChange}
                        w_Form.pOutPutOrderIn = w_SortArray
                    Else
                        w_SortArray = New Object() {m_SortDefault, m_SortChange}
                        w_Form.pOutPutOrderIn = w_SortArray
                    End If
                End If

                w_Form.pDispMode = General.G_DISPMODE_ENUM.G_DISPMODE3

                .ShowDialog(pProcessObj)

                If .pEndStatus = General.G_ENDSTATUS_ENUM.G_END_OK Then
                    m_FromYMD = .pNendoGet
                    m_ToYMD = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, m_FromYMD))
                    m_BaseYMD = .pBaseDateGet

                    m_FromLng = General.DateUtil.paGetDateIntegerFromDate(m_FromYMD)
                    m_ToLng = General.DateUtil.paGetDateIntegerFromDate(m_ToYMD)
                    m_BaseLng = General.DateUtil.paGetDateIntegerFromDate(m_BaseYMD)

                    w_KinmuDept = .pKinmuDeptCDGet '勤務部署情報
                    m_KinmuDeptCD = w_KinmuDept(0) '選択勤務部署コード
                    m_KinmuDeptNM = w_KinmuDept(1) '選択勤務部署名称

                    If m_SortChange = "1" Then
                        m_SortDefault = w_Form.pOutPutOrderOut
                        m_SaveSortDefault = w_Form.pOutPutOrderOut
                        If m_SortDefault = 0 Then
                            m_SortDefault = 1
                        End If
                    Else
                        If m_SortDefault = "" Then
                            m_SortDefault = 1
                        End If
                    End If
                ElseIf .pEndStatus = General.G_ENDSTATUS_ENUM.G_END_CANCEL Then
                    m_objComPrint.pStopFlg = True
                End If
            End With

            w_Form.Dispose()

        Catch ex As Exception
            Call General.paDllTrpMsg(Convert.ToString(Err.Number), General.g_ErrorProc)
        End Try
    End Sub

    Private Sub GetItemValue()
        Try
            General.g_ErrorProc = "clsNSW0560P GetItemValue"

            Dim w_Int As Integer
            Dim w_str As String
            Dim w_strCD() As String

            w_str = General.paGetItemValue(General.G_STRMAINKEY11, m_ListID, "CHIEFCD", Convert.ToString(0), m_HospitalCD)
            w_strCD = Split(w_str, ",")
            If UBound(w_strCD) >= 0 Then
                For w_Int = 0 To UBound(w_strCD)
                    m_ChiefCD = m_ChiefCD & General.paFormatSpace(w_strCD(w_Int), 10) & ","
                Next
            End If

            w_str = General.paGetItemValue(General.G_STRMAINKEY11, m_ListID, "NAGANIKKINKINNMUCD", Convert.ToString(0), m_HospitalCD)
            w_strCD = Split(w_str, ",")
            If UBound(w_strCD) >= 0 Then
                For w_Int = 0 To UBound(w_strCD)
                    m_NagaNikkinKinmuCD = m_NagaNikkinKinmuCD & General.paFormatSpace(w_strCD(w_Int), 10) & ","
                Next
            End If

            w_str = General.paGetItemValue(General.G_STRMAINKEY11, m_ListID, "NAGAYAKINKINMUCD", Convert.ToString(0), m_HospitalCD)
            w_strCD = Split(w_str, ",")
            If UBound(w_strCD) >= 0 Then
                For w_Int = 0 To UBound(w_strCD)
                    m_NagaYakinKinmuCD = m_NagaYakinKinmuCD & General.paFormatSpace(w_strCD(w_Int), 10) & ","
                Next
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    '職員情報を取得する関数
    Private Function GetStaff() As Boolean
        Try
            General.g_ErrorProc = "clsNSW0560P GetItemValue"

            Dim w_Count As Short
            Dim w_Int As Short
            Dim w_StaffCount As Short
            Dim w_TempStaffID As String
            Dim w_IdoHistoryCount As Short
            Dim w_blnResult As Boolean
            Dim w_StaffID As String
            Dim w_HaizokuFlg As Boolean
            Dim w_Int2 As Short
            Dim w_CurrentPost As Integer
            Dim w_TmpInt As Integer
            Dim w_strMsg() As String

            Dim w_StaffWork() As Output_Type

            Call m_objComPrint.mPrintForm_Caption(General.g_strSelKinmuDeptNm & " 職員データ取得中...")

            GetStaff = False

            '初期化
            Erase w_StaffWork

            With General.g_objGetData
                '---------- 出力対象職員情報の取得用のプロパティ設定 ---------
                .p病院CD = m_HospitalCD
                .pソート順 = 0
                .p優先順位 = 0
                .p日付区分 = 1
                .p開始年月日 = m_FromLng
                .p終了年月日 = m_ToLng
                .p対象CD = m_KangoTaniCD
                .p指定表示順 = m_SortDefault
                '---------------------------------------------------------

                '職員情報取得
                w_blnResult = .mGetStaff

                '配列変数の初期化
                ReDim w_StaffWork(0)
                If w_blnResult = False Then
                    '[出力対象の職員が存在しない場合(エラー)]

                    '---------- エラーメッセージ表示 -------------------------
                    ReDim w_strMsg(1)
                    w_strMsg(1) = m_KangoTaniName & Space(2) & "出力対象の職員"
                    Call General.paMsgDsp("NS0008", w_strMsg)
                    '--------------------------------------------------------

                    Exit Function
                Else
                    '[出力対象の職員が存在する場合]

                    w_Count = .f職員件数

                    'データ件数分の配列確保
                    ReDim w_StaffWork(w_Count)

                    '初期化
                    w_TempStaffID = ""
                    w_StaffCount = 0

                    'データ件数分Loop
                    For w_Int = 1 To w_Count
                        System.Windows.Forms.Application.DoEvents()

                        .p職員索引 = w_Int
                        w_StaffID = .f職員管理番号2

                        '看護単位出力か対象職員の場合のみ以下の処理を行う
                        If w_TempStaffID <> w_StaffID Then
                            '[新しい職員の場合]
                            w_StaffCount = w_StaffCount + 1
                            w_StaffWork(w_StaffCount).PostCD = .f役職CD2
                            ReDim w_StaffWork(w_StaffCount).SaiyoHistory(0)

                            '---------- 採用歴取得用のプロパティ設定 --------------------
                            General.g_objIdoData.pHospitalCD = m_HospitalCD
                            General.g_objIdoData.pStaffMngID = w_StaffID
                            General.g_objIdoData.pDateFlg = 1
                            General.g_objIdoData.pDateFrom = m_FromLng
                            General.g_objIdoData.pDateTo = m_ToLng
                            General.g_objIdoData.pSortFlg = 0
                            '------------------------------------------------------

                            '採用歴取得と取得結果の判別
                            If General.g_objIdoData.mGetSaiyoIdo = True Then
                                '[採用履歴が取得できた場合]

                                ReDim w_StaffWork(w_StaffCount).SaiyoHistory(General.g_objIdoData.fSI_SaiyoCount)

                                '採用履歴の数だけループ
                                For w_Int2 = 1 To General.g_objIdoData.fSI_SaiyoCount
                                    'インデックス値取得
                                    General.g_objIdoData.mSI_SaiyoIdx = w_Int2

                                    '採用コード取得
                                    w_StaffWork(w_StaffCount).SaiyoHistory(w_Int2).strCD = General.g_objIdoData.fSI_EmpCD
                                    '採用年月日取得
                                    w_StaffWork(w_StaffCount).SaiyoHistory(w_Int2).StartDate = General.g_objIdoData.fSI_EmpDate
                                    '転退年月日取得
                                    w_StaffWork(w_StaffCount).SaiyoHistory(w_Int2).EndDate = General.g_objIdoData.fSI_RetireDate
                                    If w_StaffWork(w_StaffCount).SaiyoHistory(w_Int2).EndDate = 0 Then
                                        '[終了日の日付が０の場合]
                                        w_StaffWork(w_StaffCount).SaiyoHistory(w_Int2).EndDate = 99999999   '置き換える
                                    End If
                                Next w_Int2
                            End If

                            '---------- 勤務部署歴取得用のプロパティ設定 ----------------
                            General.g_objIdoData.pHospitalCD = m_HospitalCD
                            General.g_objIdoData.pStaffMngID = w_StaffID
                            General.g_objIdoData.pDateFlg = 0
                            General.g_objIdoData.pDateFrom = m_BaseLng
                            General.g_objIdoData.pSortFlg = 0
                            '---------------------------------------------------------

                            If General.g_objIdoData.mGetKinmuDeptIdo = True Then
                                General.g_objIdoData.mKI_KinmuDeptIdx = General.g_objIdoData.fKI_KinmuDeptCount
                                w_StaffWork(w_StaffCount).KangoTaniCD = General.g_objIdoData.fKI_CD
                            End If

                            '長期休暇理由(略称) 
                            If General.g_objIdoData.mGetChoukyuInfo = True Then
                                For w_Int2 = 1 To General.g_objIdoData.fLL_ChoukyuCount
                                    General.g_objIdoData.mLL_ChoukyuIdx = w_Int2
                                    If m_BaseLng >= General.g_objIdoData.fLL_DateFrom And m_BaseLng <= General.g_objIdoData.fLL_DateTo Then
                                        w_StaffWork(w_StaffCount).SecName = General.g_objIdoData.fLL_SecName
                                    End If
                                Next w_Int2
                            End If

                            ReDim w_StaffWork(w_StaffCount).IdoHistory(0)
                            '---------- 勤務部署歴取得用のプロパティ設定 ----------------
                            General.g_objIdoData.pHospitalCD = m_HospitalCD
                            General.g_objIdoData.pStaffMngID = w_StaffID
                            General.g_objIdoData.pDateFlg = 0
                            General.g_objIdoData.pDateFrom = m_ToLng
                            General.g_objIdoData.pSortFlg = 0
                            '---------------------------------------------------------

                            '勤務部署歴取得と取得結果の判別
                            If General.g_objIdoData.mGetKinmuDeptIdo = True Then
                                '[勤務部署歴を取得できた場合]

                                For w_Int2 = 1 To General.g_objIdoData.fKI_KinmuDeptCount
                                    w_IdoHistoryCount = UBound(w_StaffWork(w_StaffCount).IdoHistory) + 1
                                    ReDim Preserve w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount)

                                    '勤務部署のインデックス値設定
                                    General.g_objIdoData.mKI_KinmuDeptIdx = w_Int2
                                    '勤務部署コード取得
                                    w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount).strCD = General.g_objIdoData.fKI_CD
                                    '勤務部署名
                                    w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount).strNm = General.g_objIdoData.fKI_Name
                                    '採用年月日取得
                                    w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount).StartDate = General.g_objIdoData.fKI_DateFrom
                                    '転退年月日取得
                                    w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount).EndDate = General.g_objIdoData.fKI_DateTo
                                    If w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount).EndDate = 0 Or w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount).EndDate > m_ToLng Then
                                        '[勤務部署異動情報の終了日が０の場合 or 指定期間の末日よりも未来日の場合]
                                        w_StaffWork(w_StaffCount).IdoHistory(w_IdoHistoryCount).EndDate = m_ToLng
                                    End If
                                Next w_Int2
                            End If
                        End If

                        '---------- 基本データの取得 -----------------------------------
                        w_StaffWork(w_StaffCount).StaffMngId = w_StaffID
                        w_StaffWork(w_StaffCount).Name = .f氏名2

                        Call m_objComPrint.mPrintForm_Disp(w_Int, w_Count)
                        '----------------------------------------------------------------

                        w_TempStaffID = w_StaffID '職員管理番号退避
                    Next w_Int
                End If
            End With

            ReDim Preserve w_StaffWork(w_StaffCount)

            w_StaffCount = 0
            For w_Int = 1 To UBound(w_StaffWork)

                '---------- 部署異動前かで絞り込み -------------------------------
                w_HaizokuFlg = False '初期化
                If w_StaffWork(w_Int).KangoTaniCD = m_KangoTaniCD Then
                    w_HaizokuFlg = True '出力対象
                End If
                '-------------------------------------------------------------

                If w_HaizokuFlg = True Then
                    w_StaffCount = w_StaffCount + 1
                    ReDim Preserve m_StaffData(w_StaffCount)
                    m_StaffData(w_StaffCount) = w_StaffWork(w_Int)
                End If
            Next w_Int

            'チーフを見つける
            w_CurrentPost = 99
            m_ChiefName = ""
            w_TmpInt = UBound(m_StaffData)
            For w_Int = 1 To w_TmpInt
                If Not m_StaffData(w_Int).PostCD = "" Then
                    w_TmpInt = InStr(m_ChiefCD, General.paFormatSpace(m_StaffData(w_Int).PostCD, 10))
                    If Not w_TmpInt = 0 And w_TmpInt < w_CurrentPost Then
                        m_ChiefName = m_StaffData(w_Int).Name
                        w_CurrentPost = w_TmpInt
                    End If
                End If
            Next w_Int

            'チーフを見つける
            If m_ChiefName = "" Then
                ReDim m_ChiefData(0)
                Call GetKenmu(m_ChiefData)
                w_TmpInt = UBound(m_ChiefData)
                For w_Int = 1 To w_TmpInt
                    If Not m_ChiefData(w_Int).PostCD = "" Then
                        w_TmpInt = InStr(m_ChiefCD, General.paFormatSpace(m_ChiefData(w_Int).PostCD, 10))
                        If Not w_TmpInt = 0 And w_TmpInt < w_CurrentPost Then
                            m_ChiefName = m_ChiefData(w_Int).Name
                            w_CurrentPost = w_TmpInt
                        End If
                    End If
                Next w_Int
            End If

            GetStaff = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    '昨年度以前の残
    Private Sub GetKinmuRemaining()
        Try
            General.g_ErrorProc = "clsNSW0560P GetKinmuRemaining"

            Dim w_Int As Integer
            Dim w_Int2 As Integer
            Dim w_Int3 As Integer
            Dim w_Count As Integer
            Dim w_Date As Integer
            Dim w_KinmuCount As List(Of String)
            Dim w_lstNikkin As List(Of String)
            Dim w_lstYakin As List(Of String)
            Dim w_TmpStr As String

            Call m_objComPrint.mPrintForm_Caption("実績データ取得中...")
            Call m_objComPrint.mPrintForm_Disp(0, UBound(m_StaffData))

            With General.g_objGetData
                For w_Int = 1 To UBound(m_StaffData)
                    '----------- ジョブを取得するためのプロパティの設定 ------------
                    .p病院CD = m_HospitalCD
                    .p職員区分 = 0
                    .p確定部署CD = ""
                    .p職員番号 = m_StaffData(w_Int).StaffMngId
                    .p日付区分 = 1
                    .p開始年月日 = m_StaffData(w_Int).SaiyoHistory(1).StartDate
                    .p終了年月日 = Integer.Parse(Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, m_FromYMD), "yyyyMMdd"))
                    .p勤務取得区分 = 1
                    '----------------------------------------------------

                    If .mGetKinmu Then
                        If m_objComPrint.mfncStopPrint Then Exit Sub
                        w_Count = .f期間日数
                        w_KinmuCount = New List(Of String)

                        '------- 各職務について、部門雇用歴で仕事が行われているかどうかを比較する -------
                        For w_Int2 = 1 To w_Count
                            w_Date = Integer.Parse(Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, w_Int2 - 1, General.paGetDateFromDateInteger(m_StaffData(w_Int).SaiyoHistory(1).StartDate)), "yyyyMMdd"))

                            For w_Int3 = 1 To UBound(m_StaffData(w_Int).IdoHistory)
                                If w_Date >= m_StaffData(w_Int).IdoHistory(w_Int3).StartDate And w_Date <= m_StaffData(w_Int).IdoHistory(w_Int3).EndDate _
                                And m_StaffData(w_Int).IdoHistory(w_Int3).strCD = m_KangoTaniCD Then
                                    .p実績年月日 = w_Date
                                    If .f実績勤務CD IsNot Nothing Then
                                        w_KinmuCount.Add(.f実績勤務CD)
                                    End If
                                    Exit For
                                End If
                            Next w_Int3
                        Next w_Int2
                        '--------------------------------------------------------------------------

                        ''-------------- フィルター作業 --------------
                        w_lstNikkin = New List(Of String)
                        w_lstYakin = New List(Of String)
                        For Each w_TmpStr In w_KinmuCount
                            If InStr(m_NagaNikkinKinmuCD, General.paFormatSpace(w_TmpStr, 10)) Then
                                w_lstNikkin.Add(w_TmpStr)
                            End If
                            If InStr(m_NagaYakinKinmuCD, General.paFormatSpace(w_TmpStr, 10)) Then
                                w_lstYakin.Add(w_TmpStr)
                            End If
                        Next w_TmpStr
                        m_StaffData(w_Int).KinmuRemaining = w_lstYakin.Count - w_lstNikkin.Count
                        '-----------------------------------------
                    End If

                    Call m_objComPrint.mPrintForm_Disp(w_Int, UBound(m_StaffData))
                Next w_Int
            End With

        Catch ex As Exception
            Throw
        End Try
    End Sub

    '月ごとの累計結果
    Private Sub GetKinmuCumulative()
        Try
            General.g_ErrorProc = "clsNSW0560P GetKinmuCumulative"

            Dim w_Int As Integer
            Dim w_Int2 As Integer
            Dim w_Int3 As Integer
            Dim w_MonthCount As Integer
            Dim w_Count As Integer
            Dim w_MonthStart As Integer
            Dim w_MonthEnd As Integer
            Dim w_Date As Integer
            Dim w_Month As Integer
            Dim w_lstKinmu As Dictionary(Of Integer, List(Of String))
            Dim w_TmpStr As String
            Dim w_lstNikkin As List(Of String)
            Dim w_lstYakin As List(Of String)
            Dim w_TmpSum As Integer
            Dim w_Kvp As KeyValuePair(Of Integer, Nullable(Of Integer))

            Call m_objComPrint.mPrintForm_Caption("実績データ取得中...")
            Call m_objComPrint.mPrintForm_Disp(0, UBound(m_StaffData))

            For w_Int = 1 To UBound(m_StaffData)
                m_StaffData(w_Int).KinmuCumulative = New Dictionary(Of Integer, Nullable(Of Integer))
                w_lstKinmu = New Dictionary(Of Integer, List(Of String))
                For w_MonthCount = 1 To 12
                    w_MonthStart = General.paGetDateIntegerFromDate(DateAdd(Microsoft.VisualBasic.DateInterval.Month, w_MonthCount - 1, m_FromYMD))
                    w_MonthEnd = General.paGetDateIntegerFromDate(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, DateAdd(Microsoft.VisualBasic.DateInterval.Month, w_MonthCount, m_FromYMD)))

                    w_Month = w_MonthStart.ToString.Substring(4, 2)
                    m_StaffData(w_Int).KinmuCumulative.Add(w_Month, Nothing)
                    w_lstKinmu.Add(w_Month, New List(Of String))

                    If Get_KakuteiF(w_MonthStart, w_MonthEnd) Then '計画に存在する
                        With General.g_objGetData
                            '----------- ジョブを取得するためのプロパティの設定 ------------
                            .p病院CD = m_HospitalCD
                            .p職員区分 = 0
                            .p確定部署CD = ""
                            .p職員番号 = m_StaffData(w_Int).StaffMngId
                            .p日付区分 = 1
                            .p開始年月日 = w_MonthStart
                            .p終了年月日 = w_MonthEnd
                            .p勤務取得区分 = 1
                            '----------------------------------------------------

                            If .mGetKinmu Then
                                If m_objComPrint.mfncStopPrint Then Exit Sub
                                w_Count = .f期間日数

                                '------- 各職務について、部門雇用歴で仕事が行われているかどうかを比較する -------
                                For w_Int2 = 1 To w_Count '目標年度からすべての仕事を得る
                                    w_Date = Integer.Parse(Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, w_Int2 - 1, General.paGetDateFromDateInteger(w_MonthStart)), "yyyyMMdd"))

                                    For w_Int3 = 1 To UBound(m_StaffData(w_Int).IdoHistory)
                                        If w_Date >= m_StaffData(w_Int).IdoHistory(w_Int3).StartDate And w_Date <= m_StaffData(w_Int).IdoHistory(w_Int3).EndDate _
                                        And m_StaffData(w_Int).IdoHistory(w_Int3).strCD = m_KangoTaniCD Then
                                            .p実績年月日 = w_Date
                                            If .f実績勤務CD IsNot Nothing Then
                                                w_lstKinmu(w_Month).Add(.f実績勤務CD)
                                            End If
                                            Exit For
                                        End If
                                    Next w_Int3
                                Next w_Int2
                                '--------------------------------------------------------------------------

                                '-------------- フィルター作業 --------------
                                w_lstNikkin = New List(Of String)
                                w_lstYakin = New List(Of String)
                                If w_lstKinmu(w_Month).Count > 0 Then
                                    For Each w_TmpStr In w_lstKinmu(w_Month)
                                        If InStr(m_NagaNikkinKinmuCD, General.paFormatSpace(w_TmpStr, 10)) Then
                                            w_lstNikkin.Add(w_TmpStr)
                                        End If
                                        If InStr(m_NagaYakinKinmuCD, General.paFormatSpace(w_TmpStr, 10)) Then
                                            w_lstYakin.Add(w_TmpStr)
                                        End If
                                    Next w_TmpStr
                                    m_StaffData(w_Int).KinmuCumulative(w_Month) = w_lstYakin.Count - w_lstNikkin.Count
                                End If
                                '-------------------------------------------
                            Else
                                m_StaffData(w_Int).KinmuCumulative(w_Month) = 0 '仕事データなし
                            End If
                        End With
                    Else
                        m_StaffData(w_Int).KinmuCumulative(w_Month) = Nothing '存在しない場合は空白
                    End If
                Next w_MonthCount

                w_TmpSum = 0
                For Each w_Kvp In m_StaffData(w_Int).KinmuCumulative '総計
                    If w_Kvp.Value IsNot Nothing Then
                        w_TmpSum = w_TmpSum + w_Kvp.Value
                    End If
                Next
                m_StaffData(w_Int).KinmuGrandTotal = m_StaffData(w_Int).KinmuRemaining + w_TmpSum

                Call m_objComPrint.mPrintForm_Disp(w_Int, UBound(m_StaffData))
            Next w_Int

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub GetKenmu(ByRef p_lstChief() As Chief_Type)
        Try
            General.g_ErrorProc = "clsNSW0560P GetKenmu"

            Dim w_Sql As String
            Dim w_FromDate As Integer
            Dim w_ToDate As Integer
            Dim w_Count As Integer
            Dim w_Int As Integer
            Dim w_Rs As ADODB.Recordset
            Dim w_Name_F As ADODB.Field
            Dim w_PostCD_F As ADODB.Field

            w_FromDate = General.DateUtil.paGetDateIntegerFromDate(m_FromYMD, General.G_DATE_ENUM.yyyyMMdd)
            w_ToDate = General.DateUtil.paGetDateIntegerFromDate(m_ToYMD, General.G_DATE_ENUM.yyyyMMdd)

            w_Sql = ""
            w_Sql = w_Sql & "SELECT"
            w_Sql = w_Sql & " KE.POSTCD "
            w_Sql = w_Sql & ",ST.STAFFNAME "
            w_Sql = w_Sql & ",ST.STAFFMNGID "
            w_Sql = w_Sql & ",KE.IDODATE "
            w_Sql = w_Sql & "FROM"
            w_Sql = w_Sql & " NS_KENMUINFO_F KE "
            w_Sql = w_Sql & ",NS_STAFFBASISINFO_F ST "
            w_Sql = w_Sql & "WHERE KE.HOSPITALCD = '" & m_HospitalCD & "' "
            w_Sql = w_Sql & "  AND KE.HOSPITALCD = ST.HOSPITALCD "
            w_Sql = w_Sql & "  AND KE.STAFFMNGID = ST.STAFFMNGID "
            w_Sql = w_Sql & "  AND KE.IDODATE <= " & w_ToDate & " "
            w_Sql = w_Sql & "  AND (KE.ENDDATE >= " & w_FromDate & " "
            w_Sql = w_Sql & "  OR KE.ENDDATE = 0"
            w_Sql = w_Sql & "  OR KE.ENDDATE = 99999999 "
            w_Sql = w_Sql & "  OR KE.ENDDATE IS NULL) "
            w_Sql = w_Sql & "  AND KE.KINMUDEPTCD = '" & m_KangoTaniCD & "' "
            w_Sql = w_Sql & "  ORDER BY KE.IDODATE, ST.STAFFMNGID"

            w_Rs = General.paDBRecordSetOpen(w_Sql)

            If w_Rs.RecordCount <= 0 Then
            Else
                With w_Rs
                    .MoveLast()
                    w_Count = .RecordCount
                    .MoveFirst()

                    w_Name_F = .Fields("STAFFNAME")
                    w_PostCD_F = .Fields("POSTCD")

                    ReDim p_lstChief(w_Count)
                    For w_Int = 1 To w_Count
                        p_lstChief(w_Int).Name = w_Name_F.Value
                        p_lstChief(w_Int).PostCD = w_PostCD_F.Value
                        .MoveNext()
                    Next w_Int
                End With
            End If

            w_Rs.Close()
            w_Rs = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Function Get_KakuteiF(ByVal p_FromLng As Integer, ByVal p_ToLng As Integer) As Boolean
        Try
            General.g_ErrorProc = "clsNSW0560P Get_KakuteiF"

            Dim w_Int As Short
            Dim w_Count As Short
            Dim w_Sql As String
            Dim w_PlanNo_F As ADODB.Field
            Dim w_Rs As ADODB.Recordset
            Dim w_Rs2 As ADODB.Recordset

            w_Sql = "SELECT PLANNO"
            w_Sql = w_Sql & " FROM NS_PLANCONTROL_F "
            w_Sql = w_Sql & " WHERE HOSPITALCD = '" & m_HospitalCD & "' "
            w_Sql = w_Sql & " AND PLANPERIODFROM <= " & p_ToLng
            w_Sql = w_Sql & " AND PLANPERIODTO > " & p_FromLng
            w_Sql = w_Sql & " ORDER BY PLANNO "

            w_Rs = General.paDBRecordSetOpen(w_Sql)

            If w_Rs.RecordCount <= 0 Then
                Get_KakuteiF = False
                w_Rs.Close()
                Exit Function
            Else
                With w_Rs
                    .MoveLast()
                    w_Count = .RecordCount
                    .MoveFirst()
                    w_PlanNo_F = w_Rs.Fields("PLANNO")
                    For w_Int = 1 To w_Count
                        w_Sql = "SELECT PLANNO "
                        w_Sql = w_Sql & " FROM NS_PLANDECISION_F "
                        w_Sql = w_Sql & " WHERE HOSPITALCD = '" & m_HospitalCD & "' "
                        w_Sql = w_Sql & " AND KINMUDEPTCD = '" & m_KangoTaniCD & "' "
                        w_Sql = w_Sql & " AND PLANNO = " & w_PlanNo_F.Value & ""

                        w_Rs2 = General.paDBRecordSetOpen(w_Sql)

                        If w_Rs2.RecordCount <= 0 Then
                            Get_KakuteiF = False
                            w_Rs.Close()
                            w_Rs2.Close()
                            Exit Function
                        End If
                        w_Rs2.Close()
                        .MoveNext()
                    Next w_Int
                End With
            End If

            w_Rs.Close()
            w_Rs = Nothing

            Get_KakuteiF = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    '基準日から年度を計算
    Private Function fncNendo(ByVal p_Date As Integer) As Integer
        Try
            General.g_ErrorProc = "clsNSW0560P fncNendo"

            Dim w_Year As Integer '年
            Dim w_MD As Integer '月日

            w_Year = General.paRoundDown(p_Date / 10000, 0)
            w_MD = p_Date Mod 10000

            If w_MD < 401 Then
                fncNendo = Integer.Parse(w_Year - 1)
            Else
                fncNendo = Integer.Parse(w_Year)
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function
End Class
