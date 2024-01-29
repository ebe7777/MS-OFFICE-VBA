Attribute VB_Name = "關於日期時間"
Option Explicit
Sub 日期相關函數()
Dim iStr As String
    '資料是否為日期
    '   已知可接受 月/日/年, 年/月/日
    IsDate (iStr)

End Sub
Sub vba暫停一段時間後繼續執行()
'方式一-最低1秒
Application.Wait (Now + TimeValue("0:00:01"))
'方式二-可到毫秒
'the numerical value 1 = 1 day
'1/24 is one hour
'1/(24*60) is one minute
'so 1/(24*60*60*2) is 1/2 second
'60# 一定要加#字號否則會產生error
Application.Wait Now + 1 / (24 * 60 * 60# * 1)
End Sub

Public Function nowTime(ByVal dateFormat As String) As String
'取得現在時間
'   依照dateFormat回傳特定格式
'       格式1：直接取用系統值
'           "GETDATE"
'           "GETTIME"
'           "GETNOW"
'       格式2：自定組合
'           6碼，分別是
'           Y 年 /M 月 /D 日/H 小時/M 分鐘/S 秒
'           忽略，以0替代
Dim nowYear As Integer, nowMonth As Integer, nowDay As Integer
Dim nowHr As Integer, nowMin As Integer, nowSec As Integer
Dim strYear As String, strMonth As String, strDay As String, strHr As String, strMin As String, strSec As String
Dim iStr1 As String
Dim i As Long
    '先計算自定組合要顯示的值
    '   年月日
    nowYear = Year(Now)
    nowMonth = Month(Now)
    nowDay = Day(Now)
    '   時分秒
    nowHr = Hour(Now)
    nowMin = Minute(Now)
    nowSec = Second(Now)
    '   小於10補0
    strYear = CStr(nowYear)
    
    If (nowMonth < 10) Then
        strMonth = "0" & CStr(nowMonth)
    Else
        strMonth = CStr(nowMonth)
    End If
    
    If (nowDay < 10) Then
        strDay = "0" & CStr(nowDay)
    Else
        strDay = CStr(nowDay)
    End If
    
    If (nowHr < 10) Then
        strHr = "0" & CStr(nowHr)
    Else
        strHr = CStr(nowHr)
    End If
    
    If (nowMin < 10) Then
        strMin = "0" & CStr(nowMin)
    Else
        strMin = CStr(nowMin)
    End If
    
    If (nowSec < 10) Then
        strSec = "0" & CStr(nowSec)
    Else
        strSec = CStr(nowSec)
    End If
    '格式1：直接取用系統值
    Select Case UCase(dateFormat)
        Case "GETDATE"
            '得到日期2018/12/26
            nowTime = CStr(Date)
        Case "GETTIME"
            ' 現在時間
            nowTime = CStr(Time())
        Case "GETNOW"
            ' 現在日期與時間
            nowTime = CStr(Now())
        Case Else
        '格式2：自定組合
        For i = 1 To 6
            iStr1 = Mid(UCase(dateFormat), i, 1)
            If (iStr1 <> "0") Then
                Select Case i
                    Case 1
                        nowTime = nowTime & strYear
                    Case 2
                        nowTime = nowTime & strMonth
                    Case 3
                        nowTime = nowTime & strDay
                    Case 4
                        nowTime = nowTime & strHr
                    Case 5
                        nowTime = nowTime & strMin
                    Case 6
                        nowTime = nowTime & strSec
                End Select
            End If
        Next i
    End Select
'''====test====
''    MsgBox nowTime("GETdate")
''    MsgBox nowTime("GETtime")
''    MsgBox nowTime("GETnow")
''    MsgBox nowTime("YMDHMS")
''    MsgBox nowTime("0MDHMS")
''    MsgBox nowTime("000HMS")
''    MsgBox nowTime("YMD000")
'''============
End Function

Public Function isLeapYearOrNor_資料庫(nowYear As Long)
'閏年(leap year)有366天,平年(common year)有365天
'閏年2月有29天,平年2月有28天
'step1 如果年份能被 4 整除，請移至步驟 2。 否則請移至步驟 5。
'step2 如果年份能被 100 整除，請移至步驟 3。 否則請移至步驟 4。
'step3 如果年份能被 400 整除，請移至步驟 4。 否則請移至步驟 5。
'step4 該年份為閏年 (有 366 天)。
'step5 該年份不是閏年 (有 365 天)
    If (nowYear Mod 4 <> 0) Then
        isLeapYearOrNor = False
    ElseIf (nowYear Mod 100 <> 0) Then
        isLeapYearOrNor = True
    ElseIf (nowYear Mod 400 <> 0) Then
        isLeapYearOrNor = True
    Else
        isLeapYearOrNo = False
    End If
End Function
Public Function howManyDaysThisMonth_資料庫(nowMonth As Long, isLeapYearOrNo As Boolean)
'依照月設定可使用日
'   閏年leap year, 平年common year, 大月odd month (奇數月), 小月even month(偶數月)
    Select Case nowMonth
        '1個月有31天
        Case 1, 3, 5, 7, 8, 10, 12
            howManyDaysThisMonth = 31
        '1個月有30天
        Case 4, 6, 9, 11
            howManyDaysThisMonth = 30
        '特殊-2月
        '   閏年2月有29天，平年2月有28天
        
        Case 2
            If (isLeapYearOrNo = True) Then
                howManyDaysThisMonth = 29
            Else
                howManyDaysThisMonth = 28
            End If
    End Select
End Function

Sub 定時炸彈_資料庫()


Dim bombText As String
'======time bomb=====
Dim bombDate As Date, nowDate As Date
nowDate = Date
bombDate = DateValue("2018/2/10")

If (nowDate - bombDate > 0) Then
    bombText = " Unexpected Error - Error code: 0x80004005"
    MsgBox bombText, vbCritical
    Exit Sub
End If
'====================

End Sub

Sub 計算某日是該年的相關資訊()
'這天是禮拜幾   Weekday函數
'   得到結果 1==>禮拜日 ，7==>禮拜六
'   https://docs.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/weekday-function
MsgBox Weekday("2022/1/1")
'這天是這年第幾個禮拜
'   https://support.microsoft.com/zh-tw/office/weeknum-%E5%87%BD%E6%95%B8-e5c43a03-b4ab-426c-b411-b18c13c75340
'    MSGBOX WEEKNUM(DATE("2023.3.17"),[return_type]
'                                     以禮拜幾為換週的規定，不輸入就是禮拜日
'                                     11 = 禮拜一,17 = 禮拜天
'                                     1/1一定是該年的第一週，如果1/2是禮拜二，[return_type]也選12，那到1/2就會換週成該年第2週
MsgBox WorksheetFunction.WeekNum("2022/1/1", 11)
End Sub


Public Function convertWeekDayToStr_資料庫(weekDayNo As Long)
'算出這天是禮拜幾用Weekday函數
'   拜日(1)....拜六(7)
    weekDayNo = Weekday("2022/1/1")
    Select Case weekDayNo
        Case 1
            convertWeekDayToStr = "週日"
        Case 2
            convertWeekDayToStr = "週一"
        Case 3
            convertWeekDayToStr = "週二"
        Case 4
            convertWeekDayToStr = "週三"
        Case 5
            convertWeekDayToStr = "週四"
        Case 6
            convertWeekDayToStr = "週五"
        Case 7
            convertWeekDayToStr = "週六"
    End Select
End Function


Sub 兩個時間中間差別多久()
Dim nowTime1 As Date, nowTime2 As Date, nowTime3 As Long
' "s" 時間差距以秒顯示
' nowTime1 一開始的時間
nowTime1 = Now()
    '.....do something
' nowTime2 結束的時間
nowTime2 = Now()
nowTime3 = DateDiff("s", nowTime1, nowTime2)

'see
'https://docs.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/datediff-function

'時間間隔引數interval 設定:
'yyyy 年
'q 季
'm 月份
'Y 一年中的一天
'd 日期
'w Weekday
'ww 週
'h 時
'n 分鐘
's 秒

End Sub

Public Function correctDateWritingInputToCustom_資料庫(dateCheckRg As Range, ByVal iYear As Long, ByVal iMonth As Long, ByVal iDay As Long)
'重新將輸入的日期排版為[年][月][日]，中間(依照設定)加上分隔符號
'(依照設定)重新將輸入的[年]排版為 西元/民國 格式
'此功能只檢查一個存儲格，故輸入的range必需為為 1 個儲存格；多個儲存格需多次呼喚此程式
Dim iStr As String
    If (iHaveErr = False) Then
        If (dateYearFormat = "AC" And iYear < 1911) Then
            iYear = iYear + 1911
        ElseIf (dateYearFormat = "ROC" And iYear >= 1911) Then
            iYear = iYear - 1911
        End If
        iStr = iYear & dateDeliSymbol & iMonth & dateDeliSymbol & iDay
        skipThis = True
        dateCheckRg.Value = iStr
        skipThis = False
    End If
End Function
Public Function correctDateWritingSysToCustom_資料庫(ByVal myDate As Date)
Dim iYear As Long, iMonth As Long, iDay As Long
'重新將系統的date值排版為[年][月][日]，中間(依照設定)加上分隔符號
'(依照設定)重新將輸入的[年]排版為 西元/民國 格式
'此功能只檢查一個存儲格，故輸入的range必需為為 1 個儲存格；多個儲存格需多次呼喚此程式
    iYear = Year(myDate)
    iMonth = Month(myDate)
    iDay = Day(myDate)
    If (dateYearFormat = "AC" And iYear < 1911) Then
        iYear = iYear + 1911
    ElseIf (dateYearFormat = "ROC" And iYear >= 1911) Then
        iYear = iYear - 1911
    End If
    correctDateWritingSysToCustom = iYear & dateDeliSymbol & iMonth & dateDeliSymbol & iDay
End Function
Private Sub 以目前日期時間新增資料夾_newFolder(saveFolderPath As String)

Dim nowYear As Integer, nowMonth As Integer, nowDay As Integer
Dim nowHr As Integer, nowMin As Integer, nowSec As Integer
Dim nowYearString As String, nowMonthString As String, nowDayString As String
Dim nowHrString As String, nowMinString As String, nowSecString As String
Dim saveFolderName As String
Dim sysSN As String



sysSN = "SYSTEM"

'定義資料夾名稱
    '抓取年月日
    nowYear = Year(Now)
    nowYearString = CStr(nowYear)
    
    nowMonth = Month(Now)
    nowMonthString = CStr(nowMonth)
    If (Len(nowMonthString) = 1) Then
        nowMonthString = "0" & nowMonthString
    End If
    
    nowDay = Day(Now)
    nowDayString = CStr(nowDay)
    If (Len(nowDayString) = 1) Then
        nowDayString = "0" & nowDayString
    End If
    
    nowHr = Hour(Now)
    nowHrString = CStr(nowHr)
    If (Len(nowHrString) = 1) Then
        nowHrString = "0" & nowHrString
    End If
    
    nowMin = Minute(Now)
    nowMinString = CStr(nowMin)
    If (Len(nowMinString) = 1) Then
        nowMinString = "0" & nowMinString
    End If
    
    nowSec = Second(Now)
    nowSecString = CStr(nowSec)
    If (Len(nowSecString) = 1) Then
        nowSecString = "0" & nowSecString
    End If
    '定義名稱
    saveFolderName = "修改後的圖_" & nowYearString & nowMonthString & nowDayString & "_" & nowHrString & nowMinString & nowSecString
'新增資料夾
    saveFolderPath = Sheets(sysSN).Cells(2, 2) & "\" & saveFolderName
    MkDir saveFolderPath
End Sub
