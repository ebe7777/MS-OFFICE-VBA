Attribute VB_Name = "關於inch和dn互轉"
Private Function inchToDn(inch)
    If (inch = "0.125" Or inch = "0.125""" Or inch = "1/8" Or inch = "1/8""") Then
        inchToDn = 6
    ElseIf (inch = "0.25" Or inch = "0.25""" Or inch = "1/4" Or inch = "1/4""") Then
        inchToDn = 8
    ElseIf (inch = "0.375" Or inch = "0.375""" Or inch = "3/8" Or inch = "3/8""") Then
        inchToDn = 10
    ElseIf (inch = "0.5" Or inch = "0.5""" Or inch = "1/2" Or inch = "1/2""") Then
        inchToDn = 15
    ElseIf (inch = "0.75" Or inch = "0.75""" Or inch = "3/4" Or inch = "3/4""") Then
        inchToDn = 20
    ElseIf (inch = "1" Or inch = "1""") Then
        inchToDn = 25
    ElseIf (inch = "1.25" Or inch = "1.25""" Or inch = "1 1/4" Or inch = "1 1/4""" Or inch = "1-1/4" Or inch = "1-1/4""") Then
        inchToDn = 32
    ElseIf (inch = "1.5" Or inch = "1.5""" Or inch = "1 1/2" Or inch = "1 1/2""" Or inch = "1-1/2" Or inch = "1-1/2""") Then
        inchToDn = 40
    ElseIf (inch = "2" Or inch = "2""") Then
        inchToDn = 50
    ElseIf (inch = "2.5" Or inch = "2.5""" Or inch = "2 1/2" Or inch = "2 1/2""" Or inch = "2-1/2" Or inch = "2-1/2""") Then
        inchToDn = 65
    ElseIf (inch = "3" Or inch = "3""") Then
        inchToDn = 80
    ElseIf (inch = "4" Or inch = "4""") Then
        inchToDn = 100
    ElseIf (inch = "5" Or inch = "5""") Then
        inchToDn = 125
    ElseIf (inch = "6" Or inch = "6""") Then
        inchToDn = 150
    ElseIf (inch = "8" Or inch = "8""") Then
        inchToDn = 200
    ElseIf (inch = "10" Or inch = "10""") Then
        inchToDn = 250
    ElseIf (inch = "12" Or inch = "12""") Then
        inchToDn = 300
    ElseIf (inch = "14" Or inch = "14""") Then
        inchToDn = 350
    ElseIf (inch = "16" Or inch = "16""") Then
        inchToDn = 400
    ElseIf (inch = "18" Or inch = "18""") Then
        inchToDn = 450
    ElseIf (inch = "20" Or inch = "20""") Then
        inchToDn = 500
    ElseIf (inch = "22" Or inch = "22""") Then
        inchToDn = 550
    ElseIf (inch = "24" Or inch = "24""") Then
        inchToDn = 600
    ElseIf (inch = "26" Or inch = "26""") Then
        inchToDn = 650
    ElseIf (inch = "28" Or inch = "28""") Then
        inchToDn = 700
    ElseIf (inch = "30" Or inch = "30""") Then
        inchToDn = 750
    ElseIf (inch = "32" Or inch = "32""") Then
        inchToDn = 800
    ElseIf (inch = "34" Or inch = "34""") Then
        inchToDn = 850
    ElseIf (inch = "36" Or inch = "36""") Then
        inchToDn = 900
    ElseIf (inch = "38" Or inch = "38""") Then
        inchToDn = 950
    ElseIf (inch = "40" Or inch = "40""") Then
        inchToDn = 1000
    ElseIf (inch = "42" Or inch = "42""") Then
        inchToDn = 1050
    ElseIf (inch = "44" Or inch = "44""") Then
        inchToDn = 1100
    ElseIf (inch = "46" Or inch = "46""") Then
        inchToDn = 1150
    ElseIf (inch = "48" Or inch = "48""") Then
        inchToDn = 1200
    ElseIf (inch = "50" Or inch = "80""") Then
        inchToDn = 1250
    ElseIf (inch = "52" Or inch = "52""") Then
        inchToDn = 1300
    ElseIf (inch = "54" Or inch = "54""") Then
        inchToDn = 1350
    ElseIf (inch = "56" Or inch = "56""") Then
        inchToDn = 1400
    ElseIf (inch = "58" Or inch = "58""") Then
        inchToDn = 1450
    ElseIf (inch = "60" Or inch = "60""") Then
        inchToDn = 1500
    ElseIf (inch = "62" Or inch = "62""") Then
        inchToDn = 1550
    ElseIf (inch = "64" Or inch = "64""") Then
        inchToDn = 1600
    ElseIf (inch = "66" Or inch = "66""") Then
        inchToDn = 1650
    ElseIf (inch = "68" Or inch = "68""") Then
        inchToDn = 1700
    ElseIf (inch = "70" Or inch = "70""") Then
        inchToDn = 1750
    ElseIf (inch = "72" Or inch = "72""") Then
        inchToDn = 1800
    ElseIf (inch = "74" Or inch = "74""") Then
        inchToDn = 1850
    ElseIf (inch = "76" Or inch = "76""") Then
        inchToDn = 1900
    ElseIf (inch = "78" Or inch = "78""") Then
        inchToDn = 1950
    ElseIf (inch = "80" Or inch = "80""") Then
        inchToDn = 2000
    ElseIf (inch = "82" Or inch = "82""") Then
        inchToDn = 2050
    ElseIf (inch = "84" Or inch = "84""") Then
        inchToDn = 2100
    ElseIf (inch = "86" Or inch = "86""") Then
        inchToDn = 2150
    ElseIf (inch = "88" Or inch = "88""") Then
        inchToDn = 2200
    ElseIf (inch = "90" Or inch = "90""") Then
        inchToDn = 2250
    ElseIf (inch = "92" Or inch = "92""") Then
        inchToDn = 2300
    ElseIf (inch = "94" Or inch = "94""") Then
        inchToDn = 2350
    ElseIf (inch = "96" Or inch = "96""") Then
        inchToDn = 2400
    ElseIf (inch = "98" Or inch = "98""") Then
        inchToDn = 2450
    ElseIf (inch = "100" Or inch = "100""") Then
        inchToDn = 2500
    ElseIf (inch = "102" Or inch = "102""") Then
        inchToDn = 2550
    ElseIf (inch = "104" Or inch = "104""") Then
        inchToDn = 2600
    ElseIf (inch = "106" Or inch = "106""") Then
        inchToDn = 2650
    ElseIf (inch = "108" Or inch = "108""") Then
        inchToDn = 2700
    ElseIf (inch = "110" Or inch = "110""") Then
        inchToDn = 2750
    ElseIf (inch = "112" Or inch = "112""") Then
        inchToDn = 2800
    ElseIf (inch = "114" Or inch = "114""") Then
        inchToDn = 2850
    ElseIf (inch = "116" Or inch = "116""") Then
        inchToDn = 2900
    ElseIf (inch = "118" Or inch = "118""") Then
        inchToDn = 2950
    ElseIf (inch = "120" Or inch = "120""") Then
        inchToDn = 3000
    ElseIf (inch = "122" Or inch = "122""") Then
        inchToDn = 3050
    ElseIf (inch = "124" Or inch = "124""") Then
        inchToDn = 3100
    ElseIf (inch = "126" Or inch = "126""") Then
        inchToDn = 3150
    ElseIf (inch = "128" Or inch = "128""") Then
        inchToDn = 3200
    ElseIf (inch = "130" Or inch = "130""") Then
        inchToDn = 3250
    ElseIf (inch = "132" Or inch = "132""") Then
        inchToDn = 3300
    ElseIf (inch = "134" Or inch = "134""") Then
        inchToDn = 3350
    ElseIf (inch = "136" Or inch = "136""") Then
        inchToDn = 3400
    ElseIf (inch = "138" Or inch = "138""") Then
        inchToDn = 3450
    End If
End Function
Private Function dnToInch1(dn)
    If (dn = 6) Then
        dnToInch1 = 0.125
    ElseIf (dn = 8) Then
        dnToInch1 = 0.25
    ElseIf (dn = 10) Then
        dnToInch1 = 0.375
    ElseIf (dn = 15) Then
        dnToInch1 = 0.5
    ElseIf (dn = 20) Then
        dnToInch1 = 0.75
    ElseIf (dn = 25) Then
        dnToInch1 = 1
    ElseIf (dn = 32) Then
        dnToInch1 = 1.25
    ElseIf (dn = 40) Then
        dnToInch1 = 1.5
    ElseIf (dn = 50) Then
        dnToInch1 = 2
    ElseIf (dn = 65) Then
        dnToInch1 = 2.5
    ElseIf (dn = 80) Then
        dnToInch1 = 3
    ElseIf (dn = 100) Then
        dnToInch1 = 4
    ElseIf (dn = 125) Then
        dnToInch1 = 5
    ElseIf (dn = 150) Then
        dnToInch1 = 6
    ElseIf (dn = 200) Then
        dnToInch1 = 8
    ElseIf (dn = 250) Then
        dnToInch1 = 10
    ElseIf (dn = 300) Then
        dnToInch1 = 12
    ElseIf (dn = 350) Then
        dnToInch1 = 14
    ElseIf (dn = 400) Then
        dnToInch1 = 16
    ElseIf (dn = 450) Then
        dnToInch1 = 18
    ElseIf (dn = 500) Then
        dnToInch1 = 20
    ElseIf (dn = 550) Then
        dnToInch1 = 22
    ElseIf (dn = 600) Then
        dnToInch1 = 24
    ElseIf (dn = 650) Then
        dnToInch1 = 26
    ElseIf (dn = 700) Then
        dnToInch1 = 28
    ElseIf (dn = 750) Then
        dnToInch1 = 30
    ElseIf (dn = 800) Then
        dnToInch1 = 32
    ElseIf (dn = 850) Then
        dnToInch1 = 34
    ElseIf (dn = 900) Then
        dnToInch1 = 36
    ElseIf (dn = 950) Then
        dnToInch1 = 38
    ElseIf (dn = 1000) Then
        dnToInch1 = 40
    ElseIf (dn = 1050) Then
        dnToInch1 = 42
    ElseIf (dn = 1100) Then
        dnToInch1 = 44
    ElseIf (dn = 1200) Then
        dnToInch1 = 48
    ElseIf (dn = 1300) Then
        dnToInch1 = 52
    ElseIf (dn = 1400) Then
        dnToInch1 = 56
    ElseIf (dn = 1500) Then
        dnToInch1 = 60
    ElseIf (dn = 1600) Then
        dnToInch1 = 64
    ElseIf (dn = 1700) Then
        dnToInch1 = 68
    ElseIf (dn = 1800) Then
        dnToInch1 = 72
    ElseIf (dn = 1900) Then
        dnToInch1 = 76
    ElseIf (dn = 2000) Then
        dnToInch1 = 80
    ElseIf (dn = 2200) Then
        dnToInch1 = 88
    ElseIf (dn = 2300) Then
        dnToInch1 = 92
    ElseIf (dn = 2400) Then
        dnToInch1 = 96
    ElseIf (dn = 2500) Then
        dnToInch1 = 102
    ElseIf (dn = 2600) Then
        dnToInch1 = 104
    ElseIf (dn = 2700) Then
        dnToInch1 = 108
    ElseIf (dn = 2800) Then
        dnToInch1 = 112
    ElseIf (dn = 2900) Then
        dnToInch1 = 116
    ElseIf (dn = 3000) Then
        dnToInch1 = 120
    ElseIf (dn = 3100) Then
        dnToInch1 = 124
    ElseIf (dn = 3200) Then
        dnToInch1 = 128
    End If
End Function
