Attribute VB_Name = "����Chart���ާ@"
Sub ��_��Ʈw()
Dim myChartName As String
Dim myChart As Chart
    '�s�W
    myChartName = "Sales"
    Charts.Add after:=Sheets(Sheets.Count)
    ActiveChart.Name = myChartName
    Set myChart = Charts(myChartName)
    '����
    Charts(myChartName).Move after:=Sheets(Sheets.Count)
    'for each
    For Each iVar In Charts
        If (iVar.Name = "Sales") Then
            MsgBox "YES"
        End If
    Next
    
    '�s�褺�e
    '   https://learn.microsoft.com/zh-tw/office/vba/api/excel.chart(object)
    '   https://learn.microsoft.com/en-us/office/vba/api/excel.chart(object)
    '�]�w��ƽd��
    '   �H��u�Ϭ��ҡA��ƨC�ӼƭȬO�ѥ��ܥk�i�ܦb�Ϫ�W
    '       xlRows �N�d�򤺪���ƥHRow���������@�ըåѥ��ܥk�̧Ǩ��X�A�M��ѥ��ܥk�i�ܦb�Ϫ�W
    '           �|�N�d�򤺪���ƥѤW��U�A�N�L�k��ܦb�Ϫ�W������@�����������D
    '       xColumnss �N�d�򤺪���ƥHColumn���������@�ըåѤW�ܤU�̧Ǩ��X�A�M��ѥ��ܥk�i�ܦb�Ϫ�W
    '           �|�N�d�򤺪���ƥѥ���k�A�N�L�k��ܦb�Ϫ�W������@�����������D
    myChart.SetSourceData Source:=Sheets("TEST1").Range("D2010:BM2011"), PlotBy:=xlRows
    myChart.SetSourceData Source:=Sheets("TEST1").Range("H2010:H2011"), PlotBy:=xlColumns
    '����/�����b���j���D�AĴ�p "���"�A�ӫD�C����ƪ����D( "1/1","2/1"...)
    myChart.ChartWizard CategoryTitle:="�����b�j���D"
    myChart.ChartWizard ValueTitle:="�����b�j���D"
    '�]�w����ƽd��A�ĴX�շ�@���D(0:�S�����D�A�C�ճ��O��ơF2:��1~2�ժ����O���D�A��3�ն}�l�����)
    myChart.ChartWizard CategoryLabels:=0
    '�O�_�u��ܬݪ������x�s�� (false �~�|�N���ê��x�s�檺�Ȩq�X�A���׬OCategoryLabels��SeriesLabels)
    myChart.PlotVisibleOnly = False
    '����/�����y�Ъ��ȱ���
    myChart.Axes(xlValue).MaximumScale = 1
    
    myChart.ChartType = xlLine
    myChart.ClearToMatchStyle
    myChart.ChartStyle = 23
    
    
    '�R��
    Charts(myChartName).Delete
End Sub


