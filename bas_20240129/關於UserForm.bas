Attribute VB_Name = "����UserForm"
Sub �q�Ϊ���()
'combobox�Blistbox�p�G�ק�᪺�ȩM�ק�e�ۦP�A���|Ĳ�o����_Click/_Change...��sub
'   ==>���ܦ���L�ȦA�ܦ��ק�᪺��
'   �ߤ@��b�S�ܰʤ]����檺sub�O DropButtonClick

'�\�h����(combobox�Blistbox�Bscrollbar)�b���g��L�\��ɥi��|���o�Ǫ����Ĳ�o_Click/_Change...��sub�A���i���U����Ĳ�o
'   ==>�b�o��sub���w���]�w�n�}�� (if true do something,if false do nothing)
End Sub
Sub �}��form()
    UserForm1.Show False
    '�p�GuserForm�����e(label��caption��)�S��ܤ��e
    DoEvents
End Sub
Sub ����form()
    UserForm1.Hide
End Sub
Sub ����form()
    '��form�i��ݭn�b�h��excel�ɤ��}�}������(�@��excel�s��{���A��Lexcel���}�ϥ�)�A�p�G�u��hide�A��A��show�ɷ|active hide��form��excel��
    '  >�bTerminate�W�� unload
    Private Sub UserForm_Terminate()
'        Unload UserForm1
'    End Sub
    Unload UserForm1
End Sub

Public Function ��Ʈw_IsFormInitialized(FormName As String) As Boolean
'���լY�W�٪�UserForm�O�_�w�gInitialized
    Dim myForm As Variant
    For Each myForm In UserForms
        If myForm.Name = FormName Then
            IsFormInitialized = True
            Exit Function
        End If
    Next
End Function

Sub ��sform()
'���Ǳ��p�Uform�|�o�ͤ����`�����p�A�w���H�U�n�Φ��B�z
'   (1)��image�Qclick��A���s�]�mpicture���image�W�ɡAimage�W���|�O�¹Ϥ���(���M��ڦ�����)
    Me.Repaint
End Sub
Sub FormActivate���q�γ]�w()
    '�NForm�\�bexcel�Ҧb�e������
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
Sub Form�̭����ݩʳ]�w����()
    
    'picture-�Ϥ��]�w
    myObj.Picture = LoadPicture("c:\123.png")
    '   �Ϥ��O�_���]�w
    If (Not Image1.Picture Is Nothing) Then
        'do something
    End If
    'foreColor-�r���C��]�w,bas��Ƨ�����RGB��
    myObj.ForeColor = RGB(204, 204, 204)
    
End Sub
Sub ����()
'�ˬdtextbox�O�_��Jinteger
'https://stackoverflow.com/questions/26138833/making-vba-form-textbox-accept-numbers-only-including-and

'�I�s��L�u�@�� Private Sub Worksheet_Change(ByVal Target As Range)
Dim myWS As Worksheet
Dim myShtCodeName As String 'myShtCodeName:�u�@��CodeName�ݩʭ�
Dim myRg As Range   '�u�@����ܮɡA���b�ק諸�d��
Set myWS = Sheet("�Y�i�u�@��")
Set myRg = myWS.Range("A1")
    myShtCodeName = myWS.CodeName
    Application.Run myShtCodeName & ".Worksheet_Change", myRg

'�I�s��Lform��commandbutton_click
'   if 'the module' is the same form module as commandbutton1_click then yes you can.
'   If it is a different module you need to replace private with public for the commandbutton1_click sub and the form itself needs to be open.
'   And to call it you would use
'   �ѦҸ��
'       https://www.access-programmers.co.uk/forums/threads/calling-commandbutton_click-event-from-a-module.275200/

Form_whatevermyformiscalled.CommandButton1_Click
End Sub
Sub ���form�W������()
'�������w
Dim myLB As MSForms.ListBox
    Set myLB = Form03Inquiry.ListBox010101
'�ϥΪ���W��
Dim myForm As UserForm, TextBoxName As String, myObj As Object
Dim i As Integer

    Set myForm = UserForm1
    TextBoxName = "TextBox1"
    
    For i = 0 To myForm.Controls.Count
        If (myForm.Controls.Item(i).Name = comboBoxName) Then
            Set myObj = myForm.Controls.Item(i)
            Exit For
        End If
    Next i
End Sub
Function �NForm������a�Jfunction���ܼƤ�(ByRef TextBoxName As String)
    With Me.Controls(TextBoxName)
            .BackStyle = fmBackStyleTransparent
    End With
End Function
Sub �]�wtextBox����()

Dim myForm As UserForm, TextBoxName As String, myObj As Object
Dim i As Integer

Set myForm = UserForm1
TextBoxName = "TextBox1"

    For i = 0 To myForm.Controls.Count
        If (myForm.Controls.Item(i).Name = comboBoxName) Then
            Set myObj = myForm.Controls.Item(i)
            Exit For
        End If
    Next i
    myObj.Value = "123"
    
    '�n����ܭȴ���n�]�w MultiLine �� True.
End Sub
Sub �]�wTextBox�����()
    'textbox�i�H�ާ@��
    TextBox1.Enabled = True
    TextBox1.BackStyle = fmBackStyleOpaque
    'textbox����ާ@��
    TextBox202.Enabled = False
    TextBox202.BackStyle = fmBackStyleTransparent
End Sub
Sub Scrollbar�]�w���e()
    '�]�wscrollbar
    ScrollBar1.min = 1
    ScrollBar1.max = i
    ScrollBar1.Value = 1
    '�p�G�����ᦳimage�n�����A�h�C�������ᶷ���榹
    Me.Repaint
End Sub
Sub multipage�]�w���e()
    'activate page (0~n)
    MultiPage1.Value = 0
    '�]�w�Ypage���ݩ�
    MultiPage1.Pages(1).Enabled = False
    

End Sub
Sub ��Ʈw_MultiPage1_Change()
'�U��J�������ɡA�e���[��check mark
Dim i As Long
    i = MultiPage1.Value
    '���]����l��
    MultiPage1.Pages(0).Caption = "�D��"
    MultiPage1.Pages(1).Caption = "�D��"
    '�ثe�����page�W�٭ק�
    Select Case i
        Case 0
            MultiPage1.Pages(0).Caption = ChrW(&H2611) & "�D��"
        Case 1
            MultiPage1.Pages(1).Caption = ChrW(&H2611) & "�D��"
    End Select
'�藍�P���Ү�form�����סB�e�ק���
'   �bForm�� Private Sub UserForm_Initialize() �]�n���ۦP���]�w
Dim myHei As Long, myWidth As Long
    Select Case MultiPage1.Value
        Case 0
            myHei = 92.5
            myWidth = 161
        Case 1
            myHei = 127.5
            myWidth = 183
    End Select
    Me.Height = myHei
    Me.Width = myWidth
End Sub
Sub �]�wcomboBox���U�Ԧ��M���()
'�p�G�n���ϥΪ̵L�k��ʭק�combobox���e�A����ݩ�[Style]�A�令 [2-fmStyleDropDownList]
'ListIndex�q0�}�l

    '��ʶ�J
    ComboBox1.AddItem "�|��"
    ComboBox1.AddItem "���q"
    ComboBox1.ListIndex = 0
    '��array��J-�������w�S�w��combobox
    ComboBox1.List = iArray
    '���o�����Combobox����
    myVal = ComboBox1.Value
    '��array��J-form�W���\�hcombobox-->���S�w��
Dim myForm As UserForm, comboBoxName As String, myObj As Object
Dim myArray()
Dim i As Integer

Set myForm = UserForm1
comboBoxName = "ComboBox1"
    
    ReDim myArray(2)
    For i = 1 To 3
        myArray(i - 1) = i
    Next i
    
    For i = 0 To myForm.Controls.Count
        If (myForm.Controls.Item(i).Name = comboBoxName) Then
            Set myObj = myForm.Controls.Item(i)
            Exit For
        End If
    Next i
    myObj.List = myArray
    '�]�w��l��
    ComboBox1.ListIndex = 0

End Sub
Sub listBox�]�w���e()
    '���olistboxt�ثeselect���󪺭�
    '   �p�G�Ĥ@���}��form�B�S���H�u�I��ListBox(�HListBox.ListIndex���N�H�u)
    '       �|�ϱo.value�줣��ȡF�ϥ�.SetFocus�i�ѨM�����D
    listBoxObj.SetFocus
    myVal = listBoxObj.Value
    '�]�wcolumns�ƥءA�S�]�w�Hlistbox���󪺭�l��
    listBoxObj.ColumnCount = 10
    '���wcolumn���e��
    '   �U�O���w�A�p�L���w�̲Τ@���t�γ]�m
    listBoxObj.ColumnWidths = "50,50,50"
    
    '�W�K���
    '   �覡1(����ϥ�ColumnHeads)�A�ϥ�.list�A������̦h10 columns
    listBoxObj.AddItem
    listBoxObj.List(i - 1, ii) = dataArray(i, ii)
    '       �M��listbox�覡
    listBoxObj.Clear
    '   �覡2(����ϥ�ColumnHeads)�A�ϥ�array�A�Scolumn����
    '       �ϥΪ�array�ݬ�2���}�C [�C,��] ��J��
    '           �C�@�����_�l�Ȭ�0
    '       �檺�ƶq�ݦۦ�]�w�A�_�h�u���1��(�w�]��)
    'listBoxObj.ColumnCount = UBound(dataArray, 2)
    '           �ѩ���e���n����A��Ʀp���h���ĳ�h�]�X��ListBox���O��U�檺���
    '       �n���������ʶb�X�{�A�ݭק��ݩ� ColumnWidths �A�令�@�Ӥj���ݩ� Width ����
    listBoxObj.List = dataArray

    '   �覡3(�i�H�ϥ�ColumnHeads)�A�ϥΤu�@������ơA�Scolumn����
    '       Adding Column Headers
    '       You can only display column headers when you use the RowSource property, not when you use an array or add items individually.
    '       To display column headers set the ColumnHeads property to True.
    '       Do not include the column headings on the worksheet in the range defined for RowSource.
    '       The row directly above the first row of the RowSource will be automatically used.
    '       !!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '       �`�N!!��ListBox��rawsource���]�w�ɡA�R���P��J��ƨ�rawsource���u�@��ɷ|��LIST���e���_��s�ө�ֵ{������t��
    '           �ҥH��n�R���ηs�W��ƨ�Ӥu�@��ɡA�����M��RAWSOURCE�]�w
    '       !!!!!!!!!!!!!!!!!!!!!!!!!!!!
    listBoxObj.ColumnHeads = True
    '       ���w�d�򪺤W�@�C<range("A1:E1")>�|�۰ʧ������ColumnHeads
    '       *�᭱�n�ϥ�(External:=True)�A�_�h�u�|���active�u�@����range��T
    listBoxObj.RowSource = Sheets("sysListDataTempSht").Range("A2:E15").address(External:=True)
    '       �M��listbox�覡
    listBoxObj.RowSource = vbNullString
    
    '�P�w�ثelistbox��w���O�ĴX����� (�Ĥ@����ƬO0)
    For i = 0 To ListBox1.ListCount - 1
        If (ListBox1.Selected(i) = True) Then
            listboxSelectedRow = i
        End If
    Next i
    
    '�NListBox��ܬ��S����������ƪ����A deselect all
    '�`�N! �p�G�O�b��ListBox���{�Ǥ��A�p ListBox1_Click�A�h����k���|���ĪG
    ListBox1.Selected(ListBox1.ListIndex) = False
    
    '���o�ثe���������
    '   ListBox�O�q0�}�l
    ListBox1.List (0)
End Function
Private Function putDict1DimArryIntoListBoxArray_��Ʈw(dictArray(), ListBoxArray())
'�NebeDictionary��getKeys/getValues���ȵ�ListBox��
'   �@���}�C�O�q1�}�l�A��ListBox�ݨD���O�q0�}�l�F���\��Ndict����ư}�C�ରListBox��
Dim i As Long
    ReDim ListBoxArray(UBound(dictArray, 1) - 1)
    For i = 1 To UBound(dictArray, 1)
        ListBoxArray(i - 1) = dictArray(i)
    Next i
End Function
Private Function listBoxSeleAllorNot_��Ʈw(myListBox As MSForms.ListBox, seleAllorNot As Boolean)
'�N���w��ListBox�����������(True)/������(False)
    For i = 0 To myListBox.ListCount - 1
        myListBox.Selected(i) = seleAllorNot
    Next i
End Function

Sub optionButton������()
'�P�@�ӵe������optionButton�|�۰ʱ���u���A�諸��@��
'   ���P��MultiPage�A����optionButton�|�U�ۦۦ��@��

'�P�_�{�b�O����optionButton�Q��� ���]�����,�W�٬O optionButton1/optionButton2
'   �Q������ȷ|�Otrue
OptionButton1.Value = True
OptionButton2.Value = False

End Sub

