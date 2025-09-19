My [ebeDictionary] is a substitute for Microsoft VBA's [Scripting.Dictionary]
Follow describe all the function of the class

Sub example()
Dim iDict1 As New ebeDictionary, iDict2 As New ebeDictionary
Dim iCount As Long
dim iValue as String
dim iBool as Boolean
dim iArray()
    'ebeDictionary contain 2 attribute "key", "value"
    '   [key] a Non-repeating data (like page number of a book)；data type is string
    '   [value] the value you want to look for by key(like chapter title on certain page of a book)；data type is string
    iDict1.Add "key1", "A"
    iDict1.Add "key2", "B"

    'copy data from iDict1 to iDict2
    iDict2.copy iDict1
    

    'get the total number of [key] (how many [key]s in the dictionary)
    iCount = iDict1.Count

    'get the [value] of a certain [key]
    iValue = iDict1.GetValue("key1")

    'check if any of [key] match the string you input
    '   true = find match, false = no match
    iBool = iDict1.Exists("key1")

    'get all the [value] of the dictionary,input them into a Array("n")
    '   "n" is the first number where data is inputted
    iArray = iDict1.GetValuesAsArray1D(1)

    'get all the [key] of the dictionary,input them into a Array("n")
    '   "n" is the first number where data is inputted
    iArray = iDict1.GetKeysAsArray1D(1)

    'get all the [key] and [value] of the dictionary,input them into a Array(2,"n")
    '   [1,N]key [2,N]Value
    '   [N,"n"]numbers of keys
    iArray = iDict1.GetKeysAndValuesAsArray2D

    'get the order of a  certain "key" in the dictionary
    '   e.g.
    '       dict key(1)= "a", key(2)= "b"
    '       the order of key "b" is 2
    iCount = iDict1.GetPageNo("key2")
 
    'get the [key] by the data "order" of a dictionary
    iValue = iDict1.GetKeyByPageNo(2)

    'get the [value] by the data "order"  of a dictionary
    iValue = iDict1.GetValueByPageNo(2)
    
    'sort the dictionary by key, you can choose to sort in ascend order or descend order
    '   ascend order = from small to big
    '   descend order = from big to samll
    ReDim iArray(3)
    iArray(3) = "C"
    iArray(1) = "A"
    iArray(2) = "B"
    iDict1.SortKey ("ascend")
  
    'get all the [key] as a string, each key is separated by "delemeter"
    iValue = iDict1.GetKeysAsString("^")
   
    'get all the [value] as a string, each value is separated by "delemeter"
    iValue = iDict1.GetValuesAsString("^")
  
    'delete the certain "key" in the dictionary
    iDict1.Remove ("key1")
 
    'Edit the [value] of a certain [key]
    iDict1.EditValue "key2", "C"
   
    'delete all data of the dictionary
    iDict1.RemoveAll


End Sub
