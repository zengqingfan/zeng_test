Module mytest
    'Public Declare Function XML_test Lib "test2.dll" () As Integer
    'Public Declare Function XML_get_record Lib "test2.dll" (ByRef filePath As String, ByRef tag As String, ByRef xml_struct As Object) As Integer
    'Public Declare Function XML_close Lib "test2.dll" () As Integer

    Public Declare Function XML_get_record_count Lib "xml5.dll" (ByVal filePath As String, ByVal tag As String, ByRef iCount As Long, ByVal sstr As xml_b) As Integer
    'Public Declare Function HelloWorld Lib "test3.dll" (ByVal aa As String) As Integer
    'Public Declare Function XML_get_record Lib "test3.dll" (ByVal filePath As String, ByVal tag As String) As Integer
    Public Declare Function XML_get_record Lib "xml5.dll" (ByRef xml_struct As xml_struct) As Integer
    Public Declare Function XML_close Lib "xml5.dll" () As Integer

    Structure xml_struct

        Dim itemName As String
        Dim Value As String
    End Structure


    Structure xml_b
        <VBFixedString(1024)> Public Name() As Char
        
    End Structure

    Sub main()

        Dim filpath As String
        'Dim xml_a As Object
        Dim xml_a As Object
        Dim i As Integer
        Dim iCount As Long

        Dim sstr As xml_b

        Dim aa As String
        Dim isDateFlg As Boolean

        Dim my_xml() As xml_struct

        'sstr = ""
        aa = "20080601000000"
        ReDim sstr()

        aa = Format(aa, "yyyy/MM/dd HH:mm:ss")
        'aa = "2008/06/01 00:00:00.000"

        isDateFlg = IsDate(aa)

        filpath = "E:\test\test_xml2.xml"
        'xml_a = New Object
        'HelloWorld("aaaa")
        'i = XML_get_record(filpath, "Root")

        i = XML_get_record_count(filpath, "Root", iCount, sstr)


    End Sub
End Module