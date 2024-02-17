Public ones As Variant
Public teens As Variant
Public tens As Variant
Public thousands As Variant

Public fones As Variant
Public fteens As Variant
Public ftens As Variant
Public fhundreds As Variant
Public fthousands As Variant

Sub Initialize()
    ones = Array("zero", "one", "two", "tree", "four", "five", "six", "seven", "eight", "nine")
    teens = Array("ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen")
    tens = Array("zero", "ten", "twenty", "thirty", "fourty", "fifty", "sixty", "seventy", "eighty", "ninety")
    thousands = Array("", "thousand", "million", "billion", "trillion")
End Sub

Sub fInitialize()
    fones = Array("ÕÝÑ", "íß", "Ïæ", "Óå", "åÇÑ", "äÌ", "ÔÔ", "åÝÊ", "åÔÊ", "äå")
    fteens = Array("Ïå", "íÇÒÏå", "ÏæÇÒÏå", "ÓíÒÏå", "åÇÑÏå", "ÇäÒÏå", "ÔÇäÒÏå", "åÝÏå", "åÌÏå", "äæÒÏå")
    ftens = Array("ÕÝÑ", "Ïå", "ÈíÓÊ", "Óí", "åá", "äÌÇå", "ÔÕÊ", "åÝÊÇÏ", "åÔÊÇÏ", "äæÏ")
    fhundreds = Array("ÕÝÑ", "íßÕÏ", "ÏæíÓÊ", "ÓíÕÏ", "åÇÑÕÏ", "ÇäÕÏ", "ÔÔÕÏ", "åÝÊÕÏ", "åÔÊÕÏ", "äåÕÏ")
    fthousands = Array("", "åÒÇÑ", "ãíáíæä", "ãíáíÇÑÏ", "ÊÑíáíæä")
End Sub

Function say3(i As Integer) As String
    Dim retval As String
    retval = ""
    
    Initialize
    ' if i is not a posetive integer less than 1000 this function can not work
    If i >= 1000 Or i < 0 Then
        say3 = " *Invalid 3-dig number* "
        Exit Function
    End If
    
    If i = 0 Then say3 = "": Exit Function
    
    If i >= 100 Then
        retval = retval + ones(i \ 100) + " hundred"
        i = i Mod 100
        If i <> 0 Then retval = retval + " and "
    End If
    If i >= 10 Then
        If i < 20 Then
            retval = retval + teens(i Mod 10)
            i = 0
        Else
            retval = retval + tens(i \ 10)
            i = i Mod 10
            If i <> 0 Then retval = retval + " "
        End If
    End If
    If i > 0 Then
        retval = retval + ones(i)
    End If
    say3 = Trim(retval)
End Function

Function say(l As Long) As String
    Dim retval As String
    Dim dig3 As Integer

    If l < 0 Then
        say = "(minus) "
        l = l * -1
    ElseIf l > 0 Then
        say = ""
    Else
        say = "zero"
    End If

    dig3 = 0
    While l > 0
        retval = say3(l Mod 1000)
        If l Mod 1000 <> 0 Then
            retval = retval + " " + thousands(dig3) + " "
            If l \ 1000 <> 0 Then
                retval = "and " + retval
            End If
        End If
        dig3 = dig3 + 1
        l = l \ 1000
        say = retval + say
    Wend
    say = Trim(say)
End Function

Function fsay3(i As Integer) As String
    Dim retval As String
    retval = ""
    
    fInitialize
    ' if i is not a posetive integer less than 1000 this function can not work
    If i >= 1000 Or i < 0 Then
        fsay3 = " *ÚÏÏ Óå ÑÞãí ÇÔÊÈÇå* "
        Exit Function
    End If
    
    If i = 0 Then fsay3 = "": Exit Function
    
    If i >= 100 Then
        retval = retval + fhundreds(i \ 100)
        i = i Mod 100
        If i <> 0 Then retval = retval + " æ "
    End If
    If i >= 10 Then
        If i < 20 Then
            retval = retval + fteens(i Mod 10)
            i = 0
        Else
            retval = retval + ftens(i \ 10)
            i = i Mod 10
            If i <> 0 Then retval = retval + " æ "
        End If
    End If
    If i > 0 Then
        retval = retval + fones(i)
    End If
    fsay3 = Trim(retval)
End Function

Function fsay(l As Long) As String
    Dim retval As String
    Dim dig3 As Integer

    If l < 0 Then
        fsay = "(ãäÝí) "
        l = l * -1
    ElseIf l > 0 Then
        fsay = ""
    Else
        fsay = "ÕÝÑ"
    End If

    dig3 = 0
    While l > 0
        retval = fsay3(l Mod 1000)
        If l Mod 1000 <> 0 Then
            retval = retval + " " + fthousands(dig3) + " "
            If l \ 1000 <> 0 Then
                retval = "æ " + retval
            End If
        End If
        dig3 = dig3 + 1
        l = l \ 1000
        fsay = retval + fsay
    Wend
    fsay = Trim(fsay)
End Function

