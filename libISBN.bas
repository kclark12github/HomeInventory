Attribute VB_Name = "libISBN"
'Algorithm taken from web site:
'http://www.isbn.org/standards/home/isbn/international/hyphenation-instructions.asp

'http://www.isbn.spk-berlin.de/
'An ISBN always consists of ten digits preceded by the letters ISBN.
'Note: In countries where the Latin alphabet is not used, an abbreviation in the characters of the local script may be used in addition to the Latin letters ISBN.
'
'The ten-digit number is divided into four parts of variable length, which must be separated clearly by hyphens or spaces5:
'
'ISBN 0 571 08989 5
'
'or
'
'isbn 90 - 70002 - 4 - 3
'
'Note: Experience suggests that hyphens are preferable to spaces.
'
'The number of digits in the first three parts of the ISBN (group identifier, publisher identifier, title identifier) varies. The number of digits in the group number and in the publisher identifier is determined by the quantity of titles planned to be produced by the publisher or publisher group. Publishers or publisher groups with large title outputs are represented by fewer digits.
'
'4.1. Group identifier
'The first part of the ISBN identifies a country, area or language area participating in the ISBN system. Some members form language areas (e.g. group number 3 = German language group) or regional units (e.g. South Pacific = group number 982). A group identifier may consist of up to 5 digits.
'
'Example: ISBN 90- ...
'
'All group identifiers are allocated by the International ISBN Agency in Berlin.
'
'4.2. Publisher identifier
'The second part of the ISBN identifies a particular publisher within a group. The publisher identifier usually indicates the exact identification of the publishing house and its address. If publishers exhaust their initial contingent of title numbers, they may be allocated an additional publisher identifier. The publisher identifier may comprise up to seven digits.
'
'Publisher identifiers are assigned by the ISBN group agency responsible for the management of the ISBN system within the country, area or language area where the publisher is officially based.
'
'Example: ISBN 90-70002- ...
'
'4.3. Title identifier
'The third part of the ISBN identifies a specific edition of a publication of a specific publisher. A title identifier may consist of up to six digits. As an ISBN must always have ten digits, blank digits are represented by leading zeros.
'
'Example: ISBN 90-70002-04- ...
'
'4.4. Check digit
'The check digit is the last digit of an ISBN. It is calculated on a modulus 11 with weights 10-2, using X in lieu of 10 where ten would occur as a check digit.
'
'This means that each of the first nine digits of the ISBN – excluding the check digit itself – is multiplied by a number ranging from 10 to 2 and that the resulting sum of the products, plus the check digit, must be divisible by 11 without a remainder.
'
'For example   ISBN 0-8436-1072-7:
'
'
'
'  Group
'Identifier Publisher
'Identifier Title
'Identifier Check
'digit
'ISBN  0 8 4 3 6   1 0 7 2 7
'Weight 10 9 8 7 6   5 4 3 2
'
'--------------------------------------------------------------------------------
'
'Products 0 + 72 + 32 + 21 + 36 + 5 + 0 + 21 + 4 + 7
'
'Total: 198
'
'As 198 can be divided by 11 without remainder 0-8436-1072-7 is a valid ISBN.
'7 is the valid check digit.
'4.5. Distribution of ranges
'The number of digits in each of the identifying parts 1, 2 and 3 is variable, although the total sum of digits contained in these parts is always 9. These nine digits, together with the check digit, make up the ten-digit ISBN.
'
'The number of digits in the group identifier will vary according to the output of books in a group. Thus, groups with an expected large output, will receive numbers of one or two digits and publishers with an expected large output will get numbers of two or three digits.
'
'For ease of reading, the four parts of the ISBN are divided by spaces or hyphens.
'
'The generation of hyphens at output by programming helps reduce work at input. It reduces the number of characters, eliminates manual checking of hyphenation, and insures accuracy of format in all ISBN listings and publications.
'
'The position of the hyphens is determined by the publisher identifier ranges established by each group agency in accordance with the book industry needs. The knowledge of the prefix ranges for each country or group of countries is necessary to develop the hyphenation output program.
'
'For example, the publisher identifier ranges of group number 0 in the English language group (Australia, English speaking Canada, Ireland, New Zealand, Puerto Rico, South Africa, Swaziland, United Kingdom, United States, and Zimbabwe) are as follows:
'
'00 - 19
'200 - 699
'7000 - 8499
' 85000 - 89999
'900000 - 949999
'9500000 - 9999999
'
'The following table is an example of the range distribution of publisher identifiers. Assuming a group identifier of one digit only, the publisher identifier ranges might be as shown in the left-hand column and the title identifiers as shown in the right-hand column.
'
  'Publisher Identifier
'
' Numbers available per publisher for
'Title identification
'
'
'--------------------------------------------------------------------------------
'
  '00 - 19
'200 - 699
'7000 - 8499
'85000 - 89999
'900000 - 949999
'9500000 - 9999999
'
'
' 1 000 000
'100 000
'10 000
'1 000
'100
'10
'
'
'
'
'Example: Group identifier "0"  If number ranges are between
' Insert hyphens after
'
'--------------------------------------------------------------------------------
'
  '00 - 19
'200 - 699
'7000 - 8499
'85000 - 89999
'900000 - 949999
'9500000 - 9999999  00 - 19
'20 - 69
'70 - 84
'85 - 89
'90 - 94
'95 - 99  1st   3rd   9th digit
'"    4th             "
'"    5th             "
'"    6th             "
'"    7th             "
'"    8th             "
'
'
'
'
'--------------------------------------------------------------------------------
'
'5 For purposes of data processing the 10-digit string is used without hyphens or spaces. Interpretation and human legible display is effectuated by means of the tables of group numbers and publisher identifier ranges.
'
'
'Group Identifier:        Country or Area:
'
'--------------------------------------------------------------------------------
'
'0 + 1 English speaking area:
  'Australia, Canada (E.), Gibraltar, Ireland, New Zealand, Puerto Rico, South Africa, Swaziland, UK, USA, Zimbabwe
'2 French speaking area:
  'France, Belgium (Fr. sp.), Canada (Fr. sp.), Luxembourg, Switzerland (Fr. sp.)
'3 German speaking area:
  'Austria, Germany, Switzerland (Germ. sp.)
'4 Japan
'5 Russian Federation:
  'Azerbaijan, Tajikistan,
'Turkmenistan, Uzbekistan,
'Armenia (and 99930),
'Belarus (and 985),
'Estonia (and 9949, 9985),
'Georgia (and 99928),
'Kazakhstan (and 9965),
'Kyrgyzstan (and 9967),
'Latvia (and 9984),
'Lithuania (and 9986),
'Moldova, Republic (and 9975),
'Ukraine (and 966)
'7 China , People 's Republic
'80 Czech Republic; Slovakia
'81 India (and 93)
'82 Norway
'83 Poland
'84 Spain
'85 Brazil
'86 Yugoslavia:
   'Bosnia and Herzegovina (and 9958),
'Croatia (and 953),
'Macedonia (and 9989),
'Slovenia (and 961)
'87 Denmark
'88 Italian speaking area:
  'Italy, Switzerland (It. sp.)
'89 Korea
'90 Netherlands
  'Netherlands,
'Belgium (Flemish)
'91 Sweden
'92 International Publishers (Unesco, EU);
'European Community Organizations
'93 India (and 81)
'950 Argentina
'951 Finland (and 952)
'952 Finland (and 951)
'953 Croatia
'954 Bulgaria
'955 Sri Lanka
'956 Chile
'957 Taiwan, China (and 986)
'958 Colombia
'959 Cuba
'960 Greece
'961 Slovenia
'962 Hong Kong (and 988)
'963 Hungary
'964 Iran
'965 Israel
'966 Ukraine (and 5)
'967 Malaysia (and 983)
'968 Mexico (and 970)
'969 Pakistan
'970 Mexico (and 968)
'971 Philippines
'972 Portugal (and 989)
'973 Romania
'974 Thailand
'975 Turkey
'976 Caribbean Community:
  'Antigua, Bahamas, Barbados, Belize, Cayman Islands, Dominica, Grenada, Guyana, Jamaica, Montserrat, St. Kitts-Nevis, St. Lucia, St. Vincent and the Grenadines, Trinidad and Tobago, Virgin Islands (Br)
'977 Egypt
'978 Nigeria
'979 Indonesia
'980 Venezuela
'981 Singapore (and 9971)
'982 South Pacific:
  'Cook Islands, Fiji, Kiribati, Marshall Islands, Nauru, Niue, Solomon Islands, Tokelau, Tonga, Tuvalu; Vanuatu, Western Samoa
'983 Malaysia (and 967)
'984 Bangladesh
'985 Belarus (and 5)
'986 Taiwan, China (and 957)
'987 Argentina
'988 Hongkong (and 962)
'989 Portugal (and 972)
'9948 United Arab Emirates
'9949 Estonia (and 5, 9985)
'9950 Palestine
'9951 Kosova
'9952 Azerbaijan
'9953 Lebanon
'9954 Morocco (and 9981)
'9955 Lithuania (and 5, 9986)
'9956 Cameroon
'9957 Jordan
'9958 Bosnia and Herzegovina
'9959 Libya
'9960 Saudi Arabia
'9961 Algeria
'9962 Panama
'9963 Cyprus
'9964 Ghana
'9965 Kazakhstan (and 5)
'9966 Kenya
'9967 Kyrgyzstan (and 5)
'9968 Costa Rica (and 9977)
'9970 Uganda
'9971 Singapore (and 981)
'9972 Peru
'9973 Tunisia
'9974 Uruguay
'9975 Moldova (and 5)
'9976 Tanzania (and 9987)
'9977 Costa Rica (and 9968)
'9978 Ecuador
'9979 Iceland
'9980 Papua New Guinea
'9982 Zambia
'9983 Gambia
'9984 Latvia (and 5)
'9985 Estonia (and 5, 9949)
'9986 Lithuania (and 5, 9955)
'9987 Tanzania (and 9976)
'9988 Ghana
'9989 Macedonia (and 86)
'99901 Bahrain
'99903 Mauritius
'99904 Netherlands Antilles
  'Aruba, Neth. Ant.
'99905 Bolivia
'99906 Kuwait
'99908 Malawi
'99909 Malta (and 99932)
'99910 Sierra Leone
'99911 Lesotho
'99912 Botswana
'99913 Andorra (and 99920)
'99914 Suriname
'99915 Maldives
'99916 Namibia
'99917 Brunei Darussalam
'99918 Faroe Islands
'99919 Benin
'99920 Andorra (and 99913)
'99921 Qatar
'99922 Guatemala (and 99939)
'99923 El Salvador
'99924 Nicaragua
'99925 Paraguay
'99926 Honduras
'99927 Albania
'99928 Georgia
'99929 Mongolia
'99930 Armenia (and 5)
'99931 Seychelles
'99932 Malta (and 99909)
'99933 Nepal
'99934 Dominican Republic
'99935 Haiti
'99936 Bhutan
'99937 Macau
'99938 Srpska
'99939 Guatemala (and 99922)
Option Explicit
Const isISBN = 0
Const isISSN = 1
Const isISMN = 2
Const isEAN = 3

Dim TypeLen(0 To 3) As Integer
Dim TypeName(0 To 3) As String
Dim mPubPref(0 To 9) As Integer
Dim chstr(0 To 299) As String
Private Sub Init()
    TypeLen(0) = 9
    TypeLen(1) = 7
    TypeLen(2) = 9
    TypeLen(3) = 12
    
    TypeName(0) = "ISBN"
    TypeName(1) = "ISSN"
    TypeName(2) = "ISMN"
    TypeName(3) = "EAN"

    mPubPref(0) = 3
    mPubPref(1) = 4
    mPubPref(2) = 4
    mPubPref(3) = 4
    mPubPref(4) = 5
    mPubPref(5) = 5
    mPubPref(6) = 5
    mPubPref(7) = 6
    mPubPref(8) = 6
    mPubPref(9) = 7

    'data obtained from http://www.isbn.spk-berlin.de/html/prefix/
    chstr(0) = "00002070859095"         'English
    chstr(2) = "00002070849095"         'French
    chstr(3) = "00002070859095"         'German
    chstr(4) = "00002070859095"         'Japan
    chstr(5) = "000020708590@@92939598" 'Russia
    chstr(7) = "000010508090"           'China
    chstr(10) = "000020708590"          'Czechoslovakia
    chstr(11) = "000020708590"          'India - also has 93 but undefined
    chstr(12) = "000020709099"          'Norway
    chstr(13) = "000020708590"          'Poland
    chstr(14) = "0000207085909597"      'Spain
    chstr(15) = "000020708590"          'Brazil
    chstr(16) = "000030708090"          'Yugoslavia
    chstr(17) = "000040708597"          'Denmark
    chstr(18) = "000020708590"          'Italian
    chstr(19) = "000025558595"          'Korea
    chstr(20) = "000020507080"          'Belgium, Netherlands
    chstr(21) = "002050708597"          'Sweden
    chstr(22) = "006080909599"          'Unesco
    chstr(50) = "0000509099"            'Argentina
    chstr(51) = "0020558995"            'Finland
    chstr(52) = "00002050@@89@@9599"    'Finland
    chstr(53) = "0010156096"            'Croatia
    chstr(54) = "000030809095"          'Bulgaria
    chstr(55) = "0020558095"            'Sri Lanka
    chstr(56) = "00002070"              'Chile
    chstr(57) = "0000448297"            'Taiwan
    chstr(58) = "0000608095"            'Colombia
    chstr(59) = "00002070"              'Cuba
    chstr(60) = "0000207085"            'Greece
    chstr(61) = "0000206090"            'Slovenia
    chstr(62) = "00002070858790"        'Hong Kong
    chstr(63) = "000020708590"          'Hungary
    chstr(64) = "0000305590"            'Iran
    chstr(65) = "0000207090"            'Israel
    chstr(66) = "0000507090"            'Ukraine
    chstr(67) = "00609099"              'Malaysia -- 999 makes 5 digits but this is now accounted for later
    chstr(68) = "00004050@@80"          'Mexico -- don't know what happens before 10 and after 899
    chstr(69) = "00204080"              'Pakistan
    chstr(70) = "0000609091"            'Mexico
    chstr(71) = "0000508591"            'Philippines
    chstr(72) = "0020558095"            'Portugal
    chstr(73) = "0020558095"            'Romania
    chstr(74) = "0000207085"            'Thailand
    chstr(75) = "0000306092"            'Turkey
    chstr(76) = "0040608095"            'Caribbean
    chstr(77) = "0000205070"            'Egypt
    chstr(78) = "00000020308090"        'Nigeria
    chstr(79) = "0020408095"            'Indonesia
    chstr(80) = "00002060"              'Venezuela
    chstr(81) = "00002030"              'Singapore
    chstr(82) = "000010@@70@@90"        'South Pacific
    chstr(83) = "000002204050809099"    'Malaysia
    chstr(84) = "0000408090"            'Bangladesh
    chstr(85) = "0000406090"            'Belarus
    chstr(87) = "0000509095"            'Argentina
    chstr(154) = "00204080"             'Morocco
    chstr(155) = "00004090"             'Lithuania
    chstr(156) = "00104090"             'Cameroun
    chstr(157) = "00004085"             'Jordan
    chstr(158) = "00105090"             'Bosnia
    chstr(160) = "00006090"             'Saudi Arabia
    chstr(161) = "00508095"             'Algeria
    chstr(162) = "00006085"             'Panama
    chstr(163) = "00305575"             'Cyprus
    chstr(164) = "007095"               'Ghana
    chstr(165) = "00004090"             'Kazakstan
    chstr(166) = "00008096"             'Kenya
    chstr(167) = "00004090"             'Kyrgyzstan
    chstr(168) = "00107097"             'Costa Rica
    chstr(170) = "00004090"             'Uganda
    chstr(171) = "00609099"             'Singapore
    chstr(172) = "0000@@@@10406090"     'Peru
    chstr(173) = "00107097"             'Tunisia
    chstr(174) = "00305575"             'Uruguay
    chstr(175) = "00509095"             'Moldova
    chstr(176) = "006090"               'Tanzania + 999 is 4 digits
    chstr(177) = "00009099"             'Costa Rica
    chstr(178) = "00009599"             'Ecuador
    chstr(179) = "00508090"             'Iceland
    chstr(180) = "004090"               'Papua New Guinea
    chstr(181) = "00001016208095"       'Morocco
    chstr(182) = "00008099"             'Zambia
    chstr(183) = "00809599"             'Gambia -- Don't know what happens before 80
    chstr(184) = "00005090"             'Latvia
    chstr(185) = "00508090"             'Estonia
    chstr(186) = "000040909497"         'Lithuania
    chstr(187) = "00004088"             'Tanzania
    chstr(188) = "00305575"             'Ghana
    chstr(189) = "00306095"             'Macedonia
    chstr(203) = "002090"               'Mauritius
    chstr(204) = "006090"               'Netherlands Antilles
    chstr(206) = "003060"               'Kuwait
    chstr(208) = "001090"               'Malawi
    chstr(209) = "004095"               'Malta
    chstr(210) = "003090"               'Sierra Leone
    chstr(211) = "000060"               'Lesotho
    chstr(212) = "006090"               'Botswana
    chstr(214) = "005090"               'Suriname
    chstr(215) = "005080"               'Maldives
    chstr(216) = "003070"               'Namibia
    chstr(217) = "003090"               'Brunei Darussalam
    chstr(218) = "004090"               'Faroes
    chstr(219) = "004090"               'Benin
    chstr(220) = "005090"               'Andorra
    chstr(221) = "002070"               'Qatar
    chstr(223) = "002080"               'El Salvador
    chstr(225) = "004080"               'Paraguay
    chstr(226) = "001060"               'Honduras
    chstr(227) = "003060"               'Albania
    chstr(228) = "001080"               'Georgia
    chstr(231) = "005080"               'Seychelles
    chstr(232) = "001060"               'Malta
End Sub
Public Function checkISBN(strISBN As String) As String
    Dim ISBN As String
    Dim iType As Integer
    Dim ckDig As String
    Dim cd As String
    Dim Message As String
    Dim EAN As String
    
    ISBN = getISBN(strISBN)
    iType = GetType(ISBN)
    If Valid(ISBN, iType) Then
        ckDig = GetCkDig(ISBN)
        Message = "This is an " & TypeName(iType) & ".  The check digit is "
        If Len(ISBN) > TypeLen(iType) Then
            cd = ckDig
            If Mid(ISBN, TypeLen(iType) + 1, 1) <> cd Then
                Message = Message & "incorrect: it should be " & ckDig & "."
            Else
                Message = Message & "correct."
            End If
        Else
            Message = Message & ckDig & "."
        End If
        Message = Message & vbCrLf
        Message = Message & "The full " & TypeName(iType) & " is " & FullNum(ISBN) & "." & vbCrLf
        If iType = isEAN And Mid(ISBN, 1, 3) = "978" Then Message = Message & "It is for a book with ISBN " & FullNum(Mid(ISBN, 4, 12)) & "."
        If iType = isEAN And Mid(ISBN, 1, 3) = "977" Then Message = Message & "It is for a serial publication with ISSN " & FullNum(Mid(ISBN, 4, 10)) & "."
        If iType = isEAN And Mid(ISBN, 1, 4) = "9790" Then Message = Message & "It is for a piece of music with ISMN " & FullNum("M" & Mid(ISBN, 5, 12)) & "."
    Else
        If Len(ISBN) > 13 And Mid(ISBN, 1, 3) = "977" And Valid(Mid(ISBN, 1, 13), isEAN) Then
            EAN = Mid(ISBN, 1, 13)
            If EAN = FullNum(EAN) Then
                Message = EAN & " is an EAN. The check digit is correct." & vbCrLf
                Message = Message & "It is for a serial publication with ISSN " & FullNum(Mid(ISBN, 4, 10)) & ", issue number " & Mid(ISBN, 14, Len(ISBN)) & "."
            Else
                Message = "You have typed in too much, too little, or letters instead of numbers."
            End If
        Else
            Message = "You have typed in too much, too little, or letters instead of numbers."
        End If
    End If
    checkISBN = Message
End Function
Private Function GetType(sNum As String) As Integer
    Dim x As Integer
    Dim l As Integer
    Dim c As String
    
    x = 0               'invalid default
    l = Len(sNum)
    c = Left(sNum, 1)
    
    If l = 9 Or l = 10 Then x = isISBN
    If l = 7 Or l = 8 Then x = isISSN
    If x = 0 And c = "M" Then x = isISMN
    If l = 12 Or l = 13 Then x = isEAN
    GetType = x
End Function
Private Function GetCkDig(sNum As String) As String
    Dim t As Integer
    Dim x As Integer
    Dim ckSum As Integer
    Dim ckDig As String
    
    t = GetType(sNum)
    If t < isISMN Then
        For x = 0 To TypeLen(t)
            ckSum = ckSum + (1 + TypeLen(t) - x) * Mid(sNum, x + 1, 1)
        Next x
        ckDig = CStr((1100 - ckSum) Mod 11)
        If ckDig = 10 Then ckDig = "X"
    Else
        If t = isISMN Then ckSum = 9
        For x = 3 - t To TypeLen(t)
            ckSum = ckSum + (3 - 2 * ((x + t) Mod 2)) * Mid(sNum, x + 1, 1)
        Next x
        ckDig = (1000 - ckSum) Mod 10
    End If
    GetCkDig = ckDig
End Function
Public Function FullNum(sNum As String) As String
    Dim strNum As String
    
    strNum = Mid(sNum, 1, TypeLen(GetType(sNum)) + 1 + GetCkDig(sNum))
    FullNum = Hyphenate(strNum)
End Function
Private Function Hyphenate(str As String) As String
    Dim ourstr As String
    Dim breaker As Integer
    Dim pref As String
    Dim p As String
    Dim mppl As Integer
    Dim i As Integer
    
    Select Case GetType(str)
        Case isISBN
            pref = Prefix(str)
            If pref <> 10 And chstr(shp(pref)) <> vbNullString Then
                p = chstr(shp(pref))
                mppl = 8 - Len(pref)
                i = Len(p) / 2
                While Mid(p, i * 2 - 2 + 1, i * 2) > Mid(str, Len(pref) + 1, Len(pref) + 2)
                    i = i - 1
                Wend
                If i = mppl + 1 And i = Len(p) / 2 Then i = mppl - 1                                                    'if it's only one over, it'll be one less
                If i > mppl And pref <> "962" Then i = i + mppl - Len(p) / 2                                            'They get bigger again, at least for Russia
                If i > mppl And (pref = "84" Or pref = "962" Or pref = "978" Or pref = "9986") Then i = 2 * mppl - i    'They get smaller in Hong Kong & Lithuania & Nigeria & Spain
                If Mid(str, 1, 6) = "967999" Or Mid(str, 1, 7) = "9976999" Then i = i + 1                               'Malaysia & Tanzania have a three-digit break
                breaker = i + Len(pref)
                ourstr = pref & "-" & Mid(str, Len(pref) + 1, breaker) & "-" & Mid(str, breaker + 1, 9) & "-" & Mid(str, 9 + 1, 1)
            Else
                ourstr = Mid(str, 1, 9) & "-" & Mid(str, 9 + 1, 1)
            End If
        Case isISSN
            ourstr = Mid(str, 1, 4) & "-" & Mid(str, 5, 8)
        Case isISMN
            breaker = 1 & mPubPref(Mid(str, 1, 1))
            ourstr = "M-" & Mid(str, 1, breaker) & "-" & Mid(str, breaker + 1, 9) & "-" & Mid(str, 9 + 1, 1)
        Case isEAN
            ourstr = str
    End Select
    Hyphenate = ourstr
End Function
Private Function Prefix(str As String) As Long
    Dim x As Long
    x = 10
    If Mid(str, 1, 1) < "8" Then x = Mid(str, 1, 1)
    If Mid(str, 1, 2) > "79" And Mid(str, 1, 2) < "94" Then x = Mid(str, 1, 2)
    If Mid(str, 1, 2) > "94" And Mid(str, 1, 2) < "99" Then x = Mid(str, 1, 3)
    If Mid(str, 1, 3) > "989" And Mid(str, 1, 3) < "999" Then x = Mid(str, 1, 4)
    If Mid(str, 1, 3) = "999" Then x = Mid(str, 1, 5)
    Prefix = x
End Function
Private Function shp(pref As String) As Long
    Dim x As Long
    If pref < 8 Then x = pref
    If pref > 7 And pref < 99 Then x = pref - 70        '80-93 -> 10-22
    If pref > 100 And pref < 999 Then x = pref - 900    '950-989 -> 50-89
    If pref > 1000 And pref < 9999 Then x = pref - 9800 '9900-9989 -> 100-189
    If pref > 10000 Then x = pref - 99700               '99900-99999 -> 200-299
    shp = x
End Function
Private Function Valid(sNum As String, iType As Integer) As Boolean
    Dim v As Boolean
    Dim x As Integer
    Dim c As String
    
    v = (iType <= isEAN)
    For x = 1 To TypeLen(iType)
        c = Mid(sNum, CLng(x) + 1, 1)
        If (x > 1 Or c <> "M" Or iType <> 2) And (c < "0" Or c > "9") Then v = False
    Next x
    Valid = v
End Function
Private Function getISBN(istring As String) As String
    Dim i As String
    Dim j As String
    Dim x As Integer
    
    For x = 1 To Len(istring)
        j = Mid(istring, CLng(x) + 1, 1)
        If j >= "0" And j <= "9" Then
            i = i & j
        ElseIf UCase(j) = "X" Then
            i = i & "X"
        ElseIf UCase(j) = "M" Then
            i = i & "M"
        End If
        Next x
    getISBN = i
End Function
Public Function FormatISBN(sISBN As String) As String
    Dim s As String
    Dim Group As String
    
    FormatISBN = sISBN
    s = Replace(sISBN, "-", vbNullString)
    If Len(s) <> 10 Then GoTo ExitSub
    
    Group = Left(s, 1)
    Select Case Group
        Case "0"
            Select Case Mid(s, 2, 2)
                Case "00" To "19"
                    FormatISBN = Format(s, "&-&&-&&&&&&-&")
                Case "20" To "69"
                    FormatISBN = Format(s, "&-&&&-&&&&&-&")
                Case "70" To "84"
                    FormatISBN = Format(s, "&-&&&&-&&&&-&")
                Case "85" To "89"
                    FormatISBN = Format(s, "&-&&&&&-&&&-&")
                Case "90" To "94"
                    FormatISBN = Format(s, "&-&&&&&&-&&-&")
                Case "95" To "99"
                    FormatISBN = Format(s, "&-&&&&&&&-&-&")
            End Select
        Case "1"
            Select Case Mid(s, 2, 2)
                Case "00" To "09"
                    FormatISBN = Format(s, "&-&&-&&&&&&-&")
                Case "10" To "39"
                    FormatISBN = Format(s, "&-&&&-&&&&&-&")
                Case "40" To "54"
                    FormatISBN = Format(s, "&-&&&&-&&&&-&")
                Case Else
                    Select Case Mid(s, 2, 4)
                        Case "5500" To "8697"
                            FormatISBN = Format(s, "&-&&&&&-&&&-&")
                        Case "8698" To "9989"
                            FormatISBN = Format(s, "&-&&&&&&-&&-&")
                        Case "9990" To "9999"
                            FormatISBN = Format(s, "&-&&&&&&&-&-&")
                    End Select
            End Select
    End Select

ExitSub:
    Exit Function
End Function
