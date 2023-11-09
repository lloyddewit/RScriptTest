Imports System.Collections.Specialized
Imports Newtonsoft.Json.Linq
Imports System.ComponentModel
Imports System.Data.Common
Imports System.Reflection.Emit
Imports System.Text.RegularExpressions
Imports Xunit

Public Class clsRScriptTestUnit
    <Fact>
    Sub TestGetLstLexemes()
        'TODO remove this comment, it's just to test GitHub push works correctly

        Dim clsRScript As RScript.clsRScript = New RScript.clsRScript("")

        'test lexeme list - identifiers and standard operators
        Dim lstExpected As New List(Of String)(New String() {
                "a", "::", "b", ":::", "ba", "$", "c", "@", "d", "^", "e", ":", "ea", "%%", "f", "%/%", "g", "%*%", "h",
                "%o%", "i", "%x%", "j", "%in%", "k", "/", "l", "*", "m", "+", "n", "-", "o",
                "<", "p", ">", "q", "<=", "r", ">=", "s", "==", "t", "!=", "u", "!", "v", "&",
                "wx", "&&", "y", "|", "z", "||", "a2", "~", "2b", "->", "c0a", "->>", "d0123456789a", "<-", "1234567890", "<<-", "e0a1b2", "=", "345f6789"})
        Dim lstActual As List(Of String) = clsRScript.GetLstLexemes(
                "a::b:::ba$c@d^e:ea%%f%/%g%*%h%o%i%x%j%in%k/l*m+n-o<p>q<=r>=s==t!=u!v&wx&&y|z||a2~2b->c0a->>d0123456789a<-1234567890<<-e0a1b2=345f6789")
        Assert.Equal(lstExpected, lstActual)

        'test lexeme list - separators, brackets, line feeds, user-defined operators and variable names with '.' and '_'
        lstExpected = New List(Of String)(New String() {",", "ae", ";", "af", vbCr,
                "ag", vbLf, "(", "ah", ")", vbCrLf, "ai", "{", "aj", "}", "ak", "[", "al", "]",
                "al", "[[", "am", "]]", "_ao", "%>%", "|>", ".ap", "%aq%", ".ar_2", "%asat%",
                "au_av.awax"})
        lstActual = clsRScript.GetLstLexemes(
                ",ae;af" & vbCr & "ag" & vbLf & "(ah)" & vbCrLf & "ai{aj}ak[al]al[[am]]_ao%>%|>.ap" &
                "%aq%.ar_2%asat%au_av.awax")
        Assert.Equal(lstExpected, lstActual)

        'test lexeme list - spaces
        lstExpected = New List(Of String)(New String() {
                " ", "+", "ay", "-", " ", "az", "  ", "::", "ba", "   ", "%*%", "   ", "bb", "   ",
                "<<-", "    ", "bc", " ", vbCr, "  ", "bd", "   ", vbLf, "    ", "be", "   ",
                vbCrLf, "  ", "bf", " "})
        lstActual = clsRScript.GetLstLexemes(" +ay- az  ::ba   %*%   bb   <<-    bc " & vbCr &
                "  bd   " & vbLf & "    be   " & vbCrLf & "  bf ")
        Assert.Equal(lstExpected, lstActual)

        'test lexeme list - string literals
        lstExpected = New List(Of String)(New String() {
                """a""", "+", """bf""", "%%", """bga""", "%/%", """bgba""", "%in%", """bgbaa""",
                ">=", """~!@#$%^&*()_[] {} \|;:',./<>? """, ",", """ bgbaaa""", vbCr, """bh""",
                vbLf, """bi""", vbCrLf, """bj""", "{", """bk""", "[[", """bl""", "%>%", """bm""",
                "%aq%", """bn""", " ", "+", """bn""", "-", " ", """bo""", "  ", "::",
                """bq""", "   ", "<<-", "    ", """br""", " ", vbCr, "  ", """bs""", "   ", vbLf,
                "    ", """bt""", "   ", vbCrLf, "  ", """bu""", " "})
        lstActual = clsRScript.GetLstLexemes("""a""+""bf""%%""bga""%/%""bgba""%in%""bgbaa"">=" &
                """~!@#$%^&*()_[] {} \|;:',./<>? "","" bgbaaa""" & vbCr & """bh""" & vbLf &
                """bi""" & vbCrLf & """bj""{""bk""[[""bl""%>%""bm""%aq%""bn"" +""bn""- ""bo""  ::" &
                """bq""   <<-    ""br"" " & vbCr & "  ""bs""   " & vbLf & "    ""bt""   " & vbCrLf &
                "  ""bu"" ")
        Assert.Equal(lstExpected, lstActual)

        'test lexeme list - comments
        lstExpected = New List(Of String)(New String() {
                "#", vbLf, "c", "#", vbLf, "ca", "#", vbLf, "+", "#", vbLf, "%/%", "#", vbLf,
                "%in%", "#", vbLf, ">=", "#~!@#$%^&*()_[]{}\|;:',./<>?#", vbLf, " ", "#", vbLf,
                "  ", "#~!@#$%^&*()_[] {} \|;:',./<>?", vbLf, "#cb", vbLf, "#cba", vbLf,
                "# "","" cbaa ", vbLf, "#", vbCr, "#cc", vbCr, "#cca", vbCrLf, "# ccaa ", vbCrLf})
        lstActual = clsRScript.GetLstLexemes(
                "#" & vbLf & "c#" & vbLf & "ca#" & vbLf & "+#" & vbLf & "%/%#" & vbLf & "%in%#" &
                vbLf & ">=#~!@#$%^&*()_[]{}\|;:',./<>?#" & vbLf & " #" & vbLf &
                "  #~!@#$%^&*()_[] {} \|;:',./<>?" & vbLf & "#cb" & vbLf & "#cba" & vbLf &
                "# "","" cbaa " & vbLf & "#" & vbCr & "#cc" & vbCr &
                "#cca" & vbCrLf & "# ccaa " & vbCrLf)
        Assert.Equal(lstExpected, lstActual)
    End Sub

    <Fact>
    Sub TestGetLstTokens()
        Dim clsRScript As RScript.clsRScript = New RScript.clsRScript("")

        'test token list - RSyntacticName
        Dim strInput As String = "._+.1+.a+a+ba+baa+a_b+c12+1234567890+2.3+1e6+" &
                "abcdefghijklmnopqrstuvwxyz+`a`+`a b`+`[[`+`d,ae;af`+`(ah)`+`ai{aj}`+" &
                "`~!@#$%^&*()_[] {} \|;:',./<>?`+`%%a_2ab%`+`%ac%`+`[[""b""]]n[[[o][p]]]`+" &
                "`if`+`else`+`while`+`repeat`+`for`+`in`+`function`+`return`+`else`+`next`+`break`"
        Dim lstInput As List(Of String) = clsRScript.GetLstLexemes(strInput)
        Dim strExpected As String = "._(RSyntacticName), +(ROperatorBinary), .1(RSyntacticName), " &
                "+(ROperatorBinary), .a(RSyntacticName), +(ROperatorBinary), a(RSyntacticName), " &
                "+(ROperatorBinary), ba(RSyntacticName), +(ROperatorBinary), baa(RSyntacticName), " &
                "+(ROperatorBinary), a_b(RSyntacticName), +(ROperatorBinary), c12(RSyntacticName), " &
                "+(ROperatorBinary), 1234567890(RSyntacticName), +(ROperatorBinary), " &
                "2.3(RSyntacticName), +(ROperatorBinary), 1e6(RSyntacticName), +(ROperatorBinary), " &
                "abcdefghijklmnopqrstuvwxyz(RSyntacticName), +(ROperatorBinary), `a`(RSyntacticName), " &
                "+(ROperatorBinary), `a b`(RSyntacticName), +(ROperatorBinary), `[[`(RSyntacticName), " &
                "+(ROperatorBinary), `d,ae;af`(RSyntacticName), +(ROperatorBinary), " &
                "`(ah)`(RSyntacticName), +(ROperatorBinary), `ai{aj}`(RSyntacticName), +(ROperatorBinary), " &
                "`~!@#$%^&*()_[] {} \|;:',./<>?`(RSyntacticName), +(ROperatorBinary), " &
                "`%%a_2ab%`(RSyntacticName), +(ROperatorBinary), `%ac%`(RSyntacticName), " &
                "+(ROperatorBinary), `[[""b""]]n[[[o][p]]]`(RSyntacticName), +(ROperatorBinary), " &
                "`if`(RSyntacticName), +(ROperatorBinary), `else`(RSyntacticName), +(ROperatorBinary), " &
                "`while`(RSyntacticName), +(ROperatorBinary), `repeat`(RSyntacticName), " &
                "+(ROperatorBinary), `for`(RSyntacticName), +(ROperatorBinary), `in`(RSyntacticName), " &
                "+(ROperatorBinary), `function`(RSyntacticName), +(ROperatorBinary), " &
                "`return`(RSyntacticName), +(ROperatorBinary), `else`(RSyntacticName), " &
                "+(ROperatorBinary), `next`(RSyntacticName), +(ROperatorBinary), " &
                "`break`(RSyntacticName), "
        Dim strActual As String = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - RBracket, RSeparator
        strInput = "d,ae;af" & vbCr & "ag" & vbLf & "(ah)" & vbCrLf & "ai{aj}"
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected = "d(RSyntacticName), ,(RSeparator), ae(RSyntacticName), ;(REndStatement), " &
                "af(RSyntacticName), " & vbCr & "(REndStatement), ag(RSyntacticName), " &
                vbLf & "(REndStatement), ((RBracket), ah(RSyntacticName), )(RBracket), " &
                vbCrLf & "(REndStatement), ai(RSyntacticName), {(RBracket), aj(RSyntacticName), }(REndScript), "
        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - RSpace
        strInput = " + ay + az + ba   +   bb   +    bc " & vbCr & "  bd   " & vbLf &
                "    be   " & vbCrLf & "  bf "
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected = " (RSpace), +(ROperatorUnaryRight),  (RSpace), ay(RSyntacticName), " &
                " (RSpace), +(ROperatorBinary),  (RSpace), az(RSyntacticName),  (RSpace), " &
                "+(ROperatorBinary),  (RSpace), ba(RSyntacticName),    (RSpace), " &
                "+(ROperatorBinary),    (RSpace), bb(RSyntacticName),    (RSpace), " &
                "+(ROperatorBinary),     (RSpace), bc(RSyntacticName),  (RSpace), " &
                vbCr & "(REndStatement),   (RSpace), bd(RSyntacticName),    (RSpace), " &
                vbLf & "(REndStatement),     (RSpace), be(RSyntacticName),    (RSpace), " &
                vbCrLf & "(REndStatement),   (RSpace), bf(RSyntacticName),  (RSpace), "
        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - RStringLiteral
        strInput = "'a',""bf"",'bga',""bgba"",'bgbaa'," &
                """~!@#$%^&*()_[] {} \|;:',./<>? "","" bgbaaa""" & vbCr & "'bh'" & vbLf &
                """bi""" & vbCrLf & "'bj'{""bk"",'bl',""bm"",'bn' ,""bn"", 'bo'  ," &
                """bq""   ,    'br' " & vbCr & "  ""bs""   " & vbLf & "    'bt'   " & vbCrLf &
                "  ""bu"" '~!@#$%^&*()_[] {} \|;:"",./<>? '"
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected = "'a'(RStringLiteral), ,(RSeparator), ""bf""(RStringLiteral), " &
                ",(RSeparator), 'bga'(RStringLiteral), ,(RSeparator), " &
                """bgba""(RStringLiteral), ,(RSeparator), 'bgbaa'(RStringLiteral), " &
                ",(RSeparator), ""~!@#$%^&*()_[] {} \|;:',./<>? ""(RStringLiteral), " &
                ",(RSeparator), "" bgbaaa""(RStringLiteral), " &
                vbCr & "(REndStatement), 'bh'(RStringLiteral), " &
                vbLf & "(REndStatement), ""bi""(RStringLiteral), " &
                vbCrLf & "(REndStatement), 'bj'(RStringLiteral), {(RBracket), ""bk""(RStringLiteral), " &
                ",(RSeparator), 'bl'(RStringLiteral), ,(RSeparator), " &
                """bm""(RStringLiteral), ,(RSeparator), 'bn'(RStringLiteral),  (RSpace), " &
                ",(RSeparator), ""bn""(RStringLiteral), ,(RSeparator),  (RSpace), " &
                "'bo'(RStringLiteral),   (RSpace), ,(RSeparator), " &
                """bq""(RStringLiteral),    (RSpace), ,(RSeparator),     (RSpace), " &
                "'br'(RStringLiteral),  (RSpace), " &
                vbCr & "(REndStatement),   (RSpace), ""bs""(RStringLiteral),    (RSpace), " &
                vbLf & "(REndStatement),     (RSpace), 'bt'(RStringLiteral),    (RSpace), " &
                vbCrLf & "(REndStatement),   (RSpace), ""bu""(RStringLiteral),  (RSpace), " &
                "'~!@#$%^&*()_[] {} \|;:"",./<>? '(RStringLiteral), "
        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - RComment 
        strInput = "#" & vbLf & "c#" & vbLf & "ca#" & vbLf & "d~#" & vbLf & " #" &
                vbLf & "  #~!@#$%^&*()_[] {} \|;:',./<>?" & vbLf & "#cb" & vbLf & "#cba" & vbLf &
                "# "","" cbaa " & vbLf & "#" & vbCr & "#cc" & vbCr &
                "#cca" & vbCrLf & "# ccaa " & vbCrLf & "#" & vbLf & "e+f#" & vbLf & " #not ignored comment"
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected = "#(RComment), " &
                vbLf & "(RNewLine), c(RSyntacticName), #(RComment), " &
                vbLf & "(REndStatement), ca(RSyntacticName), #(RComment), " &
                vbLf & "(REndStatement), d(RSyntacticName), ~(ROperatorUnaryLeft), #(RComment), " &
                vbLf & "(REndStatement),  (RSpace), #(RComment), " &
                vbLf & "(RNewLine),   (RSpace), #~!@#$%^&*()_[] {} \|;:',./<>?(RComment), " &
                vbLf & "(RNewLine), #cb(RComment), " &
                vbLf & "(RNewLine), #cba(RComment), " &
                vbLf & "(RNewLine), # "","" cbaa (RComment), " &
                vbLf & "(RNewLine), #(RComment), " &
                vbCr & "(RNewLine), #cc(RComment), " &
                vbCr & "(RNewLine), #cca(RComment), " &
                vbCrLf & "(RNewLine), # ccaa (RComment), " &
                vbCrLf & "(RNewLine), #(RComment), " &
                vbLf & "(RNewLine), e(RSyntacticName), +(ROperatorBinary), f(RSyntacticName), #(RComment), " &
                vbLf & "(REndScript),  (RSpace), #not ignored comment(RComment), "

        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - standard operators ROperatorUnaryLeft, ROperatorUnaryRight, ROperatorBinary
        strInput = "a::b:::ba$c@d^e:ea%%f%/%g%*%h%o%i%x%j%in%k/l*m+n-o<p>q<=r>=s==t!=u!v&wx&&y|z" &
                "||a2~2b->c0a->>d0123456789a<-1234567890<<-e0a1b2=345f6789+a/(b)*((c))+(d-e)/f*g" &
                "+(((d-e)/f)*g)+f1(a,b~,c,~d,e~(f+g),h~!i)"
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected = "a(RSyntacticName), ::(ROperatorBinary), b(RSyntacticName), " &
                ":::(ROperatorBinary), ba(RSyntacticName), $(ROperatorBinary), " &
                "c(RSyntacticName), @(ROperatorBinary), d(RSyntacticName), ^(ROperatorBinary), " &
                "e(RSyntacticName), :(ROperatorBinary), ea(RSyntacticName), %%(ROperatorBinary), " &
                "f(RSyntacticName), %/%(ROperatorBinary), g(RSyntacticName), %*%(ROperatorBinary), " &
                "h(RSyntacticName), %o%(ROperatorBinary), i(RSyntacticName), %x%(ROperatorBinary), " &
                "j(RSyntacticName), %in%(ROperatorBinary), k(RSyntacticName), /(ROperatorBinary), " &
                "l(RSyntacticName), *(ROperatorBinary), m(RSyntacticName), +(ROperatorBinary), " &
                "n(RSyntacticName), -(ROperatorBinary), o(RSyntacticName), <(ROperatorBinary), " &
                "p(RSyntacticName), >(ROperatorBinary), q(RSyntacticName), <=(ROperatorBinary), " &
                "r(RSyntacticName), >=(ROperatorBinary), s(RSyntacticName), ==(ROperatorBinary), " &
                "t(RSyntacticName), !=(ROperatorBinary), u(RSyntacticName), !(ROperatorBinary), " &
                "v(RSyntacticName), &(ROperatorBinary), wx(RSyntacticName), &&(ROperatorBinary), " &
                "y(RSyntacticName), |(ROperatorBinary), z(RSyntacticName), ||(ROperatorBinary), " &
                "a2(RSyntacticName), ~(ROperatorBinary), 2b(RSyntacticName), ->(ROperatorBinary), " &
                "c0a(RSyntacticName), ->>(ROperatorBinary), d0123456789a(RSyntacticName), " &
                "<-(ROperatorBinary), 1234567890(RSyntacticName), <<-(ROperatorBinary), " &
                "e0a1b2(RSyntacticName), =(ROperatorBinary), 345f6789(RSyntacticName), " &
                "+(ROperatorBinary), a(RSyntacticName), /(ROperatorBinary), ((RBracket), " &
                "b(RSyntacticName), )(RBracket), *(ROperatorBinary), ((RBracket), ((RBracket), " &
                "c(RSyntacticName), )(RBracket), )(RBracket), +(ROperatorBinary), ((RBracket), " &
                "d(RSyntacticName), -(ROperatorBinary), e(RSyntacticName), )(RBracket), " &
                "/(ROperatorBinary), f(RSyntacticName), *(ROperatorBinary), g(RSyntacticName), " &
                "+(ROperatorBinary), ((RBracket), ((RBracket), ((RBracket), d(RSyntacticName), " &
                "-(ROperatorBinary), e(RSyntacticName), )(RBracket), /(ROperatorBinary), " &
                "f(RSyntacticName), )(RBracket), *(ROperatorBinary), g(RSyntacticName), " &
                ")(RBracket), +(ROperatorBinary), f1(RFunctionName), ((RBracket), " &
                "a(RSyntacticName), ,(RSeparator), b(RSyntacticName), ~(ROperatorUnaryLeft), " &
                ",(RSeparator), c(RSyntacticName), ,(RSeparator), ~(ROperatorUnaryRight), " &
                "d(RSyntacticName), ,(RSeparator), e(RSyntacticName), ~(ROperatorBinary), " &
                "((RBracket), f(RSyntacticName), +(ROperatorBinary), g(RSyntacticName), " &
                ")(RBracket), ,(RSeparator), h(RSyntacticName), ~(ROperatorBinary), " &
                "!(ROperatorUnaryRight), i(RSyntacticName), )(RBracket), "
        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - user-defined operators
        strInput = ".a%%a_2ab%/%ac%*%aba%o%aba2%x%abaa%in%abaaa%>%abcdefg%mydefinedoperator%hijklmnopqrstuvwxyz"
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected = ".a(RSyntacticName), %%(ROperatorBinary), " &
                "a_2ab(RSyntacticName), %/%(ROperatorBinary), " &
                "ac(RSyntacticName), %*%(ROperatorBinary), " &
                "aba(RSyntacticName), %o%(ROperatorBinary), " &
                "aba2(RSyntacticName), %x%(ROperatorBinary), " &
                "abaa(RSyntacticName), %in%(ROperatorBinary), " &
                "abaaa(RSyntacticName), %>%(ROperatorBinary), " &
                "abcdefg(RSyntacticName), %mydefinedoperator%(ROperatorBinary), " &
                "hijklmnopqrstuvwxyz(RSyntacticName), "
        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - ROperatorBracket
        strInput = "a[1]-b[c(d)+e]/f(g[2],h[3],i[4]*j[5])-k[l[m[6]]];df[[""a""]];lst[[""a""]]" &
                "[[""b""]]n[[[o][p]]]"
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected = "a(RSyntacticName), [(ROperatorBracket), 1(RSyntacticName), " &
                "](ROperatorBracket), -(ROperatorBinary), b(RSyntacticName), [(ROperatorBracket), " &
                "c(RFunctionName), ((RBracket), d(RSyntacticName), )(RBracket), +(ROperatorBinary), " &
                "e(RSyntacticName), ](ROperatorBracket), /(ROperatorBinary), f(RFunctionName), " &
                "((RBracket), g(RSyntacticName), [(ROperatorBracket), 2(RSyntacticName), " &
                "](ROperatorBracket), ,(RSeparator), h(RSyntacticName), [(ROperatorBracket), " &
                "3(RSyntacticName), ](ROperatorBracket), ,(RSeparator), i(RSyntacticName), " &
                "[(ROperatorBracket), 4(RSyntacticName), ](ROperatorBracket), *(ROperatorBinary), " &
                "j(RSyntacticName), [(ROperatorBracket), 5(RSyntacticName), ](ROperatorBracket), " &
                ")(RBracket), -(ROperatorBinary), k(RSyntacticName), [(ROperatorBracket), " &
                "l(RSyntacticName), [(ROperatorBracket), m(RSyntacticName), [(ROperatorBracket), " &
                "6(RSyntacticName), ](ROperatorBracket), ](ROperatorBracket), ](ROperatorBracket), " &
                ";(REndStatement), df(RSyntacticName), [[(ROperatorBracket), ""a""(RStringLiteral), " &
                "]](ROperatorBracket), ;(REndStatement), lst(RSyntacticName), [[(ROperatorBracket), " &
                """a""(RStringLiteral), ]](ROperatorBracket), [[(ROperatorBracket), " &
                """b""(RStringLiteral), ]](ROperatorBracket), n(RSyntacticName), [[(ROperatorBracket), " &
                "[(ROperatorBracket), o(RSyntacticName), ](ROperatorBracket), [(ROperatorBracket), " &
                "p(RSyntacticName), ](ROperatorBracket), ]](ROperatorBracket), "
        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - end statement excluding key words
        strInput = "complete" &
                vbLf & "complete()" &
                vbLf & "complete(a[b],c[[d]])" &
                vbLf & "complete #" &
                vbLf & "complete " &
                vbLf & "complete + !e" &
                vbLf & "complete() -f" &
                vbLf & "complete() * g~" &
                vbLf & "incomplete::" &
                vbLf &
                vbLf & "incomplete::h i::: " & vbLf & "ia" &
                vbLf & "incomplete %>% #comment" & vbLf & "ib" &
                vbLf & "incomplete(" & vbLf & "ic)" &
                vbLf & "incomplete()[id " & vbLf & "]" &
                vbLf & "incomplete([[j[k]]]  " & vbLf & ")" &
                vbLf & "incomplete >= " & vbLf & "  #comment " & vbLf & vbLf & "l" & vbLf
        lstInput = clsRScript.GetLstLexemes(strInput)
        strExpected =
                "complete(RSyntacticName), " &
                vbLf & "(REndStatement), complete(RFunctionName), ((RBracket), )(RBracket), " &
                vbLf & "(REndStatement), complete(RFunctionName), ((RBracket), a(RSyntacticName), " &
                "[(ROperatorBracket), b(RSyntacticName), ](ROperatorBracket), ,(RSeparator), " &
                "c(RSyntacticName), [[(ROperatorBracket), d(RSyntacticName), ]](ROperatorBracket), )(RBracket), " &
                vbLf & "(REndStatement), complete(RSyntacticName),  (RSpace), #(RComment), " &
                vbLf & "(REndStatement), complete(RSyntacticName),  (RSpace), " &
                vbLf & "(REndStatement), complete(RSyntacticName),  (RSpace), +(ROperatorBinary), " &
                " (RSpace), !(ROperatorUnaryRight), e(RSyntacticName), " &
                vbLf & "(REndStatement), complete(RFunctionName), ((RBracket), )(RBracket), " &
                " (RSpace), -(ROperatorBinary), f(RSyntacticName), " &
                vbLf & "(REndStatement), complete(RFunctionName), ((RBracket), )(RBracket), " &
                " (RSpace), *(ROperatorBinary),  (RSpace), g(RSyntacticName), ~(ROperatorUnaryLeft), " &
                vbLf & "(REndStatement), incomplete(RSyntacticName), ::(ROperatorBinary), " &
                vbLf & "(RNewLine), " &
                vbLf & "(RNewLine), incomplete(RSyntacticName), ::(ROperatorBinary), " &
                "h(RSyntacticName),  (RSpace), i(RSyntacticName), :::(ROperatorBinary),  (RSpace), " &
                vbLf & "(RNewLine), ia(RSyntacticName), " &
                vbLf & "(REndStatement), incomplete(RSyntacticName),  (RSpace), %>%(ROperatorBinary), " &
                " (RSpace), #comment(RComment), " &
                vbLf & "(RNewLine), ib(RSyntacticName), " &
                vbLf & "(REndStatement), incomplete(RFunctionName), ((RBracket), " &
                vbLf & "(RNewLine), ic(RSyntacticName), )(RBracket), " &
                vbLf & "(REndStatement), incomplete(RFunctionName), ((RBracket), )(RBracket), " &
                "[(ROperatorBracket), id(RSyntacticName),  (RSpace), " &
                vbLf & "(RNewLine), ](ROperatorBracket), " &
                vbLf & "(REndStatement), incomplete(RFunctionName), ((RBracket), [[(ROperatorBracket), " &
                "j(RSyntacticName), [(ROperatorBracket), k(RSyntacticName), ](ROperatorBracket), " &
                "]](ROperatorBracket),   (RSpace), " &
                vbLf & "(RNewLine), )(RBracket), " &
                vbLf & "(REndStatement), incomplete(RSyntacticName),  (RSpace), >=(ROperatorBinary),  (RSpace), " &
                vbLf & "(RNewLine),   (RSpace), #comment (RComment), " &
                vbLf & "(RNewLine), " &
                vbLf & "(RNewLine), l(RSyntacticName), " &
                vbLf & "(REndScript), "
        strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        Assert.Equal(strExpected, strActual)

        'test token list - key words with curly brackets
        'TODO uncomment
        'strInput = "if(x > 10){" & vbLf & "    fn1(paste(x, ""is greater than 10""))" & vbLf & "}" &
        '        vbLf & "else" & vbLf & "{" & vbLf & "    fn2(paste(x, ""Is less than 10""))" &
        '        vbLf & "} " &
        '        vbLf & "while (val <= 5 )" & vbLf & "{" & vbLf & "    # statements" &
        '        vbLf & "    fn3(val)" & vbLf & "    val = val + 1" & vbLf & "}" &
        '        vbLf & "repeat" & vbLf & "{" & vbLf & "    if(val > 5) break" & vbLf & "}" &
        '        vbLf & "for (val in 1:5) {}" &
        '        vbLf & "evenOdd = function(x){" &
        '        vbLf & "if(x %% 2 == 0)" & vbLf & "    return(""even"")" & vbLf & "else" &
        '        vbLf & "    return(""odd"")" & vbLf & "}" &
        '        vbLf & "for (i in val)" & vbLf & "{" & vbLf & "    if (i == 8)" &
        '        vbLf & "        next" & vbLf & "    if(i == 5)" & vbLf & "        break" & vbLf & "}"
        'lstInput = clsRScript.GetLstLexemes(strInput)
        'strExpected =
        '        "if(RKeyWord), ((RBracket), x(RSyntacticName),  (RSpace), >(ROperatorBinary), " &
        '        " (RSpace), 10(RSyntacticName), )(RBracket), {(RBracket), " &
        '        vbLf & "(REndStatement),     (RSpace), fn1(RFunctionName), ((RBracket), " &
        '        "paste(RFunctionName), ((RBracket), x(RSyntacticName), ,(RSeparator),  (RSpace), " &
        '        """is greater than 10""(RStringLiteral), )(RBracket), )(RBracket), " &
        '        vbLf & "(REndStatement), }(REndScript), " &
        '        vbLf & "(RNewLine), else(RKeyWord), " &
        '        vbLf & "(RNewLine), {(RBracket), " &
        '        vbLf & "(REndStatement),     (RSpace), fn2(RFunctionName), ((RBracket), " &
        '        "paste(RFunctionName), ((RBracket), x(RSyntacticName), ,(RSeparator), " &
        '        " (RSpace), ""Is less than 10""(RStringLiteral), )(RBracket), )(RBracket), " &
        '        vbLf & "(REndStatement), }(REndScript),  (RSpace), " &
        '        vbLf & "(RNewLine), " &
        '        "while(RKeyWord),  (RSpace), ((RBracket), val(RSyntacticName),  (RSpace), " &
        '        "<=(ROperatorBinary),  (RSpace), 5(RSyntacticName),  (RSpace), )(RBracket), " &
        '        vbLf & "(RNewLine), {(RBracket), " &
        '        vbLf & "(REndStatement),     (RSpace), # statements(RComment), " &
        '        vbLf & "(RNewLine),     (RSpace), fn3(RFunctionName), ((RBracket), " &
        '        "val(RSyntacticName), )(RBracket), " &
        '        vbLf & "(REndStatement),     (RSpace), val(RSyntacticName),  (RSpace), =(ROperatorBinary), " &
        '        " (RSpace), val(RSyntacticName),  (RSpace), +(ROperatorBinary),  (RSpace), 1(RSyntacticName), " &
        '        vbLf & "(REndStatement), }(REndScript), " &
        '        vbLf & "(RNewLine), " &
        '        "repeat(RKeyWord), " &
        '        vbLf & "(RNewLine), {(RBracket), " &
        '        vbLf & "(REndStatement),     (RSpace), if(RKeyWord), ((RBracket), val(RSyntacticName),  (RSpace), " &
        '        ">(ROperatorBinary),  (RSpace), 5(RSyntacticName), )(RBracket),  (RSpace), break(RKeyWord), " &
        '        vbLf & "(REndStatement), }(REndScript), " &
        '        vbLf & "(REndStatement), " &
        '        "for(RKeyWord),  (RSpace), ((RBracket), val(RSyntacticName),  (RSpace), in(RKeyWord), " &
        '        " (RSpace), 1(RSyntacticName), :(ROperatorBinary), 5(RSyntacticName), )(RBracket), " &
        '        " (RSpace), {(RBracket), }(REndScript), " &
        '        vbLf & "(REndStatement), evenOdd(RSyntacticName),  (RSpace), =(ROperatorBinary),  (RSpace), " &
        '        "function(RKeyWord), ((RBracket), x(RSyntacticName), )(RBracket), {(RBracket), " &
        '        vbLf & "(REndStatement), if(RKeyWord), ((RBracket), x(RSyntacticName),  (RSpace), " &
        '        "%%(ROperatorBinary),  (RSpace), 2(RSyntacticName),  (RSpace), ==(ROperatorBinary), " &
        '        " (RSpace), 0(RSyntacticName), )(RBracket), " &
        '        vbLf & "(RNewLine),     (RSpace), return(RFunctionName), ((RBracket), " &
        '        """even""(RStringLiteral), )(RBracket), " &
        '        vbLf & "(REndScript), else(RKeyWord), " &
        '        vbLf & "(RNewLine),     (RSpace), return(RFunctionName), ((RBracket), " &
        '        """odd""(RStringLiteral), )(RBracket), " &
        '        vbLf & "(REndScript), }(REndScript), " &
        '        vbLf & "(REndStatement), for(RKeyWord),  (RSpace), ((RBracket), i(RSyntacticName), " &
        '        " (RSpace), in(RKeyWord),  (RSpace), val(RSyntacticName), )(RBracket), " &
        '        vbLf & "(RNewLine), {(RBracket), " &
        '        vbLf & "(REndStatement),     (RSpace), if(RKeyWord),  (RSpace), ((RBracket), " &
        '        "i(RSyntacticName),  (RSpace), ==(ROperatorBinary),  (RSpace), 8(RSyntacticName), " &
        '        ")(RBracket), " &
        '        vbLf & "(RNewLine),         (RSpace), next(RKeyWord), " &
        '        vbLf & "(REndScript),     (RSpace), if(RKeyWord), ((RBracket), i(RSyntacticName), " &
        '        " (RSpace), ==(ROperatorBinary),  (RSpace), 5(RSyntacticName), )(RBracket), " &
        '        vbLf & "(RNewLine),         (RSpace), break(RKeyWord), " &
        '        vbLf & "(REndScript), }(REndScript), "
        'strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        'Assert.Equal(strExpected, strActual)

        'test token list - if statement
        ' TODO uncomment
        'strInput = "if(a<b){c}" &
        '        vbLf & "if(d<=e){f}" &
        '        vbLf & "if(g==h) { #1" &
        '        vbLf & " i } #2" &
        '        vbLf & " if (j >= k)" &
        '        vbLf & "{" &
        '        vbLf & "l   #3  " &
        '        vbLf & "}" &
        '        vbLf & "if (m)" & vbLf & "#4" &
        '        vbLf & "  n+" & vbLf & "  o  #5" &
        '        vbLf & "if(p!=q)" &
        '        vbLf & "{" &
        '        vbLf & "incomplete()[id " & vbLf & "]" &
        '        vbLf & "incomplete([[j[k]]]  " & vbLf & ")" &
        '        vbLf & "}" & vbLf
        'lstInput = clsRScript.GetLstLexemes(strInput)
        'strExpected =
        '        "if(RKeyWord), ((RBracket), a(RSyntacticName), <(ROperatorBinary), b(RSyntacticName), )(RBracket), {(RBracket), c(RSyntacticName), }(REndScript), " &
        '        vbLf & "(REndStatement), if(RKeyWord), ((RBracket), d(RSyntacticName), <=(ROperatorBinary), e(RSyntacticName), )(RBracket), {(RBracket), f(RSyntacticName), }(REndScript), " &
        '        vbLf & "(REndStatement), if(RKeyWord), ((RBracket), g(RSyntacticName), ==(ROperatorBinary), h(RSyntacticName), )(RBracket),  (RSpace), {(RBracket),  (RSpace), #1(RComment), " &
        '        vbLf & "(REndStatement),  (RSpace), i(RSyntacticName),  (RSpace), }(REndScript),  (RSpace), #2(RComment), " &
        '        vbLf & "(REndStatement),  (RSpace), if(RKeyWord),  (RSpace), ((RBracket), j(RSyntacticName),  (RSpace), >=(ROperatorBinary),  (RSpace), k(RSyntacticName), )(RBracket), " &
        '        vbLf & "(RNewLine), {(RBracket), " &
        '        vbLf & "(REndStatement), l(RSyntacticName),    (RSpace), #3  (RComment), " &
        '        vbLf & "(REndStatement), }(REndScript), " &
        '        vbLf & "(REndStatement), if(RKeyWord),  (RSpace), ((RBracket), m(RSyntacticName), )(RBracket), " &
        '        vbLf & "(RNewLine), #4(RComment), " &
        '        vbLf & "(RNewLine),   (RSpace), n(RSyntacticName), +(ROperatorBinary), " &
        '        vbLf & "(RNewLine),   (RSpace), o(RSyntacticName),   (RSpace), #5(RComment), " &
        '        vbLf & "(REndScript), if(RKeyWord), ((RBracket), p(RSyntacticName), !=(ROperatorBinary), q(RSyntacticName), )(RBracket), " &
        '        vbLf & "(RNewLine), {(RBracket), " &
        '        vbLf & "(REndStatement), incomplete(RFunctionName), ((RBracket), )(RBracket), [(ROperatorBracket), id(RSyntacticName),  (RSpace), " &
        '        vbLf & "(RNewLine), ](ROperatorBracket), " &
        '        vbLf & "(REndStatement), incomplete(RFunctionName), ((RBracket), [[(ROperatorBracket), j(RSyntacticName), [(ROperatorBracket), k(RSyntacticName), ](ROperatorBracket), ]](ROperatorBracket),   (RSpace), " &
        '        vbLf & "(RNewLine), )(RBracket), " &
        '        vbLf & "(REndStatement), }(REndScript), " &
        '        vbLf & "(REndStatement), "
        'strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        'Assert.Equal(strExpected, strActual)

        '"f11(a,b)" & vbLf &
        '"f12(a,b,c)" & vbLf &
        '"f1(f2(),f3(a),f4(b=1),f5(c=2,3),f6(4,d=5),f7(,),f8(,,),f9(,,,),f10(a,,))" & vbLf &
        '"f0(f1(), f2(a), f3(f4()), f5(f6(f7(b))))" & vbLf &
        '"a/(b)*((c))+(d-e)/f*g+(((d-e)/f)*g)"
        ' 
        'test token list - nested key words with no { TODO doesn't pass yet, see my notes of 03/01/21 for an idea on how to fix
        'strInput =
        '               "for(a in b)" &
        '        vbLf & "    while(c<d)" &
        '        vbLf & "        repeat" &
        '        vbLf & "            if(e=f)" &
        '        vbLf & "                break" &
        '        vbLf & "            else" &
        '        vbLf & "                next" &
        '        vbLf & "if (function(fn1(g,fn2=function(h)fn3(i/sum(j)*100)))))" &
        '        vbLf & "    return(k)" & vbLf
        'lstInput = clsRScript.GetLstLexemes(strInput)
        'strExpected =
        '        ""
        'strActual = GetLstTokensAsString(clsRScript.GetLstTokens(lstInput))
        'Assert.Equal(strExpected, strActual)

    End Sub

    <Fact>
    Sub TestGetAsExecutableScript()
        Dim strInput, strActual As String
        Dim lstScriptPos As ICollection

        strInput = "x[3:5]<-13:15;names(x)[3]<-"" Three""" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(2, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(14, lstScriptPos(1))

        strInput = " f1(f2(),f3(a),f4(b=1),f5(c=2,3),f6(4,d=5),f7(,),f8(,,),f9(,,,),f10(a,,))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(" f1(f2(),f3(a),f4(b =1),f5(c =2,3),f6(4,d =5),f7(,),f8(,,),f9(,,,),f10(a,,))" & vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(1, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))

        strInput = "f0(f1(),f2(a),f3(f4()),f5(f6(f7(b))))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "f0(o4a=o4b,o4c=(o8a+o8b)*(o8c-o8d),o4d=f4a(o6e=o6f,o6g=o6h))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("f0(o4a =o4b,o4c =(o8a+o8b)*(o8c-o8d),o4d =f4a(o6e =o6f,o6g =o6h))" & vbLf, strActual)

        strInput = "a+b+c" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "2+1-10/5*3" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "1+2-3*10/5" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "(a-b)*(c+d)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a/(b)*((c))+(d-e)/f*g+(((d-e)/f)*g)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("a/(b)*(c)+(d-e)/f*g+(((d-e)/f)*g)" & vbLf, strActual)

        strInput = "var1<-pkg1::var2" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "var1<-pkg1::obj1$obj2$var2" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "var1<-pkg1::obj1$fun1(para1,para2)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a<-b::c(d)+e" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "f1(~a,b~,-c,+d,e~(f+g),!h,i^(-j),k+(~l),m~(~n),o/-p,q*+r)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)



        strInput = "a[1]-b[c(d)+e]/f(g[2],h[3],i[4]*j[5])-k[l[m[6]]]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[[1]]-b[[c(d)+e]]/f(g[[2]],h[[3]],i[[4]]*j[[5]])-k[[l[[m[6]]]]]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "df[[""a""]]" & vbLf & "lst[[""a""]][[""b""]]" & vbLf 'same as 'df$a' and 'lst$a$b'
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(2, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(10, lstScriptPos(1))

        strInput = "x<-""a"";df[x]" & vbLf 'same as 'df$a' and 'lst$a$b'
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(2, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(7, lstScriptPos(1))

        strInput = "df<-data.frame(x = 1:10, y = 11:20, z = letters[1:10])" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "x[3:5]<-13:15;names(x)[3]<-""Three""" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(2, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(14, lstScriptPos(1))

        strInput = "x[3:5]<-13:15;" & vbLf & "names(x)[3]<-""Three""" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(3, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(14, lstScriptPos(1))
        Assert.Equal(15, lstScriptPos(2))

        strInput = "x[3:5]<-13:15;" & vbCrLf & "names(x)[3]<-""Three""" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("x[3:5]<-13:15;" & vbLf & "names(x)[3]<-""Three""" & vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(3, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(14, lstScriptPos(1))
        Assert.Equal(16, lstScriptPos(2))

        strInput = "x[3:5]<-13:15;#comment" & vbLf & "names(x)[3]<-""Three""" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(3, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(14, lstScriptPos(1))
        Assert.Equal(23, lstScriptPos(2))

        strInput = "a[]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[,]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[,,]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[,,,]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[b,]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[,c]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[b,c]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[""b"",]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[,""c"",1]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[-1,1:2,,x<5|x>7]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = " a[]#comment" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a [,]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[ ,,] #comment" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[, ,,]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[b, ]   #comment" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("a[b,]   #comment" & vbLf, strActual)

        strInput = "a [  ,   c    ]     " & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("a [  ,   c]     " & vbLf, strActual)

        strInput = "#comment" & vbLf & "a[b,c]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[ ""b""  ,]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a[,#comment" & vbLf & """c"",  1 ]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("a[,#comment" & vbLf & """c"",  1]" & vbLf, strActual)

        strInput = "a[ -1 , 1  :   2    ,     ,      x <  5   |    x      > 7  ]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("a[ -1 , 1  :   2    ,     ,      x <  5   |    x      > 7]" & vbLf, strActual)

        'https://github.com/lloyddewit/RScript/issues/18
        strInput = "weather[,1]<-As.Date(weather[,1],format = ""%m/%d/%Y"")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = " weather  [   ,  #comment" & vbLf & "  1     ] <-  As.Date   (weather     [#comment" & vbLf & " ,  1   ]    ,    format =  ""%m/%d/%Y""    )     " & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(" weather  [   ,  #comment" & vbLf & "  1] <-  As.Date(weather     [#comment" & vbLf & " ,  1],    format =  ""%m/%d/%Y"")     " & vbLf, strActual)

        strInput = "dat <- dat[order(dat$tree, dat$dir), ]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("dat <- dat[order(dat$tree, dat$dir),]" & vbLf, strActual)

        'https://github.com/africanmathsinitiative/R-Instat/pull/8551
        strInput = "d22 <- d22[order(d22$tree, d22$day),]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "res <- MCA(poison[,3:8],excl =c(1,3))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)



        strInput = "data_book$display_daily_table(data_name = ""dodoma"", climatic_element = ""rain"", " &
                   "date_col = ""Date"", year_col = ""year"", Misscode = ""m"", monstats = c(sum = ""sum""))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "stringr::str_split_fixed(string = date,pattern = "" - "",n = ""5 "")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "ggplot2::ggplot(data = c(sum = ""sum""),mapping = ggplot2::aes(x = fert,y = size,colour = variety))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "last_graph<-ggplot2::ggplot(data = survey,mapping = ggplot2::aes(x = fert,y = size,colour = variety))" &
            "+ggplot2::geom_line()" &
            "+ggplot2::geom_rug(colour = ""orange"")" &
            "+theme_grey()" &
            "+ggplot2::theme(axis.text.x = ggplot2::element_text())" &
            "+ggplot2::facet_grid(facets = village~variety,space = ""fixed"")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "dodoma <- data_book$get_data_frame(data_name = ""dodoma"", stack_data = TRUE, measure.vars = c(""rain"", ""tmax"", ""tmin""), id.vars = c(""Date""))" & vbLf &
                   "last_graph <- ggplot2::ggplot(data = dodoma, mapping = ggplot2::aes(x = date, y = value, colour = variable)) + ggplot2::geom_line() + " &
                       "ggplot2::geom_rug(data = dodoma%>%filter(is.na(value)), colour = ""red"") + theme_grey() + ggplot2::theme(axis.text.x = ggplot2::element_text(), legend.position = ""none"") + " &
                       "ggplot2::facet_wrap(scales = ""free_y"", ncol = 1, facet = ~variable) + ggplot2::xlab(NULL)" & vbLf &
                   "data_book$add_graph(graph_name = ""last_graph"", graph = last_graph, data_name = ""dodoma"")" & vbLf &
                   "data_book$get_graphs(data_name = ""dodoma"", graph_name = ""last_graph"")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(4, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(139, lstScriptPos(1))
        Assert.Equal(534, lstScriptPos(2))
        Assert.Equal(623, lstScriptPos(3))

        strInput = "a->b" & vbLf & "c->>d" & vbLf & "e<-f" & vbLf & "g<<-h" & vbLf & "i=j" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(5, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(5, lstScriptPos(1))
        Assert.Equal(11, lstScriptPos(2))
        Assert.Equal(16, lstScriptPos(3))
        Assert.Equal(22, lstScriptPos(4))

        strInput = "x<-df$`a b`" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "names(x)<-c(""a"",""b"")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a<-b" & vbCr & "c(d)" & vbCrLf & "e->>f+g" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("a<-b" & vbLf & "c(d)" & vbLf & "e->>f+g" & vbLf, strActual)

        strInput = " f1(  f2(),   f3( a),  f4(  b =1))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "  f0(   o4a = o4b,  o4c =(o8a   + o8b)  *(   o8c -  o8d),   o4d = f4a(  o6e =   o6f, o6g =  o6h))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = " a  /(   b)*( c)  +(   d- e)  /   f *g  +(((   d- e)  /   f)* g)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = " a  +   b    +     c" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(" a  +   b  +     c" & vbLf, strActual)

        strInput = " var1  <-   pkg1::obj1$obj2$var2" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "    pkg ::obj1 $obj2$fn1 (a ,b=1, c    = 2 )" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("    pkg::obj1$obj2$fn1(a,b =1, c = 2)" & vbLf, strActual)

        strInput = " f1(  ~   a,    b ~,  -   c,    + d,  e   ~(    f +  g),   !    h, i  ^(   -    j), k  +(   ~    l), m  ~(   ~    n), o  /   -    p, q  *   +    r)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = " a  [" & vbCr & "   1" & vbLf & "] -  b   [c (  d   )+ e  ]   /f (  g   [2 ]  ,   h[ " & vbCrLf & "3  ]  " & vbLf & " ,i [  4   ]* j  [   5] )  -   k[ l  [   m[ 6  ]   ]   ]" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(" a  [" & vbCr & "   1] -  b   [c(  d)+ e]   /f(  g   [2],   h[ " & vbCrLf & "3],i [  4]* j  [   5]) -   k[ l  [   m[ 6]]]" & vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strInput & "x" & vbLf).dctRStatements.Keys
        Assert.Equal(2, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(129, lstScriptPos(1))

        strInput = "#precomment1" & vbLf &
                   " # precomment2" & vbLf &
                   "  #  precomment3" & vbLf &
                   " f1(  f2(),   f3( a),  f4(  b =1))#comment1~!@#$%^&*()_[] {} \|;:',./<>?" & vbLf &
                   "  f0(   o4a = o4b,  o4c =(o8a   + o8b)  *(   o8c -  o8d),   o4d = f4a(  o6e =   o6f, o6g =  o6h)) # comment2"","" cbaa " & vbLf &
                   " a  /(   b)*( c)  +(   d- e)  /   f *g  +(((   d- e)  /   f)* g)   #comment3" & vbLf &
                   "#comment 4" & vbLf &
                   " a  +   b  +     c" & vbLf & vbLf & vbLf &
                   "  #comment5" & vbLf &
                   "   # comment6 #comment7" & vbLf &
                   "endSyntacticName" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(5, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(118, lstScriptPos(1))
        Assert.Equal(236, lstScriptPos(2))
        Assert.Equal(313, lstScriptPos(3))
        Assert.Equal(343, lstScriptPos(4))

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal("f1(f2(),f3(a),f4(b =1))" & vbLf &
                     "f0(o4a =o4b,o4c =(o8a+o8b)*(o8c-o8d),o4d =f4a(o6e =o6f,o6g =o6h))" & vbLf &
                     "a/(b)*(c)+(d-e)/f*g+(((d-e)/f)*g)" & vbLf &
                     "a+b+c" & vbLf &
                     "endSyntacticName" & vbLf,
                     strActual)

        strInput = "#comment1" & vbLf & "a#comment2" & vbCr & " b #comment3" & vbCrLf & "#comment4" & vbLf & "  c  " & vbCrLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("#comment1" & vbLf & "a#comment2" & vbLf & " b #comment3" & vbLf & "#comment4" & vbLf & "  c  " & vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(3, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(21, lstScriptPos(1))
        Assert.Equal(35, lstScriptPos(2))

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal("a" & vbLf & "b" & vbLf & "c" & vbLf, strActual)

        strInput = "#not ignored comment"
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput & vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(1, lstScriptPos.Count)

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal(vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strActual).dctRStatements.Keys
        Assert.Equal(1, lstScriptPos.Count)

        strInput = "#not ignored comment" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(1, lstScriptPos.Count)

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal(vbLf, strActual)

        strInput = "f1()" & vbLf & "# not ignored comment" & vbCrLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(2, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(5, lstScriptPos(1))

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal("f1()" & vbLf & vbLf, strActual)

        strInput = "f1()" & vbLf & "# not ignored comment" & vbLf & "# not ignored comment2" & vbCr & " " & vbCrLf & "# not ignored comment3"
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput & vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(2, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(5, lstScriptPos(1))

        'issue lloyddewit/rscript#20
        strInput = "# Code run from Script Window (all text)" & Environment.NewLine & "1"
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput & vbLf, strActual)

        strInput = vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(vbLf, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(1, lstScriptPos.Count)

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal(vbLf, strActual)

        strInput = ""
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("", strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(0, lstScriptPos.Count)

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal("", strActual)

        strInput = Nothing
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("", strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(0, lstScriptPos.Count)

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal("", strActual)

        ' Test string constants that contain line breaks
        strInput = "x <- ""a" & vbLf & """" & vbLf &
                   "fn1(""bc " & vbLf & "d"", e)" & vbLf &
                   "fn2( ""f gh" & vbLf & """,i)" & vbLf &
                   "x <- 'a" & vbCr & vbCr & "'" & vbLf &
                   "fn1('bc " & vbCr & vbCr & vbCr & vbCr & "d', e)" & vbLf &
                   "fn2( 'f gh" & vbCr & "',i)" & vbLf &
                   "x <- `a" & vbCrLf & "`" & vbLf &
                   "fn1(`bc " & vbCrLf & "j" & vbCrLf & "d`, e)" & vbLf &
                   "fn2( `f gh" & vbCrLf & "kl" & vbCrLf & "mno" & vbCrLf & "`,i)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(9, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(10, lstScriptPos(1))
        Assert.Equal(26, lstScriptPos(2))
        Assert.Equal(42, lstScriptPos(3))
        Assert.Equal(53, lstScriptPos(4))
        Assert.Equal(72, lstScriptPos(5))
        Assert.Equal(88, lstScriptPos(6))
        Assert.Equal(99, lstScriptPos(7))
        Assert.Equal(119, lstScriptPos(8))

        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript(False)
        Assert.Equal("x<-""a" & vbLf & """" & vbLf &
                     "fn1(""bc " & vbLf & "d"",e)" & vbLf &
                     "fn2(""f gh" & vbLf & """,i)" & vbLf &
                     "x<-'a" & vbCr & vbCr & "'" & vbLf &
                     "fn1('bc " & vbCr & vbCr & vbCr & vbCr & "d',e)" & vbLf &
                     "fn2('f gh" & vbCr & "',i)" & vbLf &
                     "x<-`a" & vbCrLf & "`" & vbLf &
                     "fn1(`bc " & vbCrLf & "j" & vbCrLf & "d`,e)" & vbLf &
                     "fn2(`f gh" & vbCrLf & "kl" & vbCrLf & "mno" & vbCrLf & "`,i)" & vbLf,
                     strActual)

        ' https://github.com/africanmathsinitiative/R-Instat/issues/7095  
        strInput = "data_book$import_data(data_tables =list(data3 =clipr::read_clip_tbl(x =""Category    Feature    Ease_of_Use     Operating Systems" &
                   vbLf & """, header =TRUE)))" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        ' https://github.com/africanmathsinitiative/R-Instat/issues/7095  
        strInput = "Data <- data_book$get_data_frame(data_name = ""Data"")" & vbLf &
                "last_graph <- ggplot2::ggplot(data = Data |> dplyr::filter(rain > 0.85), mapping = ggplot2::aes(y = rain, x = make_factor("""")))" &
                    " + ggplot2::geom_boxplot(varwidth = TRUE, coef = 2) + theme_grey()" &
                    " + ggplot2::theme(axis.text.x = ggplot2::element_text(angle = 90, hjust = 1, vjust = 0.5))" &
                    " + ggplot2::xlab(NULL) + ggplot2::facet_wrap(facets = ~ Name, drop = FALSE)" & vbLf &
                "data_book$add_graph(graph_name = ""last_graph"", graph = last_graph, data_name = ""Data"")" & vbLf &
                "data_book$get_graphs(data_name = ""Data"", graph_name = ""last_graph"")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)
        Assert.Equal(strInput, strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(4, lstScriptPos.Count)
        Assert.Equal(0, lstScriptPos(0))
        Assert.Equal(53, lstScriptPos(1))
        Assert.Equal(412, lstScriptPos(2))
        Assert.Equal(499, lstScriptPos(3))

        ' https://github.com/africanmathsinitiative/R-Instat/issues/7095  
        strInput = "ifelse(year_2 > 30, 1, 0)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        ' https://github.com/africanmathsinitiative/R-Instat/issues/7377
        strInput = "(year-1900)*(year<2000)+(year-2000)*(year>1999)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        'Test string constants that have quotes embedded inside
        strInput = "a("""", ""\""\"""", ""b"", ""c(\""d\"")"", ""'"", ""''"", ""'e'"", ""`"", ""``"", ""`f`"")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a('', '\'\'', 'b', 'c(\'d\')', '""', '""""', '""e""', '`', '``', '`f`')" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "a(``, `\`\``, `b`, `c(\`d\`)`, `""`, `""""`, `""e""`, `'`, `''`, `'f'`)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "x<-""she said 'hello'""" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "read_clip_tbl(x = ""Ease_of_Use" & vbTab & "Hides R by default to prevent \""code shock\""" & vbTab & "  1"", header = TRUE)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        'issue lloyddewit/rscript#17
        strInput = "?log" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        'issue lloyddewit/rscript#21
        strInput = "?a" & vbLf &
                   "? b" & vbLf &
                   " +  c" & vbLf &
                   "  -   d +#comment1" & vbLf &
                   "(!e) - #comment2" & vbLf &
                   "(~f) +" & vbLf &
                   "(+g) - " & vbLf &
                   "(-h)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal("?a" & vbLf &
                     "? b" & vbLf &
                     " +  c" & vbLf &
                     "  -   d +(!e) -(~f) +(+g) -(-h)" & vbLf,
                     strActual)
        lstScriptPos = New RScript.clsRScript(strInput).dctRStatements.Keys
        Assert.Equal(4, lstScriptPos.Count)

        'issue lloyddewit/rscript#32
        strInput = "??log" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "??a" & vbLf &
                   "?? b" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        'issue lloyddewit/rscript#16
        strInput = """a""+""b""" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "  tfrmt(" & vbLf &
           "  # specify columns in the data" & vbLf &
           "  group = c(rowlbl1, grp)," & vbLf &
           "  label = rowlbl2," & vbLf &
           "  column = column, " & vbLf &
           "  param = param," & vbLf &
           "  value = value," & vbLf &
           "  sorting_cols = c(ord1, ord2)," & vbLf &
           "  # specify value formatting " & vbLf &
           "  body_plan = body_plan(" & vbLf &
           "  frmt_structure(group_val = "".default"", label_val = "".default"", frmt_combine(""{n} ({pct} %)""," & vbLf &
           "                                                                    n = frmt(""xxx"")," & vbLf &
           "                                                                                pct = frmt(""xx.x"")))," & vbLf &
           "    frmt_structure(group_val = "".default"", label_val = ""n"", frmt(""xxx""))," & vbLf &
           "    frmt_structure(group_val = "".default"", label_val = c(""Mean"", ""Median"", ""Min"", ""Max""), frmt(""xxx.x""))," & vbLf &
           "    frmt_structure(group_val = "".default"", label_val = ""SD"", frmt(""xxx.xx""))," & vbLf &
           "    frmt_structure(group_val = "".default"", label_val = "".default"", p = frmt_when("">0.99"" ~ "">0.99""," & vbLf &
           "                                                                                 ""<0.001"" ~ ""<0.001""," & vbLf &
           "                                                                                 TRUE ~ frmt(""x.xxx"", missing = """"))))) %>% " & vbLf &
           "  print_to_gt(data_demog) %>% " & vbLf &
           "  tab_options(" & vbLf &
           "    container.width = 900)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        'issue lloyddewit/rscript#14
        'Examples from https://www.tidyverse.org/blog/2023/04/base-vs-magrittr-pipe/
        strInput = "x %>% f(1, .)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "x |> f(1, y = _)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "df %>% split(.$var)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        'TODO curly brackets not yet supported
        'strInput = "df %>% {split(.$x, .$y)}" & vbLf
        'strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        'Assert.Equal(strInput, strActual)

        strInput = "mtcars %>% .$cyl" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)


        'Examples from https://stackoverflow.com/questions/67744604/what-does-pipe-greater-than-mean-in-r
        strInput = "c(1:3, NA_real_) |> sum(na.rm = TRUE)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "split(x = iris[-5], f = iris$Species) |> lapply(min) |> Do.call(what = rbind)" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "iris[iris$Sepal.Length > 7,] %>% subset(.$Species==""virginica"")" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

        strInput = "1:3 |> sum" & vbLf
        strActual = New RScript.clsRScript(strInput).GetAsExecutableScript()
        Assert.Equal(strInput, strActual)

    End Sub

    Private Function GetLstTokensAsString(lstRTokens As List(Of RScript.clsRToken)) As String

        If lstRTokens Is Nothing OrElse lstRTokens.Count = 0 Then
            'TODO throw exception
            Return Nothing
        End If

        Dim strNew As String = ""
        For Each clsRTokenNew In lstRTokens
            strNew &= clsRTokenNew.strTxt & "("
            Select Case clsRTokenNew.enuToken
                Case RScript.clsRToken.typToken.RSyntacticName
                    strNew &= "RSyntacticName"
                Case RScript.clsRToken.typToken.RFunctionName
                    strNew &= "RFunctionName"
                Case RScript.clsRToken.typToken.RKeyWord
                    strNew &= "RKeyWord"
                Case RScript.clsRToken.typToken.RConstantString
                    strNew &= "RStringLiteral"
                Case RScript.clsRToken.typToken.RComment
                    strNew &= "RComment"
                Case RScript.clsRToken.typToken.RSpace
                    strNew &= "RSpace"
                Case RScript.clsRToken.typToken.RBracket
                    strNew &= "RBracket"
                Case RScript.clsRToken.typToken.RSeparator
                    strNew &= "RSeparator"
                Case RScript.clsRToken.typToken.RNewLine
                    strNew &= "RNewLine"
                Case RScript.clsRToken.typToken.REndStatement
                    strNew &= "REndStatement"
                Case RScript.clsRToken.typToken.REndScript
                    strNew &= "REndScript"
                Case RScript.clsRToken.typToken.ROperatorUnaryLeft
                    strNew &= "ROperatorUnaryLeft"
                Case RScript.clsRToken.typToken.ROperatorUnaryRight
                    strNew &= "ROperatorUnaryRight"
                Case RScript.clsRToken.typToken.ROperatorBinary
                    strNew &= "ROperatorBinary"
                Case RScript.clsRToken.typToken.ROperatorBracket
                    strNew &= "ROperatorBracket"
            End Select
            strNew &= "), "
        Next
        Return strNew
    End Function


End Class


