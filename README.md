# PutCall
Paridade put call vba

Código escrito abaixo para paridade put call do mercado financeiro de opções. Para completar o programa basta somente fazer o userform e inserir as variáveis nas textbox corretas.



![Image](https://user-images.githubusercontent.com/67772460/198428705-9cd8f973-6079-4e26-9d10-da4c0e5a2e21.png)








`Private Sub CommandButton1_Click()
Dim portfolio_a As Single, s As Single, c As Single, r As Single, X As Single, p As Single, t As Single, portfolio_b As Single, res As String, ab As String, ba As String


s = TextBox2.Value
X = TextBox3.Value
t = TextBox4.Value
r = TextBox5.Value
c = TextBox6.Value
p = TextBox7.Value

t = t / 252

portfolio_a = c + (X * Exp(1) ^ (-r * t))
portfolio_b = s + p


ab = portfolio_a - portfolio_b
ba = portfolio_b - portfolio_a

UserForm1.TextBox8 = FormatNumber(portfolio_a, 2)
UserForm1.TextBox9 = FormatNumber(portfolio_b, 2)



If portfolio_a > portfolio_b Then
    UserForm1.TextBox10 = ("Vender Portfolio A e comprar B") + vbCrLf + vbCrLf + "Prêmio de " & FormatNumber(ab, 2) & " reais"
ElseIf portfolio_a < portfolio_b Then
    UserForm1.TextBox10 = ("Vender Portfolio B e comprar A") + vbCrLf + vbCrLf + "Prêmio de " & FormatNumber(ba, 2) & " reais"
ElseIf portfolio_a = portfolio_b Then
    UserForm1.TextBox10 = ("Não há prêmio")
End If




End Sub




Private Sub CommandButton2_Click()
Dim ctl As Control
For Each ctl In Me.Controls
  If TypeName(ctl) = "TextBox" Then
    ctl.Text = vbNullString
  End If
Next ctl
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub TextBox10_Change()

End Sub

Private Sub TextBox2_Change()

End Sub`
