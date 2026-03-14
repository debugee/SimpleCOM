set simple = CreateObject("DualInterfaceComponent.1")

MsgBox simple.Add(1, 0)

MsgBox simple.Subtract(3, 1)

MsgBox simple.Multiply(3, 1)

MsgBox simple.Divide(8, 2)

MsgBox simple.GetLength("hello")

MsgBox simple.ToUpper("hello world")

MsgBox simple.ToLower("HELLO WORLD")

MsgBox simple.Reverse("hello world")