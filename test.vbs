

Class TestClass
  Private Sub Class_Initialize(x)  ' 设置 Initialize 事件，相当于构造函数
    MsgBox("TestClass started")
	MsgBox x
  End Sub
  Private Sub Class_Terminate  ' 设置 Terminate 事件，相当于析构函数
    MsgBox("TestClass terminated")
  End Sub
End Class

Set X = New TestClass("1111")  ' 创建一个 TestClass 实例
Set X = Nothing   ' 删除实例







