Dim fso, f, NeedPress, temp, arrParts, i, intRightParenthesisPos, PutKey, PressKeyList(), j, breaktm, k, Sttime
On Error Resume Next
set fso = CreateObject("Scripting.FileSystemObject")
set f = fso.OpenTextFile("songs\play.txt", 1, false) '第二个参数 1 表示只读打开，第三个参数表示目标文件不存在时是否创建
set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 3000

'请根据谱子自由修改这个（每一节的长短）
Sttime = 400

ReDim PressKeyList(0)
do
	temp = UCase(trim(f.ReadLine()))
	NeedPress = temp
	if Len(NeedPress) > 0 then '跳过空行
		if StrComp(left(temp,1),"#") then '跳过注释行
'现在NeedPress中就是需要按下的按钮了，正式处理程序开始

temp = NeedPress 
arrParts = Split(temp, "/")
For i = 0 To UBound(arrParts) - 1 ' UBound 函数获取数组最大索引
	temp = LTrim(arrParts(i))
	j = 0
	do
		if Len(temp) > 0 then 
			NeedPress = left(temp,1)
			if StrComp(NeedPress,"(") then
				PutKey = NeedPress
				temp = Mid(temp, 2)
			else
				NeedPress = temp
				intRightParenthesisPos = InStr(1, NeedPress, ")") - 1
				PutKey = Mid(NeedPress,2,intRightParenthesisPos)
				intRightParenthesisPos = Len(temp) - intRightParenthesisPos - 1
				if Len(temp)>intRightParenthesisPos then
					temp = Right(temp, intRightParenthesisPos)
				else
					temp = ""
				end if
			end if
			ReDim Preserve PressKeyList(j)
			PressKeyList(j) = PutKey
			j = j + 1
		else
			'真正的按键-PressKeyList模拟
			If j = 0 then
				WScript.Sleep Sttime
			ElseIf j = 1 then
				WshShell.SendKeys CStr(PressKeyList(0))
				WScript.Sleep Sttime
			else
				breaktm = Int(Sttime / j)
				j = j - 1
				For k = 0 To j
					WshShell.SendKeys CStr(PressKeyList(k))
					WScript.Sleep breaktm
				Next
			end if
			ReDim PressKeyList(0)
			exit do
		end if
	loop
Next

'正式处理程序结束
		end if
	end if
	if f.atEndOfStream then
		exit do
	end if
loop
f.Close()
set f = nothing
set fso = nothing