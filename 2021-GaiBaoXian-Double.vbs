''在indesign 把线改成宝线的工具
''说明1：这个程序可以重复运行，直到完成所有替代。
''说明2：这个程序是会改།  །成༑  ༑
''作者：ཀརྨ་སྦྱིན་པ་ལེགས་བཤད།
''2014.10.8
''2021.6.26 修改
Set myIndesign = CreateObject("InDesign.Application")
Dim arr()
for myCounter_00=1 to myIndesign.Documents.count() step 1
	for myCounter_0=1 to myIndesign.Documents.item(myCounter_00).Stories.count() step 1
		for myCounter = 1 to myIndesign.Documents.item(myCounter_00).Stories.item(myCounter_0).Paragraphs.count() step 1
				for each i in myIndesign.Documents.item(myCounter_00).Stories.item(myCounter_0).Paragraphs.item(myCounter).lines
					HNR = i.contents 'HNR 就是每一行的内容了
					Zishu = len(HNR)
					ReDim arr(len(HNR)) '定义一个以每行字符数为下标的数组'
					for j = 0 to len(HNR)-1 step 1 '把行里面的字符分散到数组'
						arr(j) = Mid(HNR, j+1, 1) 
						if arr(j)="༑" then  arr(j)="།"
					next
					HNRnew = Join(arr,"") '把数组结合成字串'
					i.contents = HNRnew
					for k = 0 to len(HNR)-1 step 1 '判断行里面的内容'
						if arr(k)="༄" then 
							exit for '离开，以便到下一行'
						elseif arr(k) = "༅" then 
							exit for
						elseif arr(k) = "་" then
							m=k+1
							if arr(m)="།" then 
								arr(m)="༑"
								if arr(m+1)="།" then
									arr(m+1)="༑"
								elseif arr(m+1) = " " then '改第二个线'
									do while arr(m+1) =" "
										if m+1 < Zishu then
											if arr(m+1)="།" then	arr(m+1) ="༑"
										end if
										m = m+1
									loop
								end if
							end if
							exit for
						elseif arr(k) = "།" then
							if k>0 then 
								arr(k)="༑"
								m=k+1
								if arr(m)="།" then
									arr(m)="༑"
								elseif arr(m) = " " then '改第二个线'
									do while arr(m) =" "
										if m+1 < Zishu then
											if arr(m+1)="།" then	arr(m+1) ="༑"
										end if
										m = m+1
									loop
								end if
								exit for
							end if
							exit for
						elseif arr(k)="ག" then
							m=k+1
							if arr(m)="།" then
								arr(m)="༑"
							elseif arr(m) = " " then '改第二个线'
								do while arr(m) =" "
										if m+1 < Zishu then
											if arr(m+1)="།" then	arr(m+1) ="༑"
										end if
										m = m+1
									loop
							end if
							exit for
						end if
					next	
					HNRnew = Join(arr,"") '把数组结合成字串'
					i.contents = HNRnew
				next
		next
	next
next
msgbox("Done!")