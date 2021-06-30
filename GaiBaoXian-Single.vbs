''在indesign 把线改成宝线的工具
''说明1：这个程序可以重复运行，直到完成所有替代。
''说明2：这个程序是会改།  །成༑  །
''作者：ཀརྨ་སྦྱིན་པ་ལེགས་བཤད།
''2014.10.8

Set myIndesign = CreateObject("InDesign.Application")
Set myDocument = myIndesign.Documents.Item(1)
''Set myStory = myDocument.Stories.Item(1)
Dim arr()
Dim SHYX

SHYX=0
for Rt= 1 to 2
	for each ms in myDocument.Stories
		for each pg in ms.Paragraphs
			for each i in pg.lines
				HNR = i.contents 'HNR 就是每一行的内容了
				ReDim arr(len(HNR)) '定义一个以每行字符数为下标的数组'
				for j = 0 to len(HNR)-1 step 1 '把行里面的字符分散到数组'
					arr(j) = Mid(HNR, j+1, 1) 
				next
				
				for k = 0 to len(HNR)-1 step 1 '判断行里面的内容'
					if arr(k)="༄" then
						for n= k to len(HNR)-1 step 1 '把该行里不必要的宝线删除'
							if arr(n)="༑" then arr(n)="།"
						next
						exit for '离开，以便到下一行'
					end if
					
					if arr(k)="༅" then
						for n= k to len(HNR)-1 step 1 '把该行里不必要的宝线删除'
							if arr(n)="༑" then arr(n)="།"
						next
						exit for
					end if
					
					if arr(k)="་" then
						if arr(k+1)="།" then arr(k+1)="༑"
						
						for n= k + 2 to len(HNR)-1 step 1 '把该行里不必要的宝线删除'
							if arr(n)="༑" then arr(n)="།"
						next
						exit for
					end if
					
					if arr(k)="།" then
						if k=0 then SHYX= SHYX+1 '计算行首出现线的情况'
						
						if k>0 then arr(k)="༑"
						for n = k+1 to len(HNR)-1 step 1 '把该行里不必要的宝线删除'
							if arr(n)="༑" then arr(n)="།"
						next
						exit for '离开，以便到下一行'
					end if
					
					if arr(k)="ག" then
						if arr(k+1)="།" then arr(k+1)="༑"
						if arr(k+1)=" "	 then
							p=k+1
							do while arr(p)<>"་"
								if arr(p)="།" then
									arr(p)="༑"
								end if
								p = p+1
							loop
							for n= p+1 to len(HNR)-1 step 1 '把该行里不必要的宝线删除'
								if arr(n)="༑" then arr(n)="།"
							next
							exit for
						end if
					end if
				next	
				HNRnew = Join(arr,"") '把数组结合成字串'
				i.contents = HNRnew
			next
		next
	next
next
msgbox("本文里还有" & SHYX & "个线在行首建议检查")