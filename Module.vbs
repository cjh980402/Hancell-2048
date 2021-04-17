Sub save
For i=1 To 4
	For j=1 To 4
		If cells(i,j)<>"" Then
			cells(9+i,j)=cells(i,j)*1
		else
			cells(9+i,j).clearcontents
		End If
	Next
Next

cells(11,10)=cells(2,10)*1
cells(13,10)=cells(4,10)*1

'cells(50,50).Select
End Sub

Sub backup
activesheet.unprotect "Tkdlqjrj"
For i=1 To 4
	For j=1 To 4
		If cells(9+i,j)<>"" Then
			cells(i,j)=cells(9+i,j)*1
		else
			cells(i,j).clearcontents
		End If
	Next
Next

cells(2,10)=cells(11,10)*1
cells(4,10)=cells(13,10)*1

Call color
activesheet.protect "Tkdlqjrj", false, true
End Sub

Sub color
For i=1 To 4
	For j=1 To 4
		cells(i,j).font.color=RGB(249,246,242)
		If cells(i,j)="" Then
			cells(i,j).interior.color=RGB(205,193,180)
			cells(i,j).font.color=RGB(205,193,180)
			
		elseIf cells(i,j)>=2 and cells(i,j)<4 Then
			cells(i,j).interior.color=RGB(238,228,218)
			cells(i,j).font.color=RGB(119,110,101)
			
		elseif cells(i,j)<8 Then
			cells(i,j).interior.color=RGB(237,224,200)
			cells(i,j).font.color=RGB(119,110,101)
			
		elseif cells(i,j)<16 Then
			cells(i,j).interior.color=RGB(242,177,121)
			
		elseif cells(i,j)<32 Then
			cells(i,j).interior.color=RGB(245,149,99)
			
		elseif cells(i,j)<64 Then
			cells(i,j).interior.color=RGB(246,124,95)
			
		elseif cells(i,j)<128 Then
			cells(i,j).interior.color=RGB(246,94,59)
			
		elseif cells(i,j)<256 Then
			cells(i,j).interior.color=RGB(237,207,114)
			
		elseif cells(i,j)<512 Then
			cells(i,j).interior.color=RGB(237,204,97)
			
		elseif cells(i,j)<1024 Then
			cells(i,j).interior.color=RGB(237,200,80)
			
		elseif cells(i,j)<2048 Then
			cells(i,j).interior.color=RGB(237,197,64)
			
		elseif cells(i,j)<4096 Then
			cells(i,j).interior.color=RGB(237,194,46)
			
		elseif cells(i,j)>=4096 Then
			cells(i,j).interior.color=RGB(60,58,50)
			cells(i,j).font.color=RGB(247,244,240)
		else
			cells(i,j).interior.color=RGB(205,193,180)
			cells(i,j).font.color=RGB(205,193,180)
		
		End If	
	Next
Next

End Sub

Sub gameover

For i=1 To 4
	For j=1 To 4
		If cells(i,j)<>"" Then
			check1=check1+1
		End If
	Next
Next

If check1=16 Then
	For i=1 To 4
		For j=1 To 3
			If cells(i,j)=cells(i,j+1) Then
				check2=1
				Exit For
			End If
		Next
	Next
	If check2=0 Then
		For i=1 To 4
			For j=1 To 3
				If cells(j,i)=cells(j+1,i) Then
					check2=1
					Exit For
				End If
			Next
		Next
	End If
End If

If check1=16 and check2=0 Then
	Call save
	MsgBox "Game Over!!" & Chr(10)&Chr(13)& "최종 점수 : "&cells(2,10)&"점", 64, "2048 게임"
	If cells(2,10) >= cells(4,10) Then
		MsgBox "최고 점수를 달성하였습니다!!", 64, "2048 게임"
	End If
End If

End Sub

Sub random(move, num)
Randomize

Call color

i=0
j=0

Dim arr(3)
arr(0)=2
arr(1)=2
arr(2)=2
arr(3)=4

Dim row(16)
Dim col(16)
Dim count

If move=1 Then
	For a=1 To num
		count=0
		For i=1 To 4
			For j=1 To 4
				If cells(i,j)="" Then
					col(count)=i
					row(count)=j
					count=count+1
				End If
			Next
		Next
		
		r=Int(count*Rnd)
		i=col(r)
		j=row(r)
		
		cells(i,j)=arr(Int(Rnd*4))*1
		cells(i,j).interior.color=RGB(130,230,255)
		cells(i,j).font.color=RGB(119,110,101)
	Next
		
End If

End Sub

Sub reset
activesheet.unprotect "Tkdlqjrj"
application.screenupdating=false
range(cells(1,1),cells(4,4)).interior.color=RGB(205,193,180)
range(cells(1,1),cells(4,4)).font.color=RGB(205,193,180)
range(cells(1,1),cells(4,4)).clearcontents
range(cells(10,1),cells(13,4)).clearcontents

cells(2,10)=0
Call random(1,2)
Call save
application.screenupdating=true
activesheet.protect "Tkdlqjrj", false, true
End Sub

Sub 상
activesheet.unprotect "Tkdlqjrj"
application.screenupdating=false
Call save

move=0

For j=1 To 4
	For i=2 To 4
		If cells(i,j)<>"" Then
			For k=i-1 To 1 step -1
				If cells(k,j)="" Then
					cells(k,j)=cells(k+1,j)*1
					move=1
					cells(k+1,j).clearcontents
				else
					Exit For
				End If
			Next
		End If
	Next
	
	For i=1 To 3
		If cells(i,j)<>"" and cells(i,j)=cells(i+1,j) Then
			move=1
			score=8
			If cells(i,j)>=1024 Then
				score=16
			End If
			
			cells(2,10)=cells(2,10)+cells(i,j)*score
			If cells(2,10)>=cells(4,10) Then
				cells(4,10)=cells(2,10)*1
				ActiveWorkbook.Save
			End If
			
			cells(i,j)=cells(i+1,j)*2
			cells(i+1,j).clearcontents
			cells(i+1,j).interior.color=RGB(205,193,180)
			
			If cells(i,j)=2048 Then
				MsgBox "2048이 만들어졌습니다!",64,"2048 게임"
			End If
			
			For k=i+1 To 3
				If cells(k+1,j)="" Then
						cells(k,j).clearcontents
					else
						cells(k,j)=cells(k+1,j)*1
				End If
				cells(k+1,j).clearcontents
			Next
		End If
	Next
Next

Call random(move,1)
Call gameover
application.screenupdating=true
activesheet.protect "Tkdlqjrj", false, true
End Sub

Sub 하
activesheet.unprotect "Tkdlqjrj"
application.screenupdating=false
Call save

move=0

For j=1 To 4
	For i=2 To 4
		If cells(5-i,j)<>"" Then
			For k=i-1 To 1 step -1
				If cells(5-k,j)="" Then
					cells(5-k,j)=cells(5-(k+1),j)*1
					move=1
					cells(5-(k+1),j).clearcontents
				else
					Exit For
				End If
			Next
		End If
	Next
	
	For i=1 To 3
		If cells(5-i,j)<>"" and cells(5-i,j)=cells(5-(i+1),j) Then
			move=1
			score=8
			If cells(5-i,j)>=1024 Then
				score=16
			End If
			
			cells(2,10)=cells(2,10)+cells(5-i,j)*score
			If cells(2,10)>=cells(4,10) Then
				cells(4,10)=cells(2,10)*1
				ActiveWorkbook.Save
			End If
			
			cells(5-i,j)=cells(5-(i+1),j)*2
			cells(5-(i+1),j).clearcontents
			cells(5-(i+1),j).interior.color=RGB(205,193,180)
			
			If cells(5-i,j)=2048 Then
				MsgBox "2048이 만들어졌습니다!",64,"2048 게임"
			End If
			
			For k=i+1 To 3
				If cells(5-(k+1),j)="" Then
						cells(5-k,j).clearcontents
					else
						cells(5-k,j)=cells(5-(k+1),j)*1
				End If
				cells(5-(k+1),j).clearcontents
			Next
		End If
	Next
Next

Call random(move,1)
Call gameover
application.screenupdating=true
activesheet.protect "Tkdlqjrj", false, true
End Sub

Sub 좌
activesheet.unprotect "Tkdlqjrj"
application.screenupdating=false
Call save

move=0

For j=1 To 4
	For i=2 To 4
		If cells(j,i)<>"" Then
			For k=i-1 To 1 step -1
				If cells(j,k)="" Then
					cells(j,k)=cells(j,k+1)*1
					move=1
					cells(j,k+1).clearcontents
				else
					Exit For
				End If
			Next
		End If
	Next
	
	For i=1 To 3
		If cells(j,i)<>"" and cells(j,i)=cells(j,i+1) Then
			move=1
			score=8
			If cells(j,i)>=1024 Then
				score=16
			End If
			
			cells(2,10)=cells(2,10)+cells(j,i)*score
			If cells(2,10)>=cells(4,10) Then
				cells(4,10)=cells(2,10)*1
				ActiveWorkbook.Save
			End If
			
			cells(j,i)=cells(j,i+1)*2
			cells(j,i+1).clearcontents
			cells(j,i+1).interior.color=RGB(205,193,180)
			
			If cells(j,i)=2048 Then
				MsgBox "2048이 만들어졌습니다!",64,"2048 게임"
			End If
			
			For k=i+1 To 3
				
				If cells(j,k+1)="" Then
						cells(j,k).clearcontents
					else
						cells(j,k)=cells(j,k+1)*1
				End If
				
				cells(j,k+1).clearcontents
			Next
		End If
	Next
Next

Call random(move,1)
Call gameover
application.screenupdating=true
activesheet.protect "Tkdlqjrj", false, true
End Sub

Sub 우
activesheet.unprotect "Tkdlqjrj"
application.screenupdating=false
Call save

move=0


For j=1 To 4
	For i=2 To 4
		If cells(j,5-i)<>"" Then
			For k=i-1 To 1 step -1
				If cells(j,5-k)="" Then
					
					cells(j,5-k)=cells(j,5-(k+1))*1
					move=1
					cells(j,5-(k+1)).clearcontents
				else
					Exit For
				End If
			Next
		End If
	Next
	
	For i=1 To 3
		If cells(j,5-i)<>"" and cells(j,5-i)=cells(j,5-(i+1)) Then
			move=1
			score=8
			If cells(j,5-i)>=1024 Then
				score=16
			End If
			
			cells(2,10)=cells(2,10)+cells(j,5-i)*score
			If cells(2,10)>=cells(4,10) Then
				cells(4,10)=cells(2,10)*1
				ActiveWorkbook.Save
			End If
			
			cells(j,5-i)=cells(j,5-(i+1))*2
			cells(j,5-(i+1)).clearcontents
			cells(j,5-(i+1)).interior.color=RGB(205,193,180)
			
			If cells(j,5-i)=2048 Then
				MsgBox "2048이 만들어졌습니다!",64,"2048 게임"
			End If
			
			For k=i+1 To 3
				If cells(j,5-(k+1))="" Then
						cells(j,5-k).clearcontents
					else
						cells(j,5-k)=cells(j,5-(k+1))*1
				End If
				cells(j,5-(k+1)).clearcontents
			Next
		End If
	Next
Next

Call random(move,1)
Call gameover
application.screenupdating=true
activesheet.protect "Tkdlqjrj", false, true
End Sub

Sub keyset
	application.onkey "{UP}","상"
	application.onkey "{DOWN}","하"
	application.onkey "{LEFT}","좌"
	application.onkey "{RIGHT}","우"
	
	application.onkey "{F5}","reset"
	
	application.onkey "{BS}","backup"
End Sub

Sub unkeyset
	application.onkey "{UP}"
	application.onkey "{DOWN}"
	application.onkey "{LEFT}"
	application.onkey "{RIGHT}"
	
	application.onkey "{F5}"
	
	application.onkey "{BS}"
End Sub
