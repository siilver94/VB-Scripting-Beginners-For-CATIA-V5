### Create a New Part and Save it in any location


**Macro**를 실행시키면 새로운 CATIA Part가 열리고 파일이 자동으로 저장되는 것이 목적 입니다.

l**earning_2** 라는 새로운 **Macro** 생성합니다.


![Untitled](https://user-images.githubusercontent.com/57824945/90393523-e39e3280-e0cb-11ea-9fb5-143aea4ac8a0.png)


Part 를 추가 시키고, 명시하는 directory 에 저장합니다.

```vbscript
Sub CATMain()

Dim newdoc As Document
Set newdoc = CATIA.Documents.Add("Part")

Dim savedoc As Document
Set savedoc = CATIA.ActiveDocument
savedoc.SaveAs ("C:\test\newtest.CATPart")

End Sub

```

