# VBA 공부 노트

## VBA 란
- 참고 링크
- 
~~~
https://www.youtube.com/watch?v=iRm2dL9Kgeg&list=PLb_vgczBBiSQ3RxE4YAzCClKjdHzad23q
~~~

## 준비하기

### VBA
- Visual Basic(Programming Language) for Application(Office Program)

### Excel에서 VBA 쓰기 위한 환경 설정 절차
- 파일 ▶ 옵션 ▶ 리본 사용자 지정 ▶ [리본 메유 사용자 지정]을 [기본 탭]으로 설정 ▶ 개발 도구 체크 박스 [체크]


### VBA 실행 방법 
- Alt + F11 Or 개발 도구에서 [클릭]


### vscode 확장 프로그램

~~~
https://marketplace.visualstudio.com/items?itemName=local-smart.excel-live-server
~~~

### VBA On Vscode 
~~~
https://www.youtube.com/watch?v=EmFIugafE4U
~~~

### "Visual Basic 프로젝트는 프로그래밍 방식으로 액세스할 수 없습니다" 오류
- XVBA - MACRO LIST 에서 Export VBA를 통해 모듈을 생성하기 위해 Development 및 다른 선택지를 눌렀을 때 해당 오류 발생
- 엑세스 접근 설정을 바꿔줌으로 해결
- 엑셀 실행 ▶ 파일 ▶ 옵션 ▶ 보안센터 ▶ 매크로 설정 ▶ 모든 매크로 포함 체크 ▶ VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음 체크 

~~~
참고 : https://m.blog.naver.com/fjqmgnsdlk/221813821712
~~~


### 객체란 무엇인가
- Ex) 자동차.색상 = 파란색, 객체.속성 = 값

### 엑셀에서의 객체는
워크북 - 워크시트 - 셀

### 변수 설정
dimension 의 줄임말 Dim

Dim 변수이름 As 데이터타입
변수이름 = ~~ 

"" = string
없으면 Number


프로시저란 ?

~~~
https://www.youtube.com/watch?v=jUJNyH1_qMc&list=PLb_vgczBBiSQ3RxE4YAzCClKjdHzad23q&index=5
~~~

### xlsm 파일  매크로 사용 워크시트
- VBA 를 사용하기 위해 매크로 사용 워크시트인 xlsm 확장자명으로 사용하려고 함
- 그래서 기본으로 생성되는 xlsx 파일을 이름 바꾸기로 바꿨더니 다음과 같은 오류 발생
"파일 형식 또는 파일 확장명이 잘못되어 파일을 열 수 없습니다. 파일이 손상되지 않았는지, 파일 확장명이 파일 형식과 일치하는지 확인하십시오."
- 다양한 방법을 시도해봤지만 안됐음
- 마지막 시도로 xlsx 파일 안에서 다른 이름으로 저장할 때 확장명을 xlsm 으로 지정해주어 저장했더니 잘 되는 것을 확인
- 이름 바꾸기와 다른 이름으로 저장으로 xlsm파일을 만들었을 때 2KB 의 크기 차이가 나는 것을 볼 수 있었음
- 예측컨대 다른 이름으로 저장할 때는 xlsm 파일을 구성하는 기본 데이터가 알아서 추가되어 저장되는 것이고 이름 바꾸기로 확장명만 바꾸니까 데이터가 없어서 파일이 손상된 것과 같은 모습인듯 

### VBA 숫자형 데이터타입 표현범위

<img src="https://user-images.githubusercontent.com/114639257/221504928-a2db9866-7cfb-4000-b82b-6c9a085a4877.png" width="600">

### 계산기 
- 개발도구 ▷ 삽입 ▷ 양식컨트롤 1번 ▷ 매크로 지정
- 데이터 타입을 Integer 로 지정하여 만들어주니 Integer의 표현범위 최대값을 넘어서는 값이 나오면 "오버플로" 오류가 발생
- 데이터타입을 Double 로 바꿔주니 해결됨
- 버튼 옮기고 싶으면 Ctrl + 좌클릭으로 지정해주면 됨

<img src="https://user-images.githubusercontent.com/114639257/221506974-38a2e75f-422e-48f0-9fee-219c435613ea.png" width="600">

```

Sub OP_PLUS()  '더하기 연산 프로시저
Dim PARM01 As Double
Dim PARM02 As Double
Dim RE01 As Double
PARM01 = Worksheets("SHEET1").Cells(3, 2).Value
PARM02 = Worksheets("SHEET1").Cells(3, 3).Value
RE01 = PARM01 + PARM02
Worksheets("SHEET1").Cells(3, 4).Value = RE01
End Sub


Sub OP_MINUS()  '빼기 연산 프로시저
Dim PARM01 As Double
Dim PARM02 As Double
Dim RE01 As Double
PARM01 = Worksheets("SHEET1").Cells(3, 2).Value
PARM02 = Worksheets("SHEET1").Cells(3, 3).Value
RE01 = PARM01 - PARM02
Worksheets("SHEET1").Cells(3, 4).Value = RE01
End Sub

Sub OP_MULTI()  '곱하기 연산 프로시저
Dim PARM01 As Double
Dim PARM02 As Double
Dim RE01 As Double
PARM01 = Worksheets("SHEET1").Cells(3, 2).Value
PARM02 = Worksheets("SHEET1").Cells(3, 3).Value
RE01 = PARM01 * PARM02
Worksheets("SHEET1").Cells(3, 4).Value = RE01
End Sub


Sub OP_DIV()  '나누기 연산 프로시저
Dim PARM01 As Double
Dim PARM02 As Double
Dim RE01 As Double
PARM01 = Worksheets("SHEET1").Cells(3, 2).Value
PARM02 = Worksheets("SHEET1").Cells(3, 3).Value
RE01 = PARM01 / PARM02
Worksheets("SHEET1").Cells(3, 4).Value = RE01
End Sub

```

### VBA 조건문을 사용한 계산기

```
Sub F07_01 ()

Dim INPUT01 As Integer
Dim INPUT02 As Integer
Dim OPP As String
Dim RE01 As Integer

INPUT01 = Worksheets("Sheet1").Cells(3, 2).Value
INPUT02 = Worksheets("Sheet1").Cells(3, 3).Value
OPP = Worksheets("Sheet1").Cells(3, 4).Value

If OPP = "+" Then
RE01 = INPUT01 + INPUT02
ElseIf OPP = "-" Then
RE01 = INPUT01 - INPUT02
ElseIf OPP = "*" Then
RE01 = INPUT01 * INPUT02
ElseIf OPP = "/" Then
RE01 = INPUT01 / INPUT02
End If

Worksheets("Sheet1").Cells(3, 5).Value = RE01

End Sub
```

### VBA 비교 연산자

- 소문자 > 대문자
- 순서가 빠를 수록 작다 Ex) a < b, 가 < 나

<img src="https://user-images.githubusercontent.com/114639257/221513789-6049e14f-d3e0-4048-96b0-c09d54a09421.png" width="600">

### VBA 논리 연산자

- and or not

### VBA 종합실습

<img src="https://user-images.githubusercontent.com/114639257/221516255-599a0a7f-16a0-4982-a0ac-8a0034458e37.png" width="600">

### VBA Select Case문
- 특정 상황에서 if문 보다 간결해짐
- 제어문 : 조건문(분기문ㆍ 반복문)

### for next
- 자동증가변수

### 배열변수
- VBA 도 숫자는 0부터
- 변수 선언 똑같은데 크기 설정만 다름 Ex) Dim test(123) As Integer = 크기가 124개인 배열 변수 선언

#### 실습 1
- 구구단
- 공부용.xlsm

~~~
Sub test()

Dim a, b As Integer
For b = 1 To 9
For a = 1 To 9
Worksheets("Sheet1").Cells(a, b).Value = a * b
Next a
Next b

~~~

#### 실습 2 
- 매출
- 공부용.xlsm

~~~
Sub 매출과성과()

Dim i, T As Integer

For i = 5 To 12

Worksheets("매출과성과").Cells(i, 4).Value = Worksheets("매출과성과").Cells(i, 3).Value * 3 / 2

Worksheets("매출과성과").Cells(i, 6).Value = Str(Int(Worksheets("매출과성과").Cells(i, 4).Value / Worksheets("매출과성과").Cells(i, 5).Value * 100)) + "%"

T = Int(Worksheets("매출과성과").Cells(i, 4).Value / Worksheets("매출과성과").Cells(i, 5).Value * 100)

If T >= 90 Then
Worksheets("매출과성과").Cells(i, 7).Value = "A"

ElseIf 90 > T And T >= 80 Then
Worksheets("매출과성과").Cells(i, 7).Value = "B"

ElseIf 80 > T And T >= 70 Then
Worksheets("매출과성과").Cells(i, 7).Value = "C"

ElseIf 70 > T And T >= 60 Then
Worksheets("매출과성과").Cells(i, 7).Value = "D"

Else
Worksheets("매출과성과").Cells(i, 7).Value = "F"

End If

Next

End Sub
~~~


### VBA 함수
- Function =F(x)
- 함수란 프로그래밍을 보다 유연하게 처리할 수 있도록 하는 것
- 자주 쓰이고 많이 활용되는 코드를 사용자가 직접 코딩하지 않도록 도움을 주는 매우 유용한 수식

#### 반올림과 난수
- 반올림 함수 : ROUND(a,b) 인자 2개 a 는 숫자, b는 반올림할 소수점 자리
- 난수 발행 함수 : RND() 인자 없음, 0~1 사이 랜덤한 값 발행
- RND 함수로 특정 범위(0~N)의 랜덤한 수 뽑는 방법

~~~
Rnd의 최솟값이 0.00....1이면 Rnd × N 의 값은 0.00....1이고 Int(Rnd × N)은 0, Int(Rnd × N) + 1 의 값은 1입니다.
Rnd의 최댓값이 0.99....9이면 Rnd × N 의 값은 9.99....9이고 Int(Rnd × N)은 N-1, Int(Rnd × N) + 1 의 값은 N입니다.
따라서 최솟값이 1, 최댓값이 10인 랜덤한 정수를 생성할 수 있게 됩니다.
~~~

