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

