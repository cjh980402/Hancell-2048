# Hancell-2048
한셀 vbs로 만든 2048 게임 코드

![image](https://user-images.githubusercontent.com/23624124/125466863-36ca7976-0dbb-4e64-8c93-47844ebdce12.png)
해당 소스는 한셀 2014버전에서 작성되었습니다.

**게임 제작 가이드**

1. 한셀스크립트 편집기에 들어가서 소스코드 목록에서 우클릭 하고 모듈 삽입 후, thisworkbook과 모듈소스를 복사해서 입력.

2. 위 스크린샷에 보이는 도형들을 추가하고 각 4방향 화살표에 상, 하, 좌, 우 매크로를 지정,
New Game 버튼에 reset 매크로를 지정, 맨 밑 화살표에 backup 매크로를 지정한 후 껐다가 다시 켜서 키보드나 화살표를 이용하여 동작을 확인.

3. 우클릭 후 매크로 지정할 때 매크로가 목록에 뜨지 않는다면 Onkey 관련된 코드를 지우거나 주석처리 하기. (단, Onkey가 없으면 키보드가 아니라 마우스로만 게임이 가능)

4. 추가로 최근 부대 한셀 버전이 올라가서 제대로 작동이 안 될수도 있지만(Onkey 미작동 등), 기본 알고리즘은 그대로 이용하면서 오류나는 부분만 수정하면 응용해서 활용 가능.
