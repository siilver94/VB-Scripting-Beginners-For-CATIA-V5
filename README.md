# VB Scripting Beginner's Basics for CATIA V5



## Programming Language vs Scripting Language



### Programming language

- **End to End** 소프트웨어 애플리케이션을 구축하는 데 사용되는 소프트웨어 언어입니다.

  *End to End :  시스템 또는 서비스를 처음부터 끝까지 가져오고 일반적으로 제 3 자로부터 아무것도 얻을 필요 없이 완전한 기능 솔루션을 제공하는 프로세스를 설명합니다. 엔드-투-엔드 솔루션은 가능한 한 많은 중간 계층이나 단계를 제거하여 비즈니스의 성능과 효율성을 최적화 합니다.*

  

- 소프트웨어 응용 프로그램을 구축하는 데 사용되는 모든 기능을 가지고  있습니다.

- 프로그래밍 언어로 만들어진 프로그램은 먼저 복잡성을 겪고 실행됩니다

- 스크립팅 언어 보다 많은 기능이 있습니다. (ex: C, C++, JAVA, .NET)






### Scripting language 

- 스크립트 또는 스크립팅 언어는 일련의 컴퓨터 언어 명령 없이 실행 될 수 있는 파일 내에서 컴파일 합니다. 서버 측 스크립팅 언어 의 좋은 예 는 **Perl , PHP** 및 **Python** 입니다. 클라이언트 측 스크립팅 언어의 가장 좋은 예는 **JavaScript** 입니다. 
- 사용자가 필요한 경우 스크립트를 보고 편집 할 수 있는 오픈 소스 파일을 컴파일 할 필요는 없지만 필요할 때 사용할 수 있습니다.
- 다른 운영 체제 간에 쉽게 포팅 할 수 있습니다 .
- 실제 프로그램보다 개발 속도가 훨씬 빠릅니다. 일부 개인과 회사는 스크립트를 실제 
  프로그램의 프로토 타입으로 작성합니다.


- 소프트웨어 응용 프로그램의 종단 간 테스트에 사용되거나 큰 응용 
  프로그램에 작은 구성 요소를 구축하는 데 사용됩니다

- 프로그래밍 언어와 비교할 때 기능이 제한 되어 있습니다.

- 복잡성과 실행을 겪지 않고 선으로 진행됩니다.

- **이해하기 쉽고 구현하기가 매우 쉽습니다.**(ex: Java script, VB Script, Shell Script)

   


##  VB 스립팅에 대해서 

- **Microsoft** 에 의해 개발 되었고 **Microsoft** 브라우저에서만 지원됩니다.

- **VB Script**는 마이크로소프트가 개발한 액티브 스크립트 언어이다. (이 언어의 구문은 마이크로소프트의 비주얼 베이직 프로그래밍 언어 계통의 일부를 반영한다.)

- **QTP**에 의해 제공되고 있는 스크립팅 언어입니다.

- **QTP**는 **Java Script** 또한 지원하지만 **VB Script** 는 **Java Script** 와 비교할 때 매우 쉽습니다.

- 응용 프로그램을 자동화 하기위한 모든 자동화 도구에는 스크립팅 또는 프로그래밍 기능이 필요합니다

- **VB Script**는 **.vbs** 파일 확장자를 가지고 있습니다.

  

  

  ## Message Box

  

  그래픽 사용자 인터페이스에서 대화 상자는 사용자에게 정보를 보여 주거나 응답을 받는 사용자 인터페이스에서 사용되는 특별한 창이다. 대화 상자라고 부르는 까닭은 컴퓨터와 사용자 사이에 대화할 수 있는 기능을 제공하기 때문이다. 

  

  사용자에게 친숙한 메시지를 표시하는 데 사용되는 내장 하위 프로그램입니다.

  **Notepad** 에서 

  *Msgbox "Welcome to VB SCripting Basics for Catia V5 by Mohammed Shakeel"*

  라는 구문 입력 후 file extension 을 **.vbs** 로 저장하면  **vbs** 파일이 생성됨.




## Inputbox:

- 사용자로부터 입력을 받는데 사용됩니다.

**Notepad** 에서 

```vbscript
Var1 = inputbox ("Enter your name")

msgbox Var1 
```


코드 입력 후 file extension 을 **.vbs** 로 저장하면 **vbs** 파일이 생성됨.
           

입력 창에 문자열을 입력 을 하면 입력 한 내용이 팝업 됨.





