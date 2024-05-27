# VB Scripting Beginner's Basics for CATIA V5

<code><img height = "200"
src = https://github.com/siilver94/vb-scripting-beginner-s-for-CATIA-v5/assets/57824945/cedc8511-4ebc-43e9-a8d1-8e665ce97bc1></code>


## Programming Language vs Scripting Language

<br/>

### Programming language

- **End to End** 소프트웨어 애플리케이션을 구축하는 데 사용되는 소프트웨어 언어입니다.

  *End to End :  시스템 또는 서비스를 처음부터 끝까지 가져오고 일반적으로 제 3 자로부터 아무것도 얻을 필요 없이 완전한 기능 솔루션을 제공하는 프로세스를 설명합니다. 엔드-투-엔드 솔루션은 가능한 한 많은 중간 계층이나 단계를 제거하여 비즈니스의 성능과 효율성을 최적화 합니다.*

  

- 소프트웨어 응용 프로그램을 구축하는 데 사용되는 모든 기능을 가지고  있습니다.

- 프로그래밍 언어로 만들어진 프로그램은 먼저 복잡성을 겪고 실행됩니다

- 스크립팅 언어 보다 많은 기능이 있습니다. (ex: C, C++, JAVA, .NET)

<br/>

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

<br/>

##  VB 스립팅에 대해서 

- **Microsoft** 에 의해 개발 되었고 **Microsoft** 브라우저에서만 지원됩니다.

- **VB Script**는 마이크로소프트가 개발한 액티브 스크립트 언어이다. (이 언어의 구문은 마이크로소프트의 비주얼 베이직 프로그래밍 언어 계통의 일부를 반영한다.)

- **QTP**에 의해 제공되고 있는 스크립팅 언어입니다.

- **QTP**는 **Java Script** 또한 지원하지만 **VB Script** 는 **Java Script** 와 비교할 때 매우 쉽습니다.

- 응용 프로그램을 자동화 하기위한 모든 자동화 도구에는 스크립팅 또는 프로그래밍 기능이 필요합니다

- **VB Script**는 **.vbs** 파일 확장자를 가지고 있습니다.

<br/>
  
## Message Box


  그래픽 사용자 인터페이스에서 대화 상자는 사용자에게 정보를 보여 주거나 응답을 받는 사용자 인터페이스에서 사용되는 특별한 창이다. 대화 상자라고 부르는 까닭은 컴퓨터와 사용자 사이에 대화할 수 있는 기능을 제공하기 때문이다. 

  

  사용자에게 친숙한 메시지를 표시하는 데 사용되는 내장 하위 프로그램입니다.

  **Notepad** 에서 

  *Msgbox "Welcome to VB SCripting Basics for Catia V5 by Mohammed Shakeel"*

  라는 구문 입력 후 file extension 을 **.vbs** 로 저장하면  **vbs** 파일이 생성됨.

<br/>


## Inputbox:

- 사용자로부터 입력을 받는데 사용됩니다.

**Notepad** 에서 

```vbscript
Var1 = inputbox ("Enter your name")

msgbox Var1 
```


코드 입력 후 file extension 을 **.vbs** 로 저장하면 **vbs** 파일이 생성됨.
           

입력 창에 문자열을 입력 을 하면 입력 한 내용이 팝업 됨.

<br/>

## Variables 



- 변수는 다른 시점에서 다른 값을 취할 수 있는 위임자 입니다.

- 변수는 String, Integer, Boolean, decimal 과 같은 다른 타입의 값을  가질 수 있습니다.

  **Notepad** 에서 

  ```vbscript
  varx = "Allen Kim"
  msgbox varx
  
  varx = 12334
  msgbox varx
  
  varx = 2000.2
  msgbox varx
  
  varx = false
  msgbox varx 
  
  ```

  

  를 입력 후 **.vbs** 확장자로 저장 후 실행을 시키면,

  *AllenKim,   12334,  2000.2 , false*  라는 문구가 차례대로 나옵니다.

  

  ### Declaration of Variables

  

  두 개 타입의 변수 선언이 있습니다.

  - **Implicit(암시적)** :  우리는 아무런 선언을 할 필요가 없습니다. **VBS**가 전부 자동으로 선언해 줍니다.

  - **Explicit(명시적)** :  우리는  **Dim(Dimension) Keyword** 를 허용하여 변수를 선언해야 합니다.

    ​                                 eg: Dim var1, dim vnumber

  

  ### Rules for declaring the variables

  

  - 변수의 이름은 **알파벳**으로 시작해야 합니다.

  - 변수의 이름은 특수문자중 **underscore** 만 사용 가능합니다.

  - **underscore** 가 맨 처음으로 오면 안됩니다.

  - 변수의 이름은 255개 이상의 character를 넘어선 안됩니다.

  - 변수의 이름은 **VBS** 의 **Keyword** 를 사용해선 안됩니다.

    

  ### Option Explicit

  - 프로그램이 시작될 때 모든 변수가 선언되도록 하는 명령문입니다.
  - 선언되지 않은 모든 변수를 수집하여 오류를 표시합니다.

  ```vbscript
  option  explicit
  
  dim a, b, c
  
  
  
  a = int(inputbox ("enter the value of a"))
  
  b = int(inputbox ("enter the value of b"))
  
  
  
  c= a+b
  
  
  
  msgbox " the sum of a and b is " &c`
  ```

  

  ## Data Types

  - 데이터 타입은 변수가 가질 수 있는 변수의 타입을 의미합니다.
  - **VBS**에서 사용할 수 있는 유일한 데이터 타입은 **variant** 입니다.
  - **variant** 는 많은 서브타입과 배열을 가집니다.

  ```vbscript
  a = " Carlos"
  msgbox typename(a)
  
  b = 200
  msgbox typename(b)
  
  c=100.01
  msgbox typename(c)
  
  d = true
  msgbox typename(d)
  ```

  *String, Integer, Double, Boolean*

  

  ### List of subtype

  

  1. **Empty**

  2. **Integer**

  3. **Long**

  4. **Double**

  5. **String**

  6. **Boolean**
  7. **Date**

  8. **Array**

  9. **Object**

  

  ##### 1. Empty

  아무런 값도 선언되지 않은 변수는 **Empty** 타입니다.

  

  ##### 2. Integer

  만약 변수가 -32768 ~ 32767  까지의 정수 이면, 데이터 타입은 **Integer** 입니다

  

  ##### 3. Long

  만약 변수가 **Integer** 변수의 한계 값 보다 높으면 **Long** 데이터 타입을 가집니다.

  

  ##### 4. Double

  **Floating point** 값을 저장합니다.

  

  ##### 5. String

  **쌍따옴표** 로 묶인 데이터들은 전부 문자열 입니다.

  **Inputbox** 를 통해 허용되는 모든 데이터는 항상 문자열 입니다.

  

  ##### 6. Boolean

  변수가 **true**  나 **false** 를 포함하면 데이터 타입은 **Boolean** 입니다.

  

  ##### 7. Date

  만약 날짜를 표시해야 된다면,  **# #** 사이에 포함 시켜야 합니다.

  

  ##### 8. Array

  이것은 하나 이상의 값을 저장하는 특별한 **subtype** 입니다.

  배열값은  **Index** 또는 **Subscripts(아래첨자)** 를 사용하여 액세스 할 수 있습니다.

  배열의 사용은 프로그램에서 사용되어지는 변수의 숫자를 줄일 수 있습니다.

  

  

  ### Exercise Scripts 

  - Write a script to calculate simple interest, the inputs should be  

    P as principle

    R as Rate of interest

    T as time in years

  ``` vbscript
  dim p, r, t 
  
  p = cint(inputbox("Enter the value of Principle amount"))
  r = cint(inputbox("Enter the value of rate of interest"))
  t = cint(inputbox("Enter the value of numver of years"))
  
  si = (p* r* t)/100
  
  msgbox " The simple interest value is " &si
  
  ```

  

  

  ## Operators

  1. **Arithmetic Operator**
  2. **Comparison Operator**
  3. **Logical Operator**
  4. **Relational Operator**
  5. **Concatenation Operator**

  

  ##### 1. Arithmetic Operator

  **Arithmetic Operator** 는 가산, 감산을 위하여 쓰인다. 그러므로 이 **Operator**는 plus, minus, divide, multiply, modulus 등이 있다.

  

  ##### **2. Comparison Operator**

  **Comparison Operators** 는=, <, >, <> 등이 있습니다.

  

  ##### 3. Logical Operator

  이 **Operator**는 AND, OR, NOT, NOT of OR, NOT of AND, XOR 등이 있습니다.

  

  ##### 5. Concatenate Operator

  **Concatenate Operator**는 숫자와 문자열로 사용될 수 있습니다. **&** 기호로 문자나 문자열을 연결할 수 있습니다.

  

  ### Order of execution of operators

  1. **Exponent**
  2. **Multiplication**
  3. **Division**
  4. **Mod**
  5. **+**
  6. **-**

  

  

  ## Control Statement

  

  ### Simple If Statements

  ```vbscript
  a = cint(inputbox ("Enter the value of a"))
  b = cint(inputbox ("Enter the value of b"))
  
  if (a>b) then
  msgbox a & " is greater"
  else
  msgbox b & " is greater"
  end if
  ```

  

  아래 코드같이 한문장으로 **If 조건문** 을 끝낼 수 있고, 이 때 **end if** 는 필요가 없습니다.

  ```vbscript
  if (a>b) then msgbox " a is greater"
  
  if (a<b) then msgbox " b is greater"
  ```

  

  ### Nested If Statements

  **If**문 안의 다른 **If**문 입니다.  한 줄로 조건문을 끝내지 않는 한, **end if** 를 써서 **if** 조건문을 끝내야 합니다.

  

  ```vbscript
  'Write a program to find the greater number among the given three numbers using nested if
  
  a = cint(inputbox ("Enter the value of a"))
  b = cint(inputbox ("Enter the value of b"))
  c = cint(inputbox ("Enter the value of c"))
  
  if (a > b) then
  	if (a > c) then
  	msgbox "A is greater"
  	else
  	msgbox "C is greater"
   	end if
  end if
  
  if (b > a) then
  	if (b > c) then
  	msgbox " B is greater"
  	else
  	msgbox " C is greater"
  	end if
  end if
  ```

  

  ### Select Case Statements

  사용자가 값에 따라 명령문 그룹을 실행 하려는 경우 사용합니다. 다른 언어의 **Switch case** 문과 동일합니다.

  1. **Select Case Statement** 는 사용자로부터 직접 입력으로 사용되고 프로그램을 **case** 로 실행합니다.
  2. 이것을 사용 하면 긴급하게 **debug**를 할 때 코드를 좀 더 이해하기 쉽게 합니다.
  3. **'Select case'** 가 키워드 이고 **'end select'** 로 끝냅니다.
  4. 관계 프로그램을 선택하기 좋습니다.
  5. 어느 프로그램 에서든 **Select Case Statement** 을 사용할 수 있습니다. **If**문을 사용하여 작성할 수도 있습니다. 하지만 그 반대의 경우가 항상 맞는 것은 아닙니다.

  

  *예시 :  거래 수단 선택.*

  ```vbscript
  varchoice = inputbox ("Enter your choice a: Card b: cash c: cheque d: DD")
  
  select case varchoice
  	case "a"
  	msgbox "Card Option is chosen"
  
  	case "b"
  	msgbox "Cash Option is chosen"
  
  	case "c"
  	msgbox "Cheque Option is chosen"
  
  	case "d"
  	msgbox "DD Option is chosen"
  	
  	case else 
  msgbox " You have entered an invalid option"
  
  end select
  ```



*예제: 전자 계산기.*

```vbscript
varchoice = lcase(inputbox ("Enter your choice of operator :"&vbnewline& " a->addition s->subtraction m->multiplication d->division i->integer division e->exponential mod->modulus"))

var1 = cdbl (inputbox("Enter the first value"))
var2 = cdbl(inputbox("Enter the second value"))

select case varchoice

	case "a"
	msgbox (var1+var2)

	case "s"
	msgbox (var1-var2)

	case "m"
	msgbox (var1/var2)

	case "d"
	msgbox (var1/var2)

	case "i"
	msgbox (var1\var2)

	case "e"
	msgbox (var1 ^ var2)

	case "mod"
	msgbox (var1 mod var2)

	case else
	msgbox "You have entered the wrong operator"

    
end select
```




