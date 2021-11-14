# Tools For Data Standard
Data Architecture 구성요소중 데이터 표준사전 관련 도구

## Standard Checker
표준단어/도메인/용어 사전을 작성하고 속성명, 데이터 유형, 길이가 정의된 표준을 준수하는지 점검하는 도구

### Usage & Source code explanation

**데이터 표준점검 도구 전체 설명 목차: https://prodtool.tistory.com/23**   
   
1. 데이터 표준점검 도구_1.개요   
https://prodtool.tistory.com/10 [취미로 코딩하는 DA(Data Architect)]  

2. 데이터 표준점검 도구_2.1.화면 구성, 2.2.표준 점검 기능   
https://prodtool.tistory.com/11 [취미로 코딩하는 DA(Data Architect)]  

3. 데이터 표준점검 도구_2.3.표준사전 구성   
https://prodtool.tistory.com/16 [취미로 코딩하는 DA(Data Architect)]  

4. 데이터 표준점검 도구_3.표준점검 케이스   
https://prodtool.tistory.com/18 [취미로 코딩하는 DA(Data Architect)]  

5. 데이터 표준점검 도구_4.별첨   
https://prodtool.tistory.com/22 [취미로 코딩하는 DA(Data Architect)]  


### 배포파일에 포함되어 있는 데이터 표준 사전   
공공데이터 공통표준용어(행정안전부고시 제2020-42호, 2020.8.11. 시행) 자료(hwp)에 포함되어 있는 표준단어, 표준도메인, 표준용어 사전을 사용하였다.   
hwp파일에서 표준단어, 표준도메인, 표준용어를 발췌하여 아래 자료(공공데이터 공통표준_202008.xlsx)에 정리하였다.  
사전 다운로드: [공공데이터 공통표준_202008.xlsx](./%EA%B3%B5%EA%B3%B5%EB%8D%B0%EC%9D%B4%ED%84%B0%20%EA%B3%B5%ED%86%B5%ED%91%9C%EC%A4%80_202008.xlsx)(2021-08-18 추가함)  
출처: https://www.mois.go.kr/frt/bbs/type001/commonSelectBoardArticle.do?bbsId=BBSMSTR_000000000016&nttId=79284  


## 데이터 표준 사전 참고자료 추가 (2021-08-18 추가함)
공공데이터 공통표준용어(행정안전부고시 제2020-66호, 2020.12.10.)에서 추가된 단어, 도메인, 용어를 엑셀파일로 작성하여 업로드하였다.  
hwp파일에서 표준단어, 표준도메인, 표준용어를 발췌하여 아래 자료(공공데이터 공통표준_202012.xlsx)에 정리하였다.  
사전 다운로드: [공공데이터 공통표준_202012.xlsx](./%EA%B3%B5%EA%B3%B5%EB%8D%B0%EC%9D%B4%ED%84%B0%20%EA%B3%B5%ED%86%B5%ED%91%9C%EC%A4%80_202012.xlsx)  
출처: https://www.mois.go.kr/frt/bbs/type001/commonSelectBoardArticle.do?bbsId=BBSMSTR_000000000016&nttId=81609  
      
   
   
## 네이버 국어사전, 영어사전 검색 도구
표준단어, 표준용어의 설명(정의) 내용을 작성하기 위하여 네이버 사전을 검색하는 도구   

### Usage & Source code explanation
1. 네이버 국어사전, 영어사전 검색 도구 사용 방법   
https://prodtool.tistory.com/28 [취미로 코딩하는 DA(Data Architect)]   

2. 네이버 국어사전, 영어사전 검색 도구 동작 방식과 소스코드 설명   
https://prodtool.tistory.com/30 [취미로 코딩하는 DA(Data Architect)]   
   
   
## Word Extractor
현행 속성명으로부터 단어(명사, 복합어)를 추출하여 초기에 표준단어사전을 생성하거나 기존 표준단어사전을 보완하기 위한 도구   
다음과 같은 경우에 도움이 될 것으로 기대한다.   

- 현행 데이터 표준 사전이 없거나 있더라도 표준단어의 개수가 적은 경우   
- 업무가 매우 독특하여 참조하기에 적합한 데이터 표준 사전이 없는 경우   
- Database table, column comment가 너무 방대하여 수작업으로 단어를 추출하는 데 많은 시간이 걸리는 경우   
- 또는 그 반대로 Database table, column comment에 내용이 거의 없어서 표준 단어를 추출하기에 부적하고 업무 매뉴얼 등의 문서에서 추출하는 것이 적합한 경우   
- 그 외, 문서로부터 단어와 빈도 추출이 필요한 경우   

   
실행에 필요한 소스코드, 글꼴, table/column 목록 예시 파일, 출력 예시 파일을 배포용 압축파일로 묶어 두었으니, 이 파일을 다운로드 받으면 된다.   
- 배표용 압축파일: word_extractor.7z   
    
    
### Usage & Source code explanation    
1. 단어 추출 도구(1): 단어 추출 도구 개요 (Word Extraction Tool(1): Overview of Word Extraction Tool)   
https://prodtool.tistory.com/76   

2. 단어 추출 도구(2): 단어 추출 도구 실행환경 구성 (Word Extraction Tool (2): Configure the Word Extraction Tool Execution Environment)   
https://prodtool.tistory.com/77   

3. 단어 추출 도구(3): 단어 추출 도구 실행, 결과 확인 방법 (Word Extraction Tool (3): How to Run Word Extraction Tool, Check Results)   
https://prodtool.tistory.com/78   

4. 단어 추출 도구(4): 단어 추출 도구 소스코드 설명(1) (Word Extraction Tool(4): Word Extraction Tool Source Code Description(1))   
https://prodtool.tistory.com/79   

5. 단어 추출 도구(5): 단어 추출 도구 소스코드 설명(2) (Word Extraction Tool(5): Word Extraction Tool Source Code Description(2))   
https://prodtool.tistory.com/80   

6. 단어 추출 도구(6): 단어 추출 도구 부가 설명 (Word Extraction Tool (6): Additional Description of Word Extraction Tool)   
https://prodtool.tistory.com/81   

7. 단어 추출 도구 설명글 목록, 목차, 다운로드 (Word Extraction Tool Description article list, table of contents, download)   
https://prodtool.tistory.com/82   

