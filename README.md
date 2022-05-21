# Webpage-Cralwer-AND-Excel-VBA
## Webpage 에서의 표를 자동 추출하고 원하는 형태로 표 변환 😎

![ezgif com-gif-maker (6)](https://user-images.githubusercontent.com/99319638/169628370-1af5bdd7-6727-4b27-91b3-6b9784ca0dd4.gif)

```
구간도 많고 날짜도 마찬가지.
원하는 데이터가 많기에 일일이 수동으로 [ctrl+c] [ctrl+v] 하는 것엔 무리가 있을 것으로 판단 🙄
자동으로 크롤링 할때 필요한 코드 정리 & 원하는 데이터의 형태로 형변환
```

<br/>

# SUMMARY

<br/>
<br/>

![테이블_크기변경(2)](https://user-images.githubusercontent.com/99319638/169635287-1a43ead5-322c-4bdf-bbd5-cbedf76219d5.png)

<details>
<summary>접기/펼치기 버튼</summary>
<div markdown="1">

### [1] 웹페이지상에서의 날것의 표
![구간별통계정보 _ 통계정보 _ 교통정보 경기도교통정보센터 - Chrome 2022-05-21 오전 9_45_48](https://user-images.githubusercontent.com/99319638/169634121-ec16f4ee-1b74-48c9-b633-aa90c3c86dec.png)

### [2] 해당 코드를 실행하고 난뒤..
![image](https://user-images.githubusercontent.com/99319638/169628732-64f5167f-13e4-46a5-8fea-7b0ca8f74b1d.png)

</div>
</details>

# table of contents
- Webpagecrawler_final.ipynb 코드로 csv 파일 저장
- Excel VBA.ipynb 매크로 이용 파일 형 변환
- 통합문서 하나로 취합
- Finish 👻

<br/>

# Webpagecrawler_final.ipynb 코드로 csv 파일 저장

```
https://gits.gg.go.kr/web/trafficInfo/webSectionStatistics.do?linename=101&linename2={}'.format(j)+
'&lineway=1&linedate=2020-07-{}'.format(i)+'&speedGraph=07
```

`linename2` 포맷팅
그리고 `linedate` 포맷팅

<br/>

![image](https://user-images.githubusercontent.com/99319638/169627659-4ca63a21-21b1-4abe-adbb-6be1d7c0bcdc.png)
![image](https://user-images.githubusercontent.com/99319638/169634757-3cca1a9a-1583-402c-a661-86cdc4b269c8.png)

## 파일 내용 😋

![image](https://user-images.githubusercontent.com/99319638/169634863-676d592e-1264-43ec-b113-2ba5b6b16f5b.png)

<br/>

# Excel VBA.ipynb 매크로 이용 파일 형 변환

(Excel VBA가 모듈에 저장되어 있는 상태에서) <br/>
`파일 하나 열고` -> `개발도구` -> `삽입` -> `단추` -> `단추 삽입` -> `매크로 지정` -> `Module2.매크로1` -> `버튼 클릭 🐠` <br/>
-> `한 폴더안에 있는 파일들 모두 형변환됨` -> /`통합문서 취합`

### **[실행]**
1. 복사 붙여넣기 (transpose)
![image](https://user-images.githubusercontent.com/99319638/169635618-d61bed11-5d8e-4377-899d-3032a8f554f6.png)

2. 테이블 아래 아래 셀에 
`=INDEX($B$53:$AQ$76,MOD(ROW(A1)-1,24)+1,QUOTIENT(ROW(A1)-1,24)+1)` <br/>
그리고 복사 붙여넣기 (값만) <br/>
![image](https://user-images.githubusercontent.com/99319638/169635657-35827c5a-064b-4774-ada6-140f1423bf4f.png)

참고로 파일안의 폴더를 열고, 1번 2번 실행시키고 하는 것 모두 <br/>
매크로로 자동으로 돌아가게 된다. 매크로이기는 하지만 클릭을 내가 몇 번 해야한다는 번거로움이 있다. <br/>
기회가 된다면 `클릭한번으로 돌아가는 프로그램`도 만들어봐야겠다............ 🔥

