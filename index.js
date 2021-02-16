const puppeteer = require("puppeteer");
const fs = require("fs");
const xlsx = require("xlsx");
/*
교양기초는 학정번호가 너무 제멋대로다. 수작업으로 할것.
대학교양도 좀 다양한데. 
그외는 x000~x260. 분반은 00~05. 소분반은 00~02.

상경 경영대
    ECO : 경제학
    STA : 통계학
    BIZ : 경영학
문과대
    HUM : 뭔지 모르겠음
    TTP : 뭔지 모르겠음
    KOR : 국어국문
    CLL : 중어중문
    ELL : 영어영문
    GER : 독어독문
    FRE : 불어불문
    RUS : 노어노문
    HIS : 사학
    PHI : 철학
    LIS : 문헌정보학
    PSY : 심리학
이과대
    SCI : 이과대 공통
    MAT : 수학 및 수학공통기초 과목들
    PHY : 물리학
    CHE : 화학
    ESS : 지구시스템
    AST : 천문우주
    ATM : 대기과학
공과대
    ENG : 공대 공통
    MEU : 기계공학
    EEE : 전기전자
    CSI : 컴퓨터과학        
    MST : 신소재공학
    CEE : 건설환경공학
    CRP : 도시공학
    ARC : 건축공학
    DAA : 화공생명공학
    IIE : 산업공학
    IIT : 글로벌융합공학
생명시스템
    BIO : 생명대 공통
    BCH : 생화학
    BTE : 생명공학
    LSB : 잘 모르겠음
신과대 
    THE : 신과대
사회과학대
    SOC : 사회학
    TTP : 여기도 TTP가 있음. 뭔지 잘 모르겠음
    POL : 정치외교학
    PUB : 행정학
    COM : 언론홍보영상
    SOW : 사회복지학
    ANT : 문화인류학
음악대
    CMP : 작곡
    MSC : 뭔지 잘 모르겠음.
    TTP : 여기도 TTP 가 있음
    CHM : 교회음악
    BSP : 피아노
    VOM : 성악
생활과학대
    CNT : 의류환경학
    TTP : 여기도 있음. TTP가 교직이수이구나.
    FNS : 식품영양학
    HID : 실내건축학
    CFS : 아동가족학
    DSN : 생활디자인
언더우드국제대 : 기존학과와 학정번호 겹치는 것들이 있기에, 중복되는건 기록안함.
    UIC
    CLC
    ISM
    SDC
    LST
    JCL
    QRM
    STP
    SED
    NSE
    IID
    UBC
    ASP
    CDM
    CTM
글로벌인재
    GIC
    GKE
    GCM
    GBL
    GAI
    GLC
    UCC
간호대
    NUR
의과대
    MED
치과대
    DEN
교양기초(19학번~)
    YCA : 기독교 수업
    YCB : 글쓰기
    YCC : 대학영어
대학교양(19학번~)
    UCB : 
    UCE :
    UCJ
    YCD
    UCL
    YCE
    UCK
    YCF : 외국어들
    YCG    
    UCI
    YCI
    UCF
    UCG
    UCH
    YCS
*/
/*
    ["ECO",
    "STA",
    "BIZ",
    "HUM",
    "TTP",
    "KOR",
    "CLL",
    "ELL",
    "GER",
    "FRE",
    "RUS",
    "HIS",
    "PHI",
    "LIS",
    "PSY",
    "SCI",
    "MAT",
    "PHY",
    "CHE",
    "ESS",
    "AST",
    "ATM",
    "ENG",
    "MEU",
    "EEE",
    "CSI",
    "MST",
    "CEE",
    "CRP",
    "ARC",
    "DAA",
    "IIE",
    "IIT",
    "BIO",
    "BCH",
    "BTE",
    "LSB",
    "THE",
    "SOC",
    "TTP",
    "POL",
    "PUB",
    "COM",
    "SOW",
    "ANT",
    "CMP",
    "MSC",
    "TTP",
    "CHM",
    "BSP",
    "VOM",
    "CNT",
    "TTP",
    "FNS",
    "HID",
    "CFS",
    "DSN",
    "UIC",
    "CLC",
    "ISM",
    "SDC",
    "LST",
    "JCL",
    "QRM",
    "STP",
    "SED",
    "NSE",
    "IID",
    "UBC",
    "ASP",
    "CDM",
    "CTM",
    "GIC",
    "GKE",
    "GCM",
    "GBL",
    "GAI",
    "GLC",
    "UCC",
    "NUR",
    "MED",
    "DEN",]

교양기초
    YCA : 기독교 수업
    YCB : 글쓰기
    YCC : 대학영어
대학교양
    UCB : 
    UCE :
    UCJ
    YCD
    UCL
    YCE
    UCK
    YCF : 외국어들
    YCG    
    UCI
    YCI
    UCF
    UCG
    UCH
    YCS
*/
// 추출한 HTML 텍스트에서 정규식을 이용해 필요한 내용을 parse 할 것임.
const curriculumRegexString = /<td align="" class="BoxText_1" colspan="3"><pre>(.*?)<\/pre>/gims;
const attendanceRegexString = /<td class="BoxText_1_C" align="center">.*<\/td>/gi;

const returnLowerValue = (inputA, inputB) => {
  if (inputA > inputB) {
    return inputB;
  } else {
    return inputA;
  }
};

const curriculumAbbreviationArray = [
  "ECO",
  "STA",
  "BIZ",
  "HUM",
  "TTP",
  "KOR",
  "CLL",
  "ELL",
  "GER",
  "FRE",
  "RUS",
  "HIS",
  "PHI",
  "LIS",
  "PSY",
  "SCI",
  "MAT",
  "PHY",
  "CHE",
  "ESS",
  "AST",
  "ATM",
  "ENG",
  "MEU",
  "EEE",
  "CSI",
  "MST",
  "CEE",
  "CRP",
  "ARC",
  "DAA",
  "IIE",
  "IIT",
  "BIO",
  "BCH",
  "BTE",
  "LSB",
  "THE",
  "SOC",
  "TTP",
  "POL",
  "PUB",
  "COM",
  "SOW",
  "ANT",
  "CMP",
  "MSC",
  "TTP",
  "CHM",
  "BSP",
  "VOM",
  "CNT",
  "TTP",
  "FNS",
  "HID",
  "CFS",
  "DSN",
  "UIC",
  "CLC",
  "ISM",
  "SDC",
  "LST",
  "JCL",
  "QRM",
  "STP",
  "SED",
  "NSE",
  "IID",
  "UBC",
  "ASP",
  "CDM",
  "CTM",
  "GIC",
  "GKE",
  "GCM",
  "GBL",
  "GAI",
  "GLC",
  "UCC",
  "NUR",
  "MED",
  "DEN",
];
let curriculumNumberArray = [];
for (let num = 1000; num < 1260; num++) {
  curriculumNumberArray.push(num);
}
for (let num = 2000; num < 2260; num++) {
  curriculumNumberArray.push(num);
}
for (let num = 3000; num < 3260; num++) {
  curriculumNumberArray.push(num);
}
for (let num = 4000; num < 4260; num++) {
  curriculumNumberArray.push(num);
}

const bbArray = ["00", "01", "02", "03", "04", "05"];
const sbbArray = ["00", "01", "02"];

const parseWeb = async () => {
  const browserOption = await {
    //slowMo: 500, // 디버깅용으로 퍼핏티어 작업을 지정된 시간(ms)만큼 늦춥니다.
    headless: true, // 디버깅용으로 false 지정하면 브라우저가 자동으로 열린다.
  };

  const browser = await puppeteer.launch(browserOption);
  const curriculumPage = await browser.newPage();

  const attendanceNumberPage = await browser.newPage();
  await attendanceNumberPage.goto(
    "http://ysweb.yonsei.ac.kr:8888/curri120601/curri_new.jsp#top"
  );

  for await (let abbreviation of curriculumAbbreviationArray) {
    let exportArray = await [];
    for await (let i of curriculumNumberArray) {
      for await (let bb of bbArray) {
        for await (let sbb of sbbArray) {
          try {
            const curriculumPageURL = await `http://ysweb.yonsei.ac.kr:8888/curri120601/curri_pop2.jsp?&hakno=${abbreviation}${i}&bb=${bb}&sbb=${sbb}&domain=H1&startyy=2020&hakgi=1`;
            await curriculumPage.goto(curriculumPageURL, {
              waitUntil: "networkidle0",
            });

            const curriculumBodyHTML = await curriculumPage.evaluate(
              () => document.body.innerHTML
            );

            let dirtyParsedDataRow = await curriculumBodyHTML.match(
              curriculumRegexString
            );

            //await console.log("parsedDataRow: ", dirtyParsedDataRow);

            let parsedDataRow = await dirtyParsedDataRow.map((element) => {
              return element
                .replace(`<td align="" class="BoxText_1" colspan="3"><pre>`, "")
                .replace(`</pre>`, "");
            });

            //await console.log("trimedParsedDataRow: ", parsedDataRow);

            //커리큘럼 스크래핑이 끝났으면 학정번호 정보를 array에 넣어준다.
            await parsedDataRow.push(abbreviation);
            await parsedDataRow.push(i);
            await parsedDataRow.push(bb);
            await parsedDataRow.push(sbb);

            // 수강인원을 알아내기 위해 수강신청 결과 popup 을 처리할 로직.
            try {
              const newPagePromise = new Promise((x) =>
                browser.once("targetcreated", (target) => x(target.page()))
              );
              await attendanceNumberPage.evaluate(
                async ([abbreviation, i, bb, sbb]) => {
                  await OpenList_mileage_result(
                    "H1",
                    "20201",
                    `${abbreviation}${i}`,
                    `${bb}`,
                    `${sbb}`
                  );
                },
                [abbreviation, i, bb, sbb]
              );

              const newPage = await newPagePromise;

              // await console.log("newPage URL: ", newPage.url());
              const attendanceBodyHTML = await newPage.evaluate(
                () => document.body.innerHTML
              );
              const parsedAttendanceData = await attendanceBodyHTML.match(
                attendanceRegexString
              );

              // 수강신청인원 데이터에서 7번째 요소는 해당 수업의 수강 정원, 8번째는 수강신청한 사람 숫자.
              // 0부터 시작하니, 코드에선 6, 7.

              const attendanceMax = await Number(
                parsedAttendanceData[6]
                  .replace(`<td class="BoxText_1_C" align="center">`, "")
                  .replace(`</td>`, "")
              );
              const attendanceAplicant = await Number(
                parsedAttendanceData[7]
                  .replace(`<td class="BoxText_1_C" align="center">`, "")
                  .replace(`</td>`, "")
              );

              //수강인원을 parsedDataRow 에 넣어준다.
              await parsedDataRow.push(
                returnLowerValue(attendanceAplicant, attendanceMax)
              );
            } catch (error) {
              // console.log(error);
            }

            // export용 배열에 parse 한 내용을 push. 2차원 배열을 만든다.
            await exportArray.push(parsedDataRow);
          } catch (error) {
            // console.log(error);
            continue;
          }
        }
      }
      await console.log(`I'm working! ${abbreviation}${i}: Done`);
    }
    // 해당 과목에 대해 loop 다 끝났으면 파일 하나로 export.
    // await console.log(exportArray);
    const workSheet = await xlsx.utils.aoa_to_sheet(exportArray);
    const workBook = await xlsx.utils.book_new();
    await xlsx.utils.book_append_sheet(workBook, workSheet, "temp");
    await xlsx.writeFile(workBook, `${abbreviation}_${Date.now()}` + ".xlsx");
  }
  await browser.close();
};

parseWeb();
