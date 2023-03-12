# fulfill

- python version = 3.10.x
- only work in windows environment

## Quick Start

1. git clone -b warehouse --single-branch https://github.com/Themath93/fulfill.git
2. 브랜치 warehouse
3. 파이썬 다운로드 www.python.org  ver : 3.10.10
4. py.ext -m pip install -r requirements.txt
    - cmd 혹은 Windows PowerShell에서 git clone한 폴더로 진입후 위 명령어 실행
5. 제한된 보기해제 옵션 보안센터 제한된 보기해제
6. **xlwings.xlam** 파일 매크로포함으로 실행
    - **매크로 사용을 위한 매크로 방지 해제**
7. 따로 다운 받은 instantclient_fulfill.zip은 fulfill 폴더안에 압축풀기
    - zip파일은 삭제
8. 작업자가 사용할 cytiva_worker.xlsm 파일은 바탕화면에 위치
    - **해당 파일 다운로드 후 매크로 사용을 위한 매크로 방지 해제**
9. ALT + F11로 개발자도구 진입 후 도구 - 참조 - xlwings 클릭으로 활성화
10. xlwings 메뉴가 생겼으면 클릭
    - 왼쪽에 PYTHONPATH: 옆에 빈칸이 있음
    - 아래 내용 복사하는데 "" 안에는 **본인 컴퓨터계정이름**을 적고 넣으면됨
    - C:\Users\"컴퓨터계정이름"\Desktop\fulfill\xlwings_job
11. 엑셀의 BEGIN 시트에서 **START버튼** 을 누르면 FORM이 불러와 지면서 작업 시작가능


## 서비스 시작작업
1. PrintPreview() - > print() 실제 print()로 변경
2. BRACNH DB의 EMAILS 컬럼에 실제 email 값 넣어 주기 " ;".join(list)로 받을 것!
3. 기존 TEMP 데이터 지워야햐는 테이블
    - DW
        - SERVICE_REQUEST
        - SO_OUT
        - SYSTEM_STOCK
    - DM 모든테이블


4. 기존 TEMP 데이터 지우고 새로운 DATA를 넣어줘야하는테이블
    - DW
        - BRANCH
        - LOCAL_LIST
        - PROD_POSE
        - SERVICE_REQUEST
        - SHIPMENT_INFORMATION
        - SVC_BIN
        - SVC_TOOL
    - WEB 사용자 재정의
    


**주의사항**
1. 엑셀작업시 셀안에 임의로 넣는 값은 저장을 하더라도 저장되지 않음
    - 시트마다 Cell change나 Data Input 기능등이 있는데 해당 기능을 사용하여야만 데이터베이스에 적용하여 이용가능

2. WMS 사용할 때 git clone폴더는 바탕화면에 fulfill폴더안에 위치할 것
3. cytiva_worker.xlsm 파일은 바탕화면에 위치할것
    - 해당 파일 파일명변경 절대금지



PatchNote 230309 warehouse ver 1.12 
1. CODE
    1. INboundReady
        - Fatfinger 방지를 위한 컨펌기능추가
    2. Shipready 
        - 로직 기능 개선 및 속도향상
2. EXCEL
    1. Cytiva.xlsm
        - SHIPMENT_INFORMAITON 시트의 STATUS 셀 크기조정 60 -> 48
3. USERGUIDE
    1. 유저가이드 PPT 생성 
        - cytitva_worker.xlsm 에서 userguide 버튼 클릭하면 실행 가능

PatchNote 230309 warehouse ver 1.134
1. CODE 
    1. 서비스 출고 반납 관련 상태변경 메일 발송 기능 추가
    2. Version 업데이트시 버천 체크자동으로 하여 업데이트 기능 추가
PatchNote 230309 warehouse ver 1.136
1. CODE
    1. Version 업데이트 방식 변경 및 업데이트 실패 시 이전 Version으로 작업 진행