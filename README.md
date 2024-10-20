# 프로젝트 구동 방법
1. node rfp-export: 2차 제안서(ppt) 내보내기
2. node rrmse-export: 1차 제안서(엑셀), dr history cbl 내보내기

# 프로젝트 배포 방법
1. nodejs 실행
2. docker 접근할때 exec로 들어가서 각각 다른 프로세스로 띄워야함

# 프로젝트 구조 설명
1. token.json: Google API을 사용하기 위한 token key 파일로, SalesTool Google 계정 인증을 완료하면 자동으로 생성됨. 삭제하면 rfp-export 구동할때 질의가 나오고 재기동하면 정상적으로 실행됨
2. rfp-sample.js: 2차 제안서 API Response 규격. 테스트용으로 만들어 놓은 샘플 데이터
3. templete.xlsx: Excel Chart을 만들기 위한 템플릿 엑셀 파일. rrmse-export에서 사용중
4. cbl-sample/js: HR 히스토리 Cbl 전체 API Response 규격. 테스트용으로 만들어 놓은 샘플 데이터

