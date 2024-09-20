# process_attendance
## 파일 입력 형식
1. user_info 폴더 생성
   1. 참가자 정보 엑셀 파일 저장
   2. filename: users.xlsx
   3. 필수 열: 이름+전화번호 뒤 4자리, 수료 여부
2. attendance_forms 폴더 생성
   1. 구글 폼 엑셀 파일 저장
   2. filename: N차 출결.xlsx
   3. 프로그램 실행 시 엑셀 파일 자동 생성됨
## 개발 환경 설치
1. python 3.12.0 설치: [official python webpage](https://www.python.org/downloads/)
2. pipenv 설치: ```pip install pipenv```
3. 개발 환경 설치: ```pipenv install```

## 프로그램 실행
1. 가상 환경 쉘 실행
   1. ``pipenv shell``
   2. ``python process_roll_book.py --{attendance 회차} | --results``
2. 가상 환경 바로 실행
   1. ```pipenv run python process_roll_book.py  --{attendance 회차} | --results```

## kwargs
1. --attendance 숫자: 숫자 회차에 해당하는 파일을 처리
2. --results: 처리된 모든 출결 파일을 가져와서 최종 결과 파일 생성