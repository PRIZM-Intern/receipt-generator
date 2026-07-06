# 월별 영수증 기안서 생성기

PDF/JPG/PNG 영수증을 업로드하면 원본 Word 양식에 맞춰 DOCX와 PDF를 생성하는 Streamlit 앱입니다.

## 주요 기능

- PDF 내장 텍스트에서 거래일시와 실제 결제금액 일부 자동 인식
- 사용자가 날짜와 인정금액을 검수·수정
- 거래일시 순으로 영수증 정렬
- 영수증 한 장당 Word 한 페이지로 배치
- DOCX 다운로드
- LibreOffice가 설치된 환경에서는 PDF도 함께 다운로드

업로드 파일과 생성 파일은 서버에 영구 저장하지 않습니다.

## 로컬 실행

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

## Streamlit Community Cloud 무료 배포

1. 이 폴더의 파일을 회사가 관리하는 GitHub 저장소에 올립니다.
2. <https://share.streamlit.io>에 로그인합니다.
3. `Create app`을 선택합니다.
4. GitHub 저장소와 브랜치를 선택합니다.
5. Main file path에 `app.py`를 입력합니다.
6. 배포 후 발급된 `https://...streamlit.app` 주소를 공유합니다.

`packages.txt`에 LibreOffice와 나눔글꼴이 지정되어 있어 배포 시 PDF 생성 환경도 함께 설치됩니다.

## 운영·인수인계

- 개인 저장소 대신 회사 GitHub 조직의 저장소를 사용합니다.
- 저장소 관리자를 최소 두 명 지정합니다.
- 담당자가 변경되면 GitHub 관리자 권한과 Streamlit 관리 권한을 함께 인계합니다.
- 앱을 공개하면 누구나 링크로 사용할 수 있습니다. 공개 전 회사의 문서 처리 정책을 확인합니다.
- 라이브러리를 업데이트한 뒤에는 샘플 영수증으로 DOCX/PDF 생성 결과를 확인합니다.

## 문서 형식

- A4, 상단 3.00cm, 하단·좌우 2.54cm
- Header: `영 수 증`, 맑은 고딕 22pt Bold, 중앙 정렬
- 영수증 이미지 최대 폭 9.70cm, 최대 높이 20.10cm
- 하단: 총 장수, 부서명, 한글 본명
