# 워드 정리 스킬 설정

## 호출 조건
- 사용자가 "워드 정리"를 입력하면 즉시 아래 흐름을 실행

## 동작 흐름

```
텔레그램 입력: "워드 정리"
  → 봇 응답: "워드 정리가 필요한 파일을 첨부해주세요~"
  → 사용자가 파일 전송 (.docx / .html / .txt)
  → word_organizer.py 실행
      1. 파일 내용 추출 (포맷별 파서)
      2. 제목 + 섹션 + 불릿 구조로 파싱
      3. content.json 생성
      4. sds-word-writer generate.py 호출 → 새 Word 파일 생성
  → 생성된 Word 파일을 텔레그램으로 전송
```

## 워드 정리 실행 트리거 (핵심 규칙)

"워드 정리" 입력 시 아래 순서를 반드시 완전 자동 실행.

### 실행 순서
1. 파일 수신 대기 상태로 전환 후 안내 메시지 전송
2. 사용자가 파일 전송 시 word_organizer.py 호출

```python
import sys; sys.path.insert(0, "/path/to/ssg-company-analyst")
from word_organizer import organize
organize(input_path, output_path)
```

3. 결과: sds-word-writer 스타일 Word 파일 자동 생성 → 텔레그램 전송

## 파이프라인 구성
- 변환 로직   : word_organizer.py
- Word 생성기 : skills/sds-word-writer/scripts/generate.py
- 파일명 규칙 : {원본파일명}_정리.docx

## 지원 파일 포맷

| 확장자 | 처리 방식 |
|--------|-----------|
| `.docx` | python-docx로 단락·스타일 추출, 볼드 단락은 헤딩으로 인식 |
| `.html` / `.htm` | BeautifulSoup으로 h1~h6(헤딩), p/li/td(본문) 추출 |
| `.txt` | 줄별 파싱, 번호/기호/대문자 패턴으로 헤딩 자동 감지 |

## 출력 Word 파일 스타일

| 항목 | 값 |
|------|----|
| 폰트 | 바탕체 |
| 장체 | 95% |
| 본문 크기 | 14pt |
| 줄간격 | single (line=240) |
| 섹션 기호 | Ⅰ Ⅱ Ⅲ … (로마 숫자) |
| 불릿 종류 | dash (- ) 기본, dot (·) 서브 항목 |
| 페이지 번호 | 하단 중앙 |
| 마감 | - 이  상 - (우측 정렬) |

## 트리거 패턴 (bot.py)

```python
_WORD_ORG_RE = re.compile(r"워드\s*정리", re.IGNORECASE)
```
