# Kingdom Watch Data Exporter


## Install

```
go get -u github.com/cielosoft/kw-data-exporter
```

## Rules
- #표시로 시작하는 경우 주석
- A1: 추출 포맷 (예: #JSON)
  - JSON: json 형식
  - SQL: SQL 쿼리 스트크립트
  - PROTOBUF: Protocol Buffers 형식
  - #!JSON 등 앞에 ! 를 붙일시 서버 전용으로 csv 파일은 추출되지 않습니다
- A2~: 필드 이름 (예: #level)
- A3~: 필드 데이터 타입 (예: #string)
  - string: 문자열
  - float32: 실수
  - int8, uint8, int16, uint16, int32, uint32: 정수
  - 타입을 생략 하면 csv 파일에서만 추출 됩니다
- 날짜 형식은 지원하지 않습니다 (반드시 텍스트로 형식으로 지정 해야 함)