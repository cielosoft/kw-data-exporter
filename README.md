# Kingdom Watch Data Exporter

[![Build Status](https://travis-ci.org/cielosoft/kw-data-exporter.svg?branch=master)](https://travis-ci.org/cielosoft/kw-data-exporter)

## Install

```
go get -u github.com/cielosoft/kw-data-exporter
```

## Usages

```
kw-data-expoprter --help
```

- no-csv: csv 포맷 사용 안함
- json: json 포맷 사용
- sql: sql 포맷 사용
- all: json, sql 포맷 사용

## Rules
- #표시로 시작하는 경우 주석
- A1: 추출 포맷 (예: #JSON)
  - JSON: json 형식
  - SQL: SQL 쿼리 스트크립트
  - KeyValue: 키 밸류 타입
- B1: 이름
- A2~: 필드 이름 (예: #level)
- A3~: CSV 전용 필드 이름 (예: #level)
- A4~: 필드 데이터 타입 (예: #string)
  - string: 문자열
  - float: 실수
  - int: 정수 (기본)
- 날짜 형식은 지원하지 않습니다 (반드시 텍스트로 형식으로 지정 해야 함)
