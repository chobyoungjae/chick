# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요
Google Apps Script 기반의 Google Sheets 자동화 도구 모음입니다. 주로 스프레드시트 데이터 처리 및 관리를 위한 스크립트들로 구성되어 있습니다.

## 파일 구조
- `Code.js`: A열 기준 셀 병합 기능
- `합치기.js`: A열 값을 기준으로 중복 행 데이터를 합치고 정렬하는 기능
- `sample_test.js`: 중복값 합치기 메인 기능 - A열 기준으로 중복 행의 숫자 데이터를 합계 처리
- `출고완료.js`: 출고완료 상태 관리 및 백업시트 연동 시스템

## 핵심 기능 구조

### 중복값 처리 시스템 (`sample_test.js`)
- **메인 함수**: `mergeDuplicateValues()` - A열 기준 중복 행 합계 처리
- **UI 메뉴**: `onOpen()` - Google Sheets 실행 시 "중복제거" 메뉴 자동 생성
- **처리 범위**: D열부터 "예약리스트" 헤더 전 열까지의 숫자 데이터
- **헤더 기준**: 6행에서 "예약리스트" 헤더 위치를 동적으로 찾아 범위 설정
- **데이터 시작**: 7행부터 실제 데이터 처리

### 출고완료 관리 시스템 (`출고완료.js`)
- **메인 프로세스**: `processCheckout()` - 출고완료 상태를 백업시트에 동기화
- **날짜 검증**: A1(시작일)과 A2(종료일)이 같은 날짜여야 실행
- **시트 연동**: "프론트앤드" 시트와 "일별 발주량 백업본" 시트 간 데이터 동기화
- **트리거 시스템**:
  - `onEdit()` - 셀 편집 시 자동 실행
  - `onChange()` - 시트 구조 변경 감지
  - 시간 기반 트리거 - 1분마다 체크박스 상태 확인
- **동시성 제어**: `LockService`로 중복 실행 방지

## Google Apps Script 환경
- Google Sheets와 밀접하게 연동된 서버리스 스크립트
- SpreadsheetApp, LockService, ScriptApp 등 GAS 전용 API 사용
- 트리거 기반 자동화 시스템으로 실시간 데이터 처리

## 개발 시 고려사항
- Google Apps Script 환경에서만 실행 가능
- 트리거는 Google Apps Script 에디터에서 수동으로 설정
- 권한 초기화는 `initializePermissions()` 함수로 최초 1회 실행 필요
- 스프레드시트 시트명은 하드코딩되어 있음 ("프론트앤드", "일별 발주량 백업본")