# Product Requirements Document (PRD): Lightweight Tetris

## 1. 개요 (Overview)
본 문서는 `demotetris` 폴더에 구현될 '가장 가벼운 형태의 테트리스(Lightweight Tetris)' 게임의 제품 요구사항을 정의합니다. 목적은 복잡한 프레임워크나 외부 라이브러리 없이 순수 HTML, CSS, JavaScript(Vanilla JS)만을 사용하여 빠르고 가볍게 동작하는 테트리스 게임을 구축하는 것입니다.

## 2. 목표 (Goals)
- **경량화**: 외부 종속성(Dependencies) 최소화. 단일 HTML 파일 또는 최소한의 파일(HTML, CSS, JS 각 1개)로 구성.
- **직관성**: 테트리스 본연의 핵심 메커니즘(블록 하강, 회전, 줄 없애기, 점수 획득)에 집중.
- **반응성**: 부드러운 애니메이션과 즉각적인 키보드 입력 반응.

## 3. 핵심 기능 (Core Features)
### 3.1. 게임 보드 (Game Board)
- 10(가로) x 20(세로) 그리드 시스템.
- HTML Canvas API 또는 단순 DOM 요소(div)를 활용한 렌더링.

### 3.2. 블록 (Tetrominoes)
- 7가지 기본 테트로미노 제공 (I, J, L, O, S, T, Z).
- 각 블록별 고유 색상 적용.
- 회전(Rotation) 및 충돌 감지(Collision Detection) 로직.

### 3.3. 게임 플레이 (Gameplay)
- **입력 (Controls)**:
  - `좌/우 화살표`: 블록 이동.
  - `위 화살표`: 블록 회전.
  - `아래 화살표`: 소프트 드롭 (빠르게 내리기).
  - `스페이스바`: 하드 드롭 (바닥으로 즉시 내리기).
- **진행 (Progression)**:
  - 일정 시간마다 블록 자동 하강.
  - 가로 줄이 꽉 차면 줄 삭제 및 점수 획득.
  - 삭제된 줄 수에 따라 레벨/속도 증가.

### 3.4. 게임 상태 관리 (State Management)
- **시작/일시정지**: 게임 시작 및 일시정지 기능.
- **게임 오버**: 새 블록이 생성될 위치에 이미 블록이 있으면 게임 오버.
- **점수 표시**: 현재 점수 및 최고 점수(Local Storage 활용 선택) UI 표시.

## 4. 기술 스택 (Tech Stack)
- **구조**: HTML5 (`index.html`)
- **스타일**: CSS3 (`style.css` - 깔끔하고 직관적인 디자인 적용)
- **로직**: Vanilla JavaScript (`script.js` 또는 `app.js` - ES6+ 문법 사용)

## 5. UI/UX 요구사항
- 심플하고 모던한 다크 테마(Dark Theme) 또는 미니멀 테마 적용.
- 보드 옆에 다음 블록(Next Block) 및 점수(Score) 패널 배치.
- 불필요한 화려한 그래픽 배제, 핵심 플레이 경험에 집중.
