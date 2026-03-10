# 3D Plot Dashboard Pipeline

엑셀(`.xlsx`, `.xls`), TSV, CSV 파일을 입력받아 브라우저에서 바로 열 수 있는 인터랙티브 3D plot 대시보드 `.html`을 생성하는 파이프라인입니다.

핵심 목표:

- 빠른 로딩: CSV/TSV는 `polars` 기반으로 처리
- 가벼운 결과물: 단일 HTML로 생성, 대용량은 자동 샘플링
- 메모리 효율: 숫자형 다운캐스트, 카테고리 사전화, 스트리밍 성격의 Excel 샘플링
- 크로스플랫폼: Ubuntu, macOS, Windows 11에서 동일하게 동작

## 기능

- 입력 지원: `.csv`, `.tsv`, `.txt`, `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`
- 3D 축 매핑: `x`, `y`, `z` 컬럼 선택
- 스타일 매핑: 색상(`--color`), 크기(`--size`), 라벨/검색(`--label`)
- 인터랙션: 회전, 줌, hover, point click 상세 보기
- 탐색 기능: 축 범위 필터, 카테고리 필터, 텍스트 검색
- 내보내기: 현재 보이는 포인트만 CSV로 다시 export
- Figure급 출력: 고해상도 3D PNG export, 2D 사분면 PNG(2x2 패널) export
- 테마/스타일: Nebula, Paper Figure, Cartoon, Molstar-like + point style 3종
- Figure preset: Default/Nature/Cell/NeurIPS 스타일 프리셋
- Unsupervised clustering: 브라우저 내장 K-means(k=2~12), 클러스터별 색상/범례 표시
- 대규모 클러스터 가속: 샘플 학습 + 전체 포인트 할당, 클러스터 중심점(centroid) 시각화
- Cluster analytics: train count, inertia, silhouette(sample), cluster size
- Undo/Redo + Snapshot 저장/복원
- Linked 2D brushing: 2D 브러시로 3D 연동 필터링
- Session preset export/import(JSON)
- 카메라 저장: 브라우저 localStorage 기반

숫자 컬럼이 3개 미만이면 파이프라인이 `__row_index`, `__code_<column>` 같은 보조 숫자 컬럼을 자동 생성해서 3D 탐색이 끊기지 않게 합니다.

## 설치

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -e .
```

Windows PowerShell:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e .
```

## 사용법

기본:

```bash
3dplot-dashboard data.csv
```

출력 파일 지정:

```bash
3dplot-dashboard data.tsv -o out/dashboard.html
```

축/스타일 초기값 지정:

```bash
3dplot-dashboard sales.xlsx \
  --sheet 0 \
  --x longitude \
  --y latitude \
  --z revenue \
  --color region \
  --size profit \
  --label customer_name
```

샘플링 상한 조정:

```bash
3dplot-dashboard huge.csv --max-points 150000
```

생성 후 브라우저 열기:

```bash
3dplot-dashboard data.csv --open
```

## CLI 옵션

```text
3dplot-dashboard INPUT
  -o, --output PATH
  --sheet SHEET
  --x COLUMN
  --y COLUMN
  --z COLUMN
  --color COLUMN
  --size COLUMN
  --label COLUMN
  --delimiter DELIMITER
  --max-points N
  --seed N
  --title TITLE
  --plotly-url URL
  --open
```

## 성능 전략

- CSV/TSV: `polars.scan_csv()`로 읽고 행 수가 많으면 systematic sampling 후 collect
- XLSX/XLS: read-only/온디맨드 방식으로 읽으면서 reservoir sampling
- float는 `Float32`, int는 가능한 작은 dtype으로 다운캐스트
- low-cardinality 문자열 컬럼은 dictionary encoding으로 HTML 크기 감소
- 클러스터링은 현재 보이는 포인트에만 적용하여 브라우저 연산량을 제한
- 클러스터링은 좌표 정규화 후 수행해 축 스케일 편향을 줄이고, 대용량은 샘플 기반 학습으로 가속
- 필터/축 범위/검색이 바뀌지 않은 렌더는 가시성 인덱스를 재사용해 상호작용 반응성 개선

## 오프라인 사용

기본 HTML은 Plotly CDN URL을 사용합니다. 완전 오프라인 환경이면 로컬 `plotly.min.js` 경로를 준비한 뒤 `--plotly-url`로 넘기면 됩니다.

예시:

```bash
3dplot-dashboard data.csv --plotly-url ./vendor/plotly-2.35.2.min.js
```

## 샘플 데이터 생성

```bash
.venv/bin/python scripts/generate_sample_data.py
```

생성 위치:

- `examples/cluster_demo.csv`
- `examples/cluster_demo.tsv`
- `examples/cluster_demo.xlsx`

## 테스트

```bash
.venv/bin/pip install -e .[dev]
.venv/bin/pytest
```
