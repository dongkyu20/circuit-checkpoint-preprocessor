# 회로 체크포인트 전처리: XLSX / CSV / TXT + Gemini CLI

XLSX의 `체크포인트_REF_GUIDE` 시트 또는 CSV 형식의 CSV/TXT/TSV 파일에서 Checkpoint와 Specification을 검증 규칙 후보로 구조화합니다. 설계치는 회로에서 추출할 정보의 양식으로 사용합니다. 원본 입력 파일을 수정하지 않습니다. BOM/Netlist 추출과 최종 PASS/FAIL 판정은 이 코드의 범위에 포함되지 않습니다.

## 설치

Python 3.10 이상이 필요합니다. Windows에서는 Python 3.12를 사용하셔도 됩니다. Node.js 20 이상과 Gemini CLI가 필요합니다.

저장소를 내려받고 해당 폴더에서 실행하세요. ZIP을 사용한다면 압축을 푼 코드 폴더에서 실행하시면 됩니다.

```powershell
git clone https://github.com/dongkyu20/circuit-checkpoint-preprocessor.git
cd circuit-checkpoint-preprocessor
```

```powershell
python -m pip install -r requirements.txt
npm install -g @google/gemini-cli
gemini
```

`gemini` 화면에서 본인의 인증 방식을 설정하고 종료합니다. 기존 Gemini CLI 인증을 그대로 사용하며, Python에서 별도 Gemini SDK/API 키를 요구하지 않습니다. 설정한 CLI 인증 방식에 따라 사용량 제한이나 과금 조건은 달라집니다.

설치 확인:

```powershell
gemini --version
gemini --prompt "Return only the word OK. Do not use tools." --output-format json
```

위 명령의 JSON에는 `response`가 있어야 합니다. CLI의 `--output-format json`은 모델의 답변 자체를 Schema로 강제하는 옵션이 아닙니다. 이 프로그램은 CLI JSON을 파싱하고, 그 안의 `response`를 다시 JSON으로 파싱한 뒤 구조를 검사합니다.

## 실행

XLSX 기본 시트 이름은 `체크포인트_REF_GUIDE`, 헤더는 첫 번째 행입니다. 필수 열은 `Checkpoint`, `Specification`, `설계치`입니다. 영문 열 이름의 대소문자와 앞뒤 공백은 허용합니다. `adminno`, `seqno`, `참조문서`, `pin_info`도 있으면 사용합니다. CSV/TXT는 시트가 없으므로 `--sheet`를 사용하지 않아도 됩니다.

실제 파일명을 지정하세요. 경로에 공백이 있으면 따옴표로 감싸세요.

### CSV를 TXT로 저장한 경우

파일 확장자로 입력 방식을 선택합니다. CSV 내용을 그대로 TXT로 저장한 경우 아래처럼 실행하세요. `rules.txt`는 실제 파일명으로 바꾸세요.

```powershell
git pull
python preprocess.py "rules.txt" --output "result" --limit 30 --dry-run
python preprocess.py "rules.txt" --output "result" --limit 30 --batch-size 5
```

쉼표, 탭, 세미콜론, 파이프 구분자는 필수 헤더를 기준으로 감지합니다. UTF-8(BOM 포함), BOM이 있는 UTF-16, CP949 인코딩은 엄격하게 디코딩합니다. 실제 사용한 인코딩과 구분자는 출력 JSON과 summary에 기록합니다. CP949 파일의 한글 내용이 의도와 맞는지는 dry-run 결과에서 확인하세요. 수동 지정도 가능합니다.

```powershell
python preprocess.py "rules.txt" --encoding cp949 --delimiter tab --output "result" --limit 30
python preprocess.py "rules.csv" --encoding utf-8-sig --delimiter comma --output "result" --limit 30
```

TXT에는 아래와 같이 헤더와 열 구분이 유지되어 있어야 합니다.

```csv
adminno,seqno,Checkpoint,Specification,설계치,참조문서
1336023553,0,"미사용 입력 핀을 VCC 또는 GND에 연결","VCC 또는 GND 연결","Unused input pin:","chip.pdf"
```

셀 안의 쉼표, 개행과 큰따옴표는 정상적인 CSV 인용 규칙을 유지해야 합니다. CSV를 메모장에서 그대로 TXT로 저장하는 것은 가능하지만, 셀을 한 줄씩 나열한 텍스트나 여러 시트의 표를 그대로 이어 붙인 파일은 입력 형식이 다릅니다. 이 경우 체크포인트 표 하나의 CSV 형태로 다시 저장하세요. 잘못된 열 수나 닫히지 않은 따옴표는 자동으로 보정하지 않고 오류를 표시합니다.

CSV/TXT의 식별자는 문자열로 보존하여 선행 0을 잃지 않습니다. 출력의 `sheet`와 `excel_row`는 null이며 `record_row`가 헤더를 포함한 논리 레코드 번호를 나타냅니다. 셀 안에 여러 줄이 있으면 실제 텍스트 줄 번호와 다를 수 있습니다. `row_uid`는 파일명과 레코드 번호로 구성됩니다. 동일한 정규화 내용을 가진 기존 XLSX 결과는 캐시를 재사용합니다.

CSV/TXT 입력만 사용할 때는 openpyxl이 필요하지 않습니다. 실제 Gemini 분석에는 설치 및 로그인된 Gemini CLI가 필요합니다.

### 중복 헤더 처리

헤더를 비교할 때 앞뒤 공백, 대소문자, 공백 차이를 정규화합니다. 예를 들어 `Admin No`와 `adminno`는 동일한 처리용 열로 인식됩니다.

- 처리에 쓰지 않는 보조 열이 중복되면 `category [col 4]`, `category [col 5]`처럼 열 번호를 붙여 두 열의 값을 모두 보존합니다. 원래 이름과 저장 이름은 출력의 `renamed_headers`에서 확인할 수 있습니다.
- 조인으로 `adminno`가 여러 개 남은 경우 값이 같으면 첫 번째 열을 식별자로 사용하고 원본 열들은 모두 보존합니다. 값이 다른 행은 자동 선택하지 않고 해당 레코드와 열 위치를 표시합니다. 사용할 열은 `--adminno-column N`으로 지정할 수 있습니다.
- 여러 `seqno`는 서로 다른 테이블에서 온 별도 필드로 취급합니다. 값이 달라도 오류를 내지 않으며, 같더라도 자동 병합하지 않습니다. 각 열의 번호, 원래 헤더, 저장 키, 값은 `seqno_fields`에 별개로 보존합니다. `seqno=0`도 정상 값입니다.
- `seqno`가 여러 개인데 대표 열을 지정하지 않으면 `checkpoint_id`는 원본 파일/시트와 행 위치로 생성하고 `checkpoint_id_basis`는 `row`가 됩니다. `normalized_source.seqno`는 빈 문자열이며 원래 값들은 `seqno_fields`와 `source_columns`에서 확인할 수 있습니다. 데이터 행을 추적하는 ID이며 어느 테이블의 업무 번호인지 임의로 결정하지 않습니다.
- 체크포인트 번호에 해당하는 열을 아는 경우 `--seqno-column N`으로 `adminno:seqno` 형식의 ID를 만들 수 있습니다. 이때도 모든 `seqno` 원본 열을 별개로 보존합니다. N은 전체 표에서 1부터 시작하는 열 번호입니다. 선택 내역은 `adminno_selection`, `seqno_selection`에 남습니다.
- `Checkpoint`, `Specification`, `설계치`, `참조문서`, `pin_info`가 중복되면 어떤 열이 실제 기준인지 추측하지 않습니다. 오류에 이름과 열 위치를 표시합니다. 사용할 열만 해당 이름으로 남기고, 다른 열은 의미가 구분되는 이름으로 바꿔 주세요.
- 이 오류는 데이터 행 중복이나 문장 정규화 오류를 뜻하지 않습니다. 입력 헤더의 이름 충돌입니다.

예를 들어 사용할 `adminno`가 전체 표의 4번째 열이라면:

```powershell
python preprocess.py "rules.txt" --adminno-column 4 --output "result" --limit 30 --dry-run
```

사용할 `seqno`가 전체 표의 6번째 열이라면 두 옵션을 함께 지정할 수 있습니다.

```powershell
python preprocess.py "rules.txt" --adminno-column 4 --seqno-column 6 --output "result" --limit 30 --dry-run
```

### 1. AI 호출 없이 읽기 확인

```powershell
python preprocess.py "checkpoint.xlsx" --output "result" --limit 30 --dry-run
```

`result/prepared.jsonl`에서 원문과 정규화한 값, 행 번호를 확인합니다. HTML entity를 복원하고 개행/공백과 일부 표시용 escape를 정리합니다. `/OE`, 단위, 조건문, 첨부 파일명은 삭제하지 않습니다.

### 2. 먼저 30행 처리

```powershell
python preprocess.py "checkpoint.xlsx" --output "result" --limit 30 --batch-size 5
```

`review.csv`를 열어 원문 충돌, 권고/필수, 외부 표 누락, 핀 목록이 적절하게 표현됐는지 검토합니다. 설명은 한국어이며 부품명과 기술 기호는 보존하도록 프롬프트에 지정했습니다.

### 3. 전체 처리

```powershell
python preprocess.py "checkpoint.xlsx" --output "result" --batch-size 10
```

처음 30행에서 성공한 결과는 같은 출력 폴더의 SQLite 캐시에서 재사용합니다. CLI 호출은 순차 실행됩니다. 최대 10개 고유 행을 한 번에 전송하며, 행 내용이 길면 자동으로 더 작은 배치를 만듭니다. 1만~1만 2천 개 고유 행이면 최초 배치 호출은 대략 1천~1천 2백 회이며, 길이 제한과 재시도 때문에 더 늘어날 수 있습니다. 의미와 추출 품질을 평가하는 데 사용하는 계정의 처리량도 고려하세요.

모델 지정이 필요하면 설치된 CLI에서 사용할 수 있는 모델 ID를 넣으세요:

```powershell
python preprocess.py "checkpoint.xlsx" --output "result" --model "YOUR_AVAILABLE_MODEL_ID"
```

같은 데이터라도 지정 모델이 바뀌면 캐시를 분리합니다. 모델을 생략하면 CLI 기본 모델을 사용합니다. 기본 모델, 인증 환경, CLI 버전 등 외부 설정을 변경해 기존 결과를 다시 생성하려면 `--cache-tag`를 바꾸세요. 프롬프트와 Schema 파일의 내용이 바뀌면 자동으로 새 요청 키를 계산합니다.

```powershell
python preprocess.py "checkpoint.xlsx" --output "result" --cache-tag "round2"
```

## 파일 구성

| 파일 | 역할 |
|---|---|
| `preprocess.py` | XLSX/CSV/TXT 읽기, CLI 호출, 구조/근거 검증, 캐시, 결과 저장 |
| `prompt.txt` | 한국어 전처리 지침. 필요하면 직접 수정 |
| `analysis.schema.json` | 모델이 출력할 analysis 구조 |
| `requirements.txt` | XLSX 읽기에 openpyxl 사용. CSV/TXT는 Python 표준 라이브러리 사용 |
| `tests/` | 오프라인 단위/통합 테스트 |

## 실행 결과

| 경로 | 용도 |
|---|---|
| `result/prepared.jsonl` | 현재 선택한 행의 원문과 정규화 데이터 |
| `result/checkpoints.jsonl` | 현재 선택한 행의 최종 전처리 결과. 실패/미처리 행도 포함 |
| `result/review.csv` | Excel에서 열 수 있는 검토용 요약 |
| `result/state.sqlite3` | 성공 결과와 실패 기록. 재실행을 위해 보존 |
| `result/summary.json` | 처리/캐시/실패/미처리 수와 실행 요약 |
| `result/cli_logs/` | 실제 CLI의 stdout/stderr 원문. 답변과 사용 통계가 포함될 수 있음 |

`prepared.jsonl`, 결과 파일과 요약은 매 실행 시 현재 선택 범위로 다시 작성합니다. `--limit 30`으로 실행하면 출력 파일은 30행 분량입니다. SQLite에는 이전에 성공한 다른 행의 캐시도 남습니다. 전체 결과 파일을 복구하려면 `--limit` 없이 다시 실행하세요.

## JSONL 구조

JSONL 한 줄은 원본의 한 행에 대응합니다. 복합 체크포인트는 `analysis.checks`의 여러 하위 규칙으로 분리됩니다.

```json
{
  "checkpoint_id": "1336023553:0",
  "row_uid": "체크포인트_REF_GUIDE:2",
  "excel_row": 2,
  "source_columns": {"Checkpoint": "원본 셀", "Specification": "원본 셀", "설계치": "원본 셀"},
  "normalized_source": {"checkpoint": "정규화한 원문"},
  "status": "normalized",
  "review_state": "candidate",
  "analysis": {
    "summary": "체크포인트 요약",
    "checks": [
      {
        "requirement_id": "체크포인트_REF_GUIDE:2:R1",
        "target_kind": "input_pin",
        "target_description": "미사용 데이터 입력 핀",
        "part_numbers": [],
        "applicability": "미사용 입력인 경우",
        "requirement": "미사용 입력은 해당 전원 또는 GND에 연결",
        "level": "mandatory",
        "parameters": [],
        "required_design_facts": ["pin_usage", "connected_net"],
        "evidence": [{"field": "checkpoint", "quote": "원문에 실제 존재하는 문구"}]
      }
    ],
    "issues": []
  }
}
```

위 JSON은 주요 필드의 구조 예시이며 실제 분석 결과가 아닙니다. 전체 필드는 실제 출력과 Schema를 확인하세요. `analysis.schema.json`은 `analysis`에 적용됩니다. 출력 시 코드가 각 하위 규칙에 `requirement_id`를 추가합니다.

| status | 의미 |
|---|---|
| `normalized` | 응답 구조와 근거 인용 검사를 통과함 |
| `error` | CLI/형식/근거 검증에 실패한 행 |
| `pending` | 아직 처리되지 않은 행 |
| `skipped` | Checkpoint/Specification이 비어 있거나 관련 셀에 수식이 있어 자동 처리하지 않은 행 |

`review_state=candidate`는 검증 규칙 후보입니다. 내용의 전기적 타당성이 승인됐다는 뜻은 아닙니다. 원문 중복 ID, 식별자 누락, 원문 충돌 등 문제가 있으면 `needs_review`입니다. 실제 설계 PASS/FAIL과는 다른 상태입니다.

## 재시작과 오류 처리

- Ctrl+C로 중단하면 성공한 결과를 유지하고 미처리 상태를 출력합니다. 같은 명령과 출력 폴더로 다시 실행하면 이어서 처리합니다.
- 강제 종료로 출력 파일이 갱신되지 않아도 커밋된 성공 결과는 SQLite에 남습니다. 프로세스가 완전히 종료됐는지 확인한 뒤 남은 `run.lock`을 삭제하고 재실행합니다.
- 누락/잘못된 형식의 행만 재시도합니다. 계속 실패하는 복수 행은 배치를 나눠 재시도합니다. 유효한 결과는 즉시 저장합니다.
- 인증/설정 오류는 전체 실행을 중단합니다. 기타 CLI 오류는 정해진 횟수만 재시도한 뒤 중단합니다. 계정 한도나 과금 조건은 프로그램이 변경하지 않습니다.
- 종료 코드 `0`은 선택 범위에서 처리 대상인 행이 전처리됐거나 dry-run이 완료됐음을 뜻합니다. `2`는 오류/미처리 행 또는 실행 오류가 남아 있음을 뜻합니다. `130`은 Ctrl+C 중단입니다. `0`이어도 `needs_review` 및 `skipped`를 검토하세요.
- 실행 중 `review.csv`를 Excel에서 열어 잠그면 Windows에서 결과 저장이 실패할 수 있습니다. 실행이 끝난 뒤 파일을 여세요.
- 한 출력 폴더에 두 프로세스를 동시에 실행하지 마세요.

헤더가 다른 위치에 있으면:

```powershell
python preprocess.py "checkpoint.xlsx" --header-row 3 --sheet "체크포인트_REF_GUIDE" --output "result"
```

CLI를 찾지 못하면:

```powershell
Get-Command gemini
python preprocess.py "checkpoint.xlsx" --gemini "C:\Users\YOUR_USER\AppData\Roaming\npm\gemini.cmd" --output "result"
```

호출 시간과 재시도 간격 조정:

```powershell
python preprocess.py "checkpoint.xlsx" --output "result" --batch-size 5 --timeout 300 --interval 5 --attempts 3 --backoff 15
```

`--max-batch-chars`는 행 JSON의 문자 수 목표이며 토큰 수나 CLI 최대 입력 크기가 아닙니다. Schema와 프롬프트는 여기에 추가됩니다. 단일 행이 목표보다 길어도 잘라서 조건을 잃지 않도록 그대로 전송합니다.

## 전처리 정책과 한계

- 원문과 전체 열 데이터는 로컬 출력에 보존합니다. 모델에는 Checkpoint, Specification, 설계치 양식, 참조문서 셀, pin_info만 전송합니다.
- 내용이 정확히 같은 전처리 입력은 한 번만 분석하고 각 원본 행에 결과를 연결합니다. 식별자가 중복되어도 행은 삭제하지 않습니다. `row_uid`와 `requirement_id`는 현재 입력의 행/레코드 위치 기준이므로 정렬 후 다시 실행하면 달라질 수 있습니다.
- 데이터용 XLSX/CSV/TSV/TXT, 결과와 CLI 로그는 .gitignore에서 제외합니다. `prompt.txt`와 `requirements.txt`는 코드 파일이므로 예외로 추적합니다.
- 수식 셀은 자동 계산하지 않습니다. 관련 수식이 있으면 결과값으로 저장한 입력 사본을 사용하거나 별도 전처리로 해소해야 합니다.
- 첨부 PDF/메일의 파일명을 문서 본문으로 취급하지 않습니다. 외부 문서 읽기는 구현하지 않았습니다. 표나 본문이 필요하면 해당 원문을 별도로 확보해야 합니다.
- Gemini CLI는 headless 모드에서도 `@file` 전처리를 할 수 있어, 전송 JSON의 `@`를 가역적인 `\u0040` escape로 표현합니다. 저장된 원문과 JSON 의미는 유지합니다.
- CLI용 임시 작업 폴더에서 내장 도구, MCP 허용 목록, hooks를 비활성화하고 `--extensions none`을 사용합니다. 기존 사용자 인증과 조직의 system 설정은 바꾸지 않습니다. YOLO 모드는 사용하지 않습니다.
- 구조 검사는 이 Schema에서 사용하는 type/required/properties/additionalProperties/enum/minLength/minItems만 지원합니다. Schema를 새 기능으로 확장하면 검증기도 같이 확장하거나 일반 JSON Schema 라이브러리로 교체하세요.
- 각 evidence 인용이 해당 원문에 실제 존재하는지 확인합니다. 이 검사만으로 번역과 요약, 수치 해석, 모든 조건의 보존까지 증명할 수는 없습니다. 먼저 30행의 의미를 검토하고 전체 결과에서도 표본 및 issues 항목을 확인하세요.
- 값과 단위는 추출하되 현재 코드에서 단위 환산이나 전기적 계산은 하지 않습니다. 권고값을 필수값으로 바꾸지 않도록 프롬프트에 지정했지만 모델 출력의 의미는 검토가 필요합니다.

## 테스트

```powershell
python -m unittest discover -s tests -v
```

오프라인 테스트는 실제 XLSX 읽기와 subprocess 호출을 사용합니다. Gemini 응답은 테스트용 실행 파일로 대체합니다. 통합 테스트의 실행 파일 fixture는 POSIX shebang을 사용하므로 Windows에서는 통합 테스트가 skip되고 단위 테스트가 실행됩니다. Windows 실제 Gemini CLI 연동은 현지 설치 환경에서 확인해야 합니다.

검증 범위: XLSX/TXT 각각 12,000행 읽기/보존, UTF-8/CP949/UTF-16, CSV 구분자, 여러 줄과 큰따옴표가 있는 셀, 선행 0, 긴 셀, CSV 오류 검출, XLSX/TXT 간 캐시 재사용, 재시작, 원문 변경 재처리, 부분 응답 재시도, 인증 실패 중단, timeout, Ctrl+C, 근거 인용과 응답 형식 검사. 실제 사용자 입력 파일과 Gemini CLI가 제공되지 않아 실제 모델 출력의 품질과 계정 처리량은 측정하지 않았습니다.

## 참고한 공식 문서

- https://geminicli.com/docs/get-started/installation/
- https://geminicli.com/docs/cli/headless/
- https://geminicli.com/docs/cli/tutorials/automation/
- https://geminicli.com/docs/reference/configuration/
- https://github.com/google-gemini/gemini-cli/blob/main/packages/cli/src/nonInteractiveCli.ts

수정일: 2026-10-14
