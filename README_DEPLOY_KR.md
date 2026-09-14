# ER_Predictor 배포 안내

## 제품과 배포 경로

이 제품은 보관된 레거시 `ERTA_Predictor`와 분리된 one-folder Windows 제품입니다. 활성 소스와 작업 루트는 `D:\research\FDA_endocrine_disruption\ER_Predictor_Code`입니다. 빌드 결과와 배포 실행 파일의 정확한 경로는 다음과 같습니다.

```text
dist\ER_Predictor\ER_Predictor.exe
D:\research\FDA_endocrine_disruption\ER_Predictor\ER_Predictor\ER_Predictor.exe
```

보관된 `D:\research\FDA_endocrine_disruption\_archive\legacy_apps\ERTA_Predictor`는 변경 불가한 레거시 롤백 패키지입니다. 이 제품의 빌드/게시 스크립트는 해당 경로를 대상 또는 작업 루트로 받지 않습니다.

## K-FDA 사용자 PC 설치

K-FDA 사용자 PC에는 Python, pip, RDKit, TensorFlow 또는 기타 Python 패키지를 별도로 설치하지 않습니다. `ER_Predictor_Setup_x64.exe`가 검증된 Python 3.10 런타임과 모든 라이브러리, 모델, AD 참조 데이터를 함께 설치합니다. 시스템 Python이나 인터넷 연결에 의존하지 않는 방식이 재현성과 운영 안정성에 더 적합합니다.

```text
installer\ER_Predictor_Setup_x64.exe
```

1. 64-bit Windows 10/11 PC에서 설치 파일을 관리자 권한으로 실행합니다.
2. 기본 설치 경로는 `C:\Program Files\ER Predictor`입니다.
3. 시작 메뉴의 **ER Predictor**를 실행합니다.
4. 직접 SMILES 예측은 완전히 오프라인입니다. CAS→SMILES 조회에만 인터넷이 필요합니다.
5. 상태/캐시는 `%LOCALAPPDATA%\ER_Predictor\v1`, 단일 예측의 플롯/내보내기는 `%USERPROFILE%\Documents\ER_Predictor\Exports`에 저장됩니다. ERTA와 ERalpha batch 결과 workbook은 선택한 입력 파일과 같은 폴더에 저장됩니다.

기관 배포 전 `installer\ER_Predictor_Setup_x64.sha256.json`의 SHA-256을 전달 파일과 대조하고, 기관 코드서명 인증서가 있으면 설치 파일에 Authenticode 서명을 추가하십시오. 서명되지 않은 내부 빌드는 Windows SmartScreen 경고가 발생할 수 있습니다.

## 탭과 지원 모델

- 시작 탭은 **ERTA**이며 기존 ERTA 분류, AD, 그래프 흐름을 유지합니다.
- **ERalpha** 탭은 FDA 요청 범위인 ERα 결합 분류만 제공합니다. ERβ와 회귀 모델은 FDA 배포 화면과 설치 파일에서 제외됩니다.
- ERTA와 ERalpha의 `Options`, `Single prediction`, `Batch prediction`은 같은
  컨트롤 순서, 위치, 크기, 여백, 행/열 확장 규칙 및 열기/닫기 동작을
  사용합니다. 의도적으로 다른 결과 용어는 ERTA의 `Positive/Negative`와
  ERalpha의 `Binding/Non-binding`뿐입니다.
- 두 `Options`에는 같은 위치에 `Model`/`Browse`/`Reload model`과
  `AD reference`/`Browse`/`Reload AD`가 있습니다. ERalpha에서 선택할 수
  있는 모델은 bundled released catalog와 실행 파일 release allowlist의
  모델 ID, 크기, SHA-256이 모두 일치하는 파일뿐이며, 이 검증은 joblib
  역직렬화 전에 끝납니다. 승인된 후속 모델도 같은 release 절차를 완료한
  경우에만 선택 목록에 들어갑니다. 임의 joblib 또는 외부 AD 파일은
  로드하지 않습니다. Reload 완료는 가장 최근 작업과 사용자가 입력한
  경로가 그대로인 경우에만 적용되며, AD Reload는 새 AD 관리자를 완전히
  fit한 뒤 교체하여 실행 중 요청의 기존 AD 상태를 변경하지 않습니다.
  승인된 기본 model/AD는 두 탭 모두 시작 시 자동으로 검증하고 로드합니다.
- ERalpha 분류는 승인된 0.5 기준과 직접 결합 확률/라벨을 사용합니다.
- `Binding`과 `Non-binding`은 직접 ERα 수용체 결합 분류만 뜻합니다. 전사 활성, 작용제/길항제 활성, 신호 전달, 공동활성인자 모집 또는 일반 내분비계 장애를 뜻하지 않습니다.
- ERalpha AD는 ERTA와 같은 계산 코드를 사용하지만 ERα 전용 학습 참조와 캐시만 사용합니다.
- ERalpha batch 결과 workbook은 `Predictions`, `Guide`, `Diagnostics`, `Input`, `Metadata` 순서이며, 열 때 `Predictions`가 첫 시트이자 활성 시트입니다. `Guide`, 행별 기술 상태를 담은 `Diagnostics`, 변경하지 않은 원본 `Input`, 실행 단위 provenance를 담은 `Metadata`를 함께 유지합니다.
- ERalpha workbook의 모든 시트는 색 채우기, 결과별 강조, 임의 열 너비 같은 장식 서식을 추가하지 않습니다. 특히 기본 `Predictions` 데이터 셀 서식은 ERTA가 pandas/openpyxl로 생성하는 평문 worksheet와 동일합니다. 같은 열 이름은 읽기 편의를 위한 표현 일치이며 endpoint의 과학적 의미까지 같다는 뜻은 아닙니다.
- `Predictions`에는 충돌하지 않는 원본 입력 열이 원래 순서대로 먼저 나오고, 신뢰된 열이 정확히 `CAS`, `SMILES`, `Canonical_SMILES`, `Mol_valid`, `Probability_Negative_0`, `Probability_Positive_1`, `Prediction`, `Prediction_label`, 아래 AD 열 9개, `PubChem_CID`, `PubChem_status` 순서로 이어집니다. `Prediction`은 숫자 `0`/`1`, `Prediction_label`은 각각 `Non-binding`/`Binding`입니다. 직접 입력했거나 PubChem 응답에 값이 없으면 PubChem provenance 셀은 비어 있으며, CID나 조회 상태를 추정하거나 다시 조회하여 채우지 않습니다.
- ERalpha에서 `Probability_Negative_0`은 **Non-binding**, `Probability_Positive_1`은 **Binding** 확률입니다. ERTA의 같은 평문 열 이름에서 `Negative`/`Positive`는 전사 활성 분류를 뜻하므로 ERalpha의 직접 수용체 결합 결과와 서로 바꾸어 해석할 수 없습니다.
- ERalpha AD 열은 `AD`, `AD_MeanDistance`, `AD_DistanceThreshold`, `AD_Distance_InDomain`, `AD_SimilarityMax`, `AD_SimilarityThreshold`, `AD_Similarity_InDomain`, `AD_PC1`, `AD_PC2`입니다. 모두 ERalpha 경로의 학습 참조와 캐시로 계산하며 ERTA AD 데이터를 섞지 않습니다.
- 예측하지 못한 행은 `Predictions`의 두 확률, `Prediction`, `Prediction_label`, AD 값이 모두 빈 셀입니다. 같은 순서의 `Diagnostics` 행에서 `Row_ID`, `CAS`, 0부터 시작하는 `row_index`, `Status_Code`, `Status_Message`, `SMILES_Provenance`, `Result_Status`, `Reason_Category`, `Reason_Description`, `Recommended_Action`을 확인하십시오.
- 실행마다 일정한 `Workflow`, `Task`, `Subtype`, `Decision_rule`, `Model_ID`, `Model_SHA256`, `Preprocessing_Policy_ID`, protocol/source/split/CV/report/caveat hash, 근거 범위와 caveat는 한 행짜리 `Metadata`에만 기록됩니다. 이 값들을 각 `Predictions` 행에 반복하지 않습니다. 분류 workbook 계약 ID는 `erba.binding.classification.excel.v3`입니다.
- workbook을 자동 처리하는 코드는 v3 `Predictions` 이름을 직접 사용하고, 행 상태/사유는 같은 데이터 행 위치(또는 0부터 시작하는 `row_index`)의 `Diagnostics`에서, 모델 SHA-256과 실행 provenance는 한 행짜리 `Metadata`에서 읽어야 합니다. v2의 소문자 `raw_smiles`, `model_smiles`, `non_binding_probability`, `binding_probability`, `binding_label`이나 기본 시트의 행별 model/hash/caveat 열은 더 이상 제공하지 않습니다.
- `Guide`는 기본 결과 열, 빈 예측값, `Diagnostics`의 제외 사유·권장 조치, 원본 `Input`, 실행 단위 `Metadata`를 설명하고 전체/사유별 행 수를 제공합니다. 이름을 정규화했을 때 입력 열과 신뢰된 결과·진단·metadata 열이 충돌하면 해당 입력 열은 `Predictions`에서 제외하고 지정된 시트의 생성 값을 신뢰합니다. 원본 값과 머리글은 `Input` 시트에서 확인할 수 있습니다.
- 이 workbook/AD 보고 개선을 위해 새 모델을 도입하거나 승인하지 않았습니다. 기존에 승인된 V7 ERalpha 분류 모델과 그 과학적 한계가 그대로 적용됩니다.

**ERBA 분류 근거는 과거에 노출된 개발 행을 사용한 내부 평가입니다. 신선한 독립/외부/시간적/모집단 검증 또는 규제 검증이 아니며, 규제 용도로 사용할 수 없습니다.**

## 입력과 저장 위치

- 모든 지원 경로에서 직접 SMILES 입력은 오프라인으로 동작합니다.
- CAS 조회만 인터넷이 필요합니다. 조회 실패가 이미 유효한 SMILES를 덮어쓰지 않습니다.
- 배포 원본 예제는 bundled `templates\test.xlsx`입니다. 배포된 one-folder
  실행 파일이 `<container>\ER_Predictor\ER_Predictor.exe`이면 두 탭은
  `<container>\test.xlsx`를 함께 사용합니다. 그 외 환경에서는 최초 실행 때
  bundled 원본을
  `%USERPROFILE%\Documents\ER_Predictor\Examples\test.xlsx`로 원자적으로
  복사하여 두 탭이 같은 쓰기 가능한 파일을 사용합니다. 기존 사용자/게시
  `test.xlsx`는 덮어쓰거나 초기화하지 않습니다. 사용자가 해당 복사본을
  삭제한 뒤 다시 실행하는 것이 bundled 원본으로 되돌리는 명시적 방법입니다.
- 배포 `test.xlsx`는 원본 25개 CAS 행과 `CARSRN` 머리글의 바이트를 그대로
  보존합니다(SHA-256
  `5a1f569f8f6a5cd47bff67a189645c3f9461fcf07bd24f4e5b4f83197f3350aa`).
  `CARSRN`은 두 batch 경로에서 `CAS` 별칭으로 인식됩니다.
  단일 입력은 첫 행의 CAS로 초기화되며 workbook에 SMILES 열이 없으므로
  SMILES는 비어 있을 수 있습니다. 각 탭의 **Example input**은 이 CAS/빈
  SMILES 상태를 서로 독립적으로 복원하며, 예측 전 CAS 조회에는 인터넷이
  필요합니다.
- 기본 `test.xlsx`는 결과를 바로 옆에 쓸 수 있는 사용자 파일입니다. 반복
  실행이나 원본 보존이 필요하면 별도의 쓰기 가능한 폴더에 복사해 선택하십시오.
  **Download template**도 응용 프로그램 폴더 밖의 쓰기 가능한 사용자
  폴더만 허용하며 보호된 경로를 다른 위치로 자동 전환하지 않습니다.
- 결합 예시는 `templates\ERTA_ERBA_example.xlsx`입니다. 첫 시트 **ERBA_Input**은 ERBA 입력용이며 정확히 `Row_ID`, `CAS`, `SMILES` 열을 사용합니다. 두 번째 **ERTA_Input**은 `CID`, `CAS`, `SMILES`, `label` 열을 보여 줍니다. ERTA batch에 넣을 때는 이 시트를 별도 workbook으로 내보내십시오.

- ERTA와 ERalpha batch 페이지에는 별도의 출력 폴더 선택기가 없습니다. 입력 workbook을 선택하면 읽기 전용 **Result folder (same as input)** 항목에 입력 파일의 폴더가 표시되고, 결과 workbook은 항상 그 폴더에 원자적으로 게시됩니다. ERTA batch 그래프는 같은 위치의 `graphs` 하위 폴더에 저장됩니다.
- ERTA 결과 이름은 `ERTA_<입력파일 stem>_<모델명>_prediction.xlsx`로
  시작합니다. 같은 이름이 있으면 확장자 앞에 `_2`, `_3`을 붙입니다.
- ERalpha batch AD 그래프는 ERalpha 경로의 AD 데이터와 Binding/Non-binding 의미만 사용하고, 결과 workbook 옆의 충돌 방지 폴더 `<workbook-stem>_graphs`에 저장됩니다. 같은 이름이 이미 있으면 `_2`, `_3` 접미사를 사용하여 이전 그래프 폴더를 덮어쓰지 않습니다.
- ERalpha batch 완료 요약은 ERTA와 같은 `Prediction result` 영역에 표시되며 전체 행, Binding, Non-binding, 예측하지 못한 행, AD In-domain/Out-of-domain 수, 정확한 workbook 경로, 그래프 수/폴더 및 내부 평가 근거의 한계를 포함합니다.
- 두 endpoint 모두 새 batch를 시작하면 이전 완료 요약을 즉시
  `Batch prediction is running.`으로 교체합니다. `Run batch` 아래의 집계
  진행 표시는 `<percent>% - <current>/<total> - <stage>` 형식이며
  `Reading input workbook`(0%), `Resolving CAS/SMILES`(10–35%),
  `Preprocessing and prediction`(35–88%), `Writing workbook`(90%),
  `Completed` 또는 `Failed`(100%)를 동일하게 사용합니다. 실제 PubChem
  행 조회 상세
  `Fetching SMILES from PubChem: <index> / <total> (<CAS>)`는 집계 진행
  표시가 아니라 두 화면의 같은 하단 상태 줄에만 표시됩니다.
- 결과 workbook이 원자적으로 게시되었으면 예측하지 못한 행이나 선택 항목인
  AD/graph의 누락이 있어도 두 endpoint 모두 파란
  `Batch prediction done` 정보 팝업을 표시합니다. 공통 팝업에는 정확한
  저장 경로, 전체/예측/`Not predicted` 행 수, endpoint별 결과 수,
  graph 수/폴더와 AD/graph 상세가 빠짐없이 표시됩니다. 모든 행을 예측하지
  못했으면 `No rows could be predicted.`라고 명시하며 모든 행이 예측됐다고
  표현하지 않습니다. 지원하지 않는 입력 행이나 선택 산출물 누락에는 일반
  노란 경고 팝업을 사용하지 않습니다.
- 시작 전 거부나 읽기/모델 실행/workbook 저장·게시 실패로 결과 파일이
  생성되지 않으면 빨간 `Batch prediction failed` 오류 팝업을 표시하고
  입력·template·실행 컨트롤을 다시 활성화하며 성공 팝업을 표시하지
  않습니다. 하단 상태 줄은 두 endpoint 모두
  `Batch prediction started.`, `Batch prediction completed: <path>` 또는
  `Batch prediction failed: <details>` 형식을 사용합니다.
- ERTA는 기존 workbook schema와 0-fingerprint 추론을 변경하지 않습니다.
  화면과 팝업의 Predicted/Positive/Negative 수는 `Mol_valid=True`인 행만
  포함하며, `Mol_valid=False` 행에 legacy 호환을 위해 남은 workbook/graph
  라벨은 사용 가능한 예측이 아님을 명시합니다. 두 endpoint의 AD
  In-domain/Out-of-domain/Unavailable 수는 predicted 행만 분모로 사용하고
  `Not predicted` 행은 제외합니다.
- ERTA batch 실행 중에는 model/AD 경로 변경과 reload가 잠기며, 반대로
  model/AD reload 중에는 batch를 시작할 수 없습니다. 성공한 reload는 이전
  batch 완료 요약과 진행 상태를 초기화하므로 다른 model/AD 결과로 오인하지
  않습니다.
- 무인 native QA는 원본/기본 파일 옆에 결과를 쓰지 않고 QA 전용 폴더의
  바이트 동일 복사본 두 개를 사용합니다. 네트워크 변동을 제거하기 위해
  25개 CAS의 PubChem 응답을 동일한 유효 SMILES `C=O`로 대체하며 이
  substitution과 전체 CAS 목록을 automation transcript에 기록합니다. 이는
  실제 물질 조회 결과 검증이 아닙니다. 실제 25개 CAS의 온라인 동작은 이
  native QA 자동화 모드 없이 배포 앱에서 별도로 실행하여 확인합니다.
- 입력 파일의 상위 폴더가 보호되어 있거나 쓰기 불가능하면 예측을 시작하지 않고, 입력 workbook을 응용 프로그램 설치 폴더 밖의 쓰기 가능한 폴더로 복사한 뒤 다시 선택하라는 안내와 함께 실패합니다. 다른 위치로 자동 전환하지 않습니다.
- 입력 workbook과 이전 batch 결과 workbook은 절대 덮어쓰지 않습니다. 같은 결과 이름이 이미 있으면 `_2`, `_3`처럼 충돌하지 않는 이름을 사용합니다.

설치 폴더와 묶인 자원은 읽기 전용입니다. 기본 상태 저장소는 `%LOCALAPPDATA%\ER_Predictor\v1`, 단일 예측 플롯/내보내기 저장소는 `%USERPROFILE%\Documents\ER_Predictor\Exports`입니다. 명시적으로 쓰기 가능한 portable 설치에서만 실행 전 `ER_PREDICTOR_PORTABLE=1`을 설정하면 실행 파일 옆 `ER_Predictor_UserData\state`와 `Exports`를 사용합니다. 이 외부 사용자 저장소 설정은 batch 결과 위치를 변경하지 않습니다. Batch 결과는 portable 모드에서도 선택한 입력 파일 옆에 저장됩니다. 어떤 경우에도 보관된 `D:\research\FDA_endocrine_disruption\_archive\legacy_apps\ERTA_Predictor` 아래를 출력 위치로 사용하지 마십시오.

## 빌드 PC 준비

Windows x64와 Python 3.10이 필요합니다. 의존성 취득과 release build는 분리된 두 단계입니다. 먼저 승인된 Python 3.10 x64에서 전체 Windows wheelhouse, 정확한 전이 의존성 pin, wheel SHA-256 lock, inventory 증적을 생성하고 오프라인 해석을 검증합니다.

```bat
py -3.10 -I prepare_build_wheelhouse.py prepare
```

`py -3.10`을 찾지 못할 때만 `ER_PREDICTOR_PYTHON`에 승인된 Python 3.10 x64 실행 파일을 지정합니다. `requirements-lock.txt`와 wheelhouse inventory는 생성 증적이며 수동 pin 목록이 아닙니다. 빌드는 비어 있거나 오래된 lock, 누락된 wheel, hash 불일치, 온라인 fallback, 잘못된 Python 버전/아키텍처를 모두 거부합니다. 선정된 released 모델이 CatBoost를 필요로 하면 모델링 환경과 동일한 정확한 pin을 `requirements.txt`에 추가하고 `app.spec`의 data/native/hidden import를 활성화한 뒤 준비 단계를 다시 실행해야 합니다. Python/RDKit/TensorFlow/모델이 없는 사용자 PC에도 전체 배포 폴더를 복사하면 됩니다.

## 빌드 증적 및 비교

빌드는 catalog의 released ERBA 자산 경로, 크기, SHA-256을 먼저 검사하며 하나라도 없거나 다르면 실패합니다. 준비 단계가 통과한 뒤 다음을 실행합니다.

```bat
build_exe.bat
```

설치 파일 생성:

```bat
build_installer.bat
```

각 실행은 서로 다른 **두 개의 빈 venv**를 만들고, 검증된 로컬 wheelhouse와 `--require-hashes`만 사용하여 설치한 뒤 PyInstaller clean build를 두 번 수행합니다. 두 빌드의 component/source/SBOM-style/model/static/distribution inventory가 동일해야 하며, 불일치 시 빌드가 실패하고 `artifacts\ER_Predictor-clean-build-comparison.json`에 결과가 남습니다. resolved requirements, wheelhouse inventory, 두 reproducibility manifest, 비교 receipt, 최종 build manifest도 `artifacts`에 기록됩니다. 배포 런타임에는 절대 연구 경로가 포함되지 않습니다.

## 실행 및 게시

실행 파일 하나가 아니라 `dist\ER_Predictor` 폴더 전체를 복사한 뒤 `ER_Predictor.exe`를 실행합니다.

감사된 빌드가 끝난 뒤 인자 없이 다음을 실행합니다.

```bat
publish_exe.bat
```

스크립트는 전체 one-folder collection을 새 루트에 staging하고 실행 파일을 검사한 뒤 원자적으로 정확한 게시 경로에만 승격합니다. 게시 전후 보관된 `D:\research\FDA_endocrine_disruption\_archive\legacy_apps\ERTA_Predictor`의 재귀 manifest를 비교하여 변경을 감지하면 실패합니다. 레거시 ERTA 패키지는 절대 복사 대상이나 게시 대상이 아니며 수정하지 않습니다.
