# Seminar Generator

논문 PDF를 기반으로 RS Team Seminar용 발표자료 초안을 생성하는 도구입니다.

작업 흐름은 다음 두 단계입니다.

1. LLM에 `llm_prompt.txt`, `config/example.yaml`, 논문 PDF를 함께 입력해 논문별 YAML 설정 파일을 생성함.
2. 생성된 YAML을 `main.py`에 입력해 PowerPoint 발표자료를 생성함.

---

## 1. 설치

프로젝트 루트에서 의존성을 설치합니다.

```bash
pip install -r requirements.txt
```

필요한 주요 패키지는 다음과 같습니다.

```text
python-pptx
PyYAML
```

---

## 2. 논문별 YAML 생성

LLM에 다음 파일들을 함께 입력합니다.

```text
llm_prompt.txt
config/example.yaml
논문 PDF
```

LLM에는 `llm_prompt.txt`의 지시를 따르게 하고, 최종 출력으로 YAML 본문만 생성하게 합니다.

생성된 YAML은 `config/` 아래에 저장합니다.

예시:

```text
config/attention_is_all_you_need.yaml
```

YAML은 반드시 다음 구조를 가져야 합니다.

```yaml
schema_version: "1.0"

deck:
    output_name: "attention_is_all_you_need"
    language: "ko"
    eod_title: "EOD"

cover:
    title: "Attention Is All You Need"
    subtitle: "NeurIPS 2017 [link]"
    date: "2026-06-10"
    venue: "RS Team Seminar"

sections:
    - id: "sec_motivation"
      title: "Motivation"
      contents:
          - "motivation_problem"

main:
    - id: "motivation_problem"
      title: "기존 sequence model의 병목"
      body: |
          RNN 기반 모델은 token을 순차적으로 처리해야 했음.
          긴 sequence에서 병렬화가 어렵고 학습 비용이 커지는 문제가 있었음.
          이 논문은 attention만으로 sequence transduction을 처리하는 구조를 제안함.
```

주의할 점은 다음과 같습니다.

- `sections[*].contents`에 적은 id는 반드시 `main[*].id`에 존재해야 함.
- `main`에 있는 모든 슬라이드는 반드시 어떤 section에서 참조되어야 함.
- `body`는 `body: |` 형식의 일반 텍스트로 작성해야 함.
- `body` 안에서는 bullet point나 번호 목록을 사용하지 않아야 함.
- 논문마다 슬라이드 수와 섹션 구성은 달라질 수 있음.

---

## 3. PPTX 생성

생성한 YAML을 `--config` 인자로 전달해 실행합니다.

```bash
python main.py --config config/attention_is_all_you_need.yaml
```

성공하면 `results/` 디렉터리에 PowerPoint 파일이 생성됩니다.

```text
results/attention_is_all_you_need.pptx
```

출력 파일명은 YAML의 `deck.output_name` 값을 기반으로 결정됩니다.

---

## 4. 실행 예시

기본 예시 YAML로 발표자료를 생성하려면 다음을 실행합니다.

```bash
python main.py --config config/example.yaml
```

생성 결과:

```text
results/paper_seminar_draft.pptx
```

---

## 5. 출력 경로 직접 지정

원하는 위치에 PPTX를 저장하려면 `--out`을 사용합니다.

```bash
python main.py --config config/attention_is_all_you_need.yaml --out results/my_seminar.pptx
```

---

## 6. 템플릿 지정

기본 템플릿은 다음 파일입니다.

```text
RS 세미나 템플릿.pptx
```

다른 템플릿을 사용하려면 `--template`을 지정합니다.

```bash
python main.py --config config/attention_is_all_you_need.yaml --template "RS 세미나 템플릿.pptx"
```

템플릿은 최소한 다음 3개의 샘플 슬라이드를 포함해야 합니다.
`main.py`는 이 3장의 레이아웃을 가져온 뒤 샘플 슬라이드를 삭제하고 새 발표자료를 생성합니다.

```text
1번째 슬라이드: cover slide
2번째 슬라이드: section slide
3번째 슬라이드: main content slide
```

---

## 7. 템플릿 구조 확인

템플릿의 placeholder 정보를 확인하려면 다음을 실행합니다.

```bash
python main.py --inspect-template
```

특정 템플릿을 확인하려면 다음처럼 실행합니다.

```bash
python main.py --inspect-template --template "RS 세미나 템플릿.pptx"
```

---

## 8. 전체 워크플로우 요약

```text
논문 PDF 준비
→ LLM에 llm_prompt.txt + config/example.yaml + 논문 PDF 입력
→ 논문별 YAML 생성
→ config/[논문제목].yaml 저장
→ python main.py --config config/[논문제목].yaml 실행
→ results/[output_name].pptx 생성
```
