# Examples

## EUV, 파장 너머의 문제

- [`euv-wavelength-beyond-platform-economics-ko.pptx`](./euv-wavelength-beyond-platform-economics-ko.pptx): 한국어 14장 고밀도 기술 발표자료
- [`euv-wavelength-beyond-platform-economics-ko.mjs`](./euv-wavelength-beyond-platform-economics-ko.mjs): `@oai/artifact-tool` 기반 생성 소스
- [`assets/euv/`](./assets/euv/): ImageGen으로 만든 무문자 scientific visual 7종과 재현용 prompt

Storyline:

`Pattern Transfer → Resolution Limit → DUV Extension → EUV Transition → Platform Economics`

기존 20장 흐름에서 같은 질문에 답하던 장을 합쳐 Resolution+k₁+OPC, DUV immersion+resolution 비교, Mo/Si+Bragg, Sn+LPP, High-NA+anamorphic, stage dynamics+economics를 각각 한 장의 D3 기술 슬라이드로 압축했습니다. 모든 본문과 핵심 라벨·수치·도식은 PowerPoint native object로 구성되어 개별 편집할 수 있습니다. 7장의 생성 이미지는 표지와 복잡한 광학·source·stage 장면에만 제한적으로 사용한 bounded visual asset이며, 완성 슬라이드를 통째로 평면화하지 않습니다. 외부 수치와 역사적 사실의 URL, 생성 이미지 provenance는 각 슬라이드 speaker notes의 `[Sources]` 블록에 들어 있습니다.

생성 소스는 다음 검증 산출물을 `work/euv-deck-dense/` 아래에 만듭니다.

- 슬라이드별 PNG와 layout JSON
- PPTX 재불러오기 후 슬라이드별 PNG
- object / text / notes inspect 결과

PPTX의 편집 가능 구조는 저장소의 validator로 확인할 수 있습니다.

```bash
node skills/tightened-slide-pptx/scripts/validate-editable-pptx.mjs \
  examples/euv-wavelength-beyond-platform-economics-ko.pptx
```
