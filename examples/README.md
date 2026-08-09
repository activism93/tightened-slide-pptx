# Examples

## EUV, 파장 너머의 문제

- [`euv-wavelength-beyond-platform-economics-ko.pptx`](./euv-wavelength-beyond-platform-economics-ko.pptx): 한국어 20장 기술 발표자료
- [`euv-wavelength-beyond-platform-economics-ko.mjs`](./euv-wavelength-beyond-platform-economics-ko.mjs): `@oai/artifact-tool` 기반 생성 소스

Storyline:

`Pattern Transfer → Resolution Limit → DUV Extension → EUV Transition → Platform Economics`

모든 본문, 선, 박스, cross-section, optical path, multilayer, source 및 High-NA schematic은 PowerPoint native object로 구성되어 개별 편집할 수 있습니다. 외부 수치와 역사적 사실의 URL은 각 슬라이드 speaker notes의 `[Sources]` 블록에 들어 있습니다.

생성 소스는 다음 검증 산출물을 `work/euv-deck/` 아래에 만듭니다.

- 슬라이드별 PNG와 layout JSON
- PPTX 재불러오기 후 슬라이드별 PNG
- object / text / notes inspect 결과

PPTX의 편집 가능 구조는 저장소의 validator로 확인할 수 있습니다.

```bash
node skills/tightened-slide-pptx/scripts/validate-editable-pptx.mjs \
  examples/euv-wavelength-beyond-platform-economics-ko.pptx
```
