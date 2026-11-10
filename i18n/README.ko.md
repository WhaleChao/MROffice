# MR Office — 한국어

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

> ### ⬇ 모든 기능을 사용하려면 MR Office 데스크톱 앱(무료)을 설치하세요
>
> MR Office 데스크톱 앱은 무료이며 이 편집기들과 전체 서비스를 갖추고 있고, 내 컴퓨터에서 오프라인으로도 작동합니다. 오픈 소스 편집기에 없는 다음 기능을 제공합니다:
>
> - Word, Excel, PowerPoint 파일(.docx, .xlsx, .pptx) 열기와 저장
> - 함수가 들어간 수식(SUM, IF, VLOOKUP 등) 계산
> - 차트
> - SmartArt 편집
> - PDF 편집 내용을 파일에 다시 저장
>
> **[Releases에서 MR Office 다운로드 →](https://github.com/dongmingyang2007-pixel/MROffice/releases)** · [또는 mroffice.ai에서](https://mroffice.ai)

<p align="center"><a href="https://mroffice.ai"><b>🌐 mroffice.ai</b></a> · <a href="https://mroffice.ai">MR Office에서 사용해 보기 →</a></p>

> MR Office의 오픈 소스 편집기: 문서, 스프레드시트, 프레젠테이션, PDF, 필기, 코드가 브라우저에서 실행됩니다.

MR Office는 노트, AI, 실제 파일을 한곳에 모은 작업 공간입니다. 이 저장소는 그 편집기의 소스 코드를 공개합니다. MR Office에서 실행되는 것과 같은 편집기이며, 직접 빌드하고 실행하고 포함할 수 있습니다. 이것은 오픈 소스 편집기이지 완전한 오픈 소스 Office 제품군이 아닙니다. Microsoft Office 파일의 읽기와 쓰기는 MR Office가 제공합니다(아래 참고).

## 편집기

이 저장소에서 각 편집기가 하는 일:

- **문서** — 페이지 단위 워드 프로세서: 스타일, 글꼴, 목록, 표, 머리글과 바닥글, 구역과 단, 도형과 수식, 표지와 목차, 메모와 변경 내용 추적, 페이지 위 필기. 고유 형식(.mrdoc)으로 저장하며 HTML, Markdown, 일반 텍스트를 엽니다.
- **스프레드시트** — 셀과 수식, 숫자 서식, 조건부 서식, 정렬과 필터, 틀 고정, 여러 시트, 선택 영역의 합계와 평균. 셀 사이의 사칙연산(=A1*B1+10 등)은 입력하면 브라우저에서 계산되고, 함수를 쓰는 수식(SUM, IF, VLOOKUP 등)은 쓴 그대로 보존됩니다. 고유 형식(.mrsheet)으로 저장하며 CSV와 TSV를 열고 저장합니다.
- **프레젠테이션** — 고유 테마 8개가 있는 슬라이드: 레이아웃, 슬라이드 마스터와 레이아웃 보기, 텍스트, 도형, 텍스트 상자, 표, WordArt와 수식, 메모, 구역, 화면 전환과 애니메이션(미리 보기), 슬라이드 쇼. 고유 형식(.mrslides)으로 저장합니다.
- **PDF** — PDF 읽기, 검색, 프레젠테이션, 인쇄, 텍스트 복사와 소리 내어 읽기. 파일이 열려 있는 동안 선택한 텍스트에 강조 표시, 밑줄, 취소선, 물결선을 달고 메모와 메모 창을 쓸 수 있습니다. 저장하면 연 그대로의 파일을 돌려줍니다.
- **필기** — 압력 감지 펜, 형광펜, 지우개, 올가미, 눈금자. 잉크를 도형으로 바꿉니다. 고유 형식(.mrink)으로 저장합니다.
- **코드** — 여러 언어의 구문 강조, 개요, 찾기 및 바꾸기를 갖춘 가벼운 코드 편집기. JavaScript/TypeScript, Python, SQL, Lua, 셸을 브라우저에서 바로 실행합니다.

함수가 들어간 수식 계산, 차트 그리기, SmartArt 레이아웃은 MR Office 자체 엔진이 처리하며 MR Office 데스크톱 앱에서 제공됩니다. 오픈 소스 편집기는 파일에 있는 수식, 차트, SmartArt를 그대로 보존하고 SmartArt는 저장된 모습대로 표시합니다.

## Word, Excel, PowerPoint, PDF 파일

Word, Excel, PowerPoint 파일(.docx, .xlsx, .pptx) 열기와 저장, PDF 편집 내용을 파일에 다시 저장하는 기능은 무료 MR Office 데스크톱 앱이 제공합니다. 데스크톱 앱에는 이 편집기들이 모두 들어 있으며 내 컴퓨터에서 오프라인으로도 작동합니다:

- **MR Office** (무료 데스크톱 앱): https://github.com/dongmingyang2007-pixel/MROffice/releases
- **웹사이트**: https://mroffice.ai

[모든 버전 →](https://github.com/dongmingyang2007-pixel/MROffice/releases)

## 편집기 실행

여섯 편집기는 `packages/`(`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, 공통 `@mrnote/core`)에 있으며 `apps/playground`에서 함께 실행됩니다:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev`는 같은 주소에서 개발 서버를 시작하고, `npm run smoke`는 빌드를 점검합니다. 각 패키지의 README에 하는 일과 포함 방법을 자세히 적었습니다.

## 라이선스

GNU Affero General Public License v3.0(LICENSE)과 NOTICE.md의 추가 조건으로 공개합니다. "MR Office" 표시와 https://mroffice.ai 링크를 유지해야 하며, MR Office 이름과 로고에 대한 권리는 부여되지 않습니다. 상업용 라이선스는 support@mr-note.com으로 문의하세요. 서드파티 소프트웨어와 데이터는 THIRD_PARTY_NOTICES.md에 있습니다. 기여 방법은 CONTRIBUTING.md를 보세요.

## 더 보기

웹사이트와 전체 제품: **https://mroffice.ai** · 질문과 의견: support@mr-note.com
