/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeLanguage, type OfficeLanguage } from "./office-languages";

export type PageFollowTextKey = "tier" | "hint" | "stateFollow" | "stateCustom";
type Words = Readonly<Record<PageFollowTextKey, string>>;

const TEXT: Record<OfficeLanguage, { office: Words; note: Words }> = {
  en: {
    office: { tier: "Follow workspace member roles", hint: "Members who can edit the workspace edit this file; members who can view it read it, including people who join later.", stateFollow: "Follows workspace member roles", stateCustom: "Custom access for this file" },
    note: { tier: "Follow notebook member roles", hint: "Members who can edit the notebook edit this file; members who can view it read it, including people who join later.", stateFollow: "Follows notebook member roles", stateCustom: "Custom access for this file" },
  },
  zh: {
    office: { tier: "跟随工作区成员权限", hint: "工作区里可编辑的成员可以编辑此文件，只读成员可以查看；以后加入的成员也一样。", stateFollow: "当前：跟随工作区成员权限", stateCustom: "当前：此文件使用自定义权限" },
    note: { tier: "跟随笔记本成员权限", hint: "笔记本里可编辑的成员可以编辑此文件，只读成员可以查看；以后加入的成员也一样。", stateFollow: "当前：跟随笔记本成员权限", stateCustom: "当前：此文件使用自定义权限" },
  },
  "zh-TW": {
    office: { tier: "跟隨工作區成員權限", hint: "工作區中可編輯的成員可以編輯此檔案，唯讀成員可以檢視；之後加入的成員也一樣。", stateFollow: "目前：跟隨工作區成員權限", stateCustom: "目前：此檔案使用自訂權限" },
    note: { tier: "跟隨筆記本成員權限", hint: "筆記本中可編輯的成員可以編輯此檔案，唯讀成員可以檢視；之後加入的成員也一樣。", stateFollow: "目前：跟隨筆記本成員權限", stateCustom: "目前：此檔案使用自訂權限" },
  },
  ja: {
    office: { tier: "ワークスペースのメンバー権限に従う", hint: "ワークスペースで編集できるメンバーはこのファイルを編集でき、閲覧のみのメンバーは閲覧できます。後から参加した人も同じです。", stateFollow: "現在：ワークスペースのメンバー権限に従う", stateCustom: "現在：このファイル独自のアクセス設定" },
    note: { tier: "ノートブックのメンバー権限に従う", hint: "ノートブックで編集できるメンバーはこのファイルを編集でき、閲覧のみのメンバーは閲覧できます。後から参加した人も同じです。", stateFollow: "現在：ノートブックのメンバー権限に従う", stateCustom: "現在：このファイル独自のアクセス設定" },
  },
  ko: {
    office: { tier: "작업 공간 멤버 권한 따르기", hint: "작업 공간에서 편집할 수 있는 멤버는 이 파일을 편집하고, 보기 전용 멤버는 볼 수 있습니다. 나중에 참여한 사람도 같습니다.", stateFollow: "현재: 작업 공간 멤버 권한을 따름", stateCustom: "현재: 이 파일만의 사용자 지정 권한" },
    note: { tier: "노트북 멤버 권한 따르기", hint: "노트북에서 편집할 수 있는 멤버는 이 파일을 편집하고, 보기 전용 멤버는 볼 수 있습니다. 나중에 참여한 사람도 같습니다.", stateFollow: "현재: 노트북 멤버 권한을 따름", stateCustom: "현재: 이 파일만의 사용자 지정 권한" },
  },
  de: {
    office: { tier: "Mitgliederrollen des Arbeitsbereichs folgen", hint: "Mitglieder, die den Arbeitsbereich bearbeiten können, bearbeiten diese Datei; Mitglieder mit Leserechten sehen sie – auch wer später beitritt.", stateFollow: "Folgt den Mitgliederrollen des Arbeitsbereichs", stateCustom: "Eigene Zugriffsrechte für diese Datei" },
    note: { tier: "Mitgliederrollen des Notizbuchs folgen", hint: "Mitglieder, die das Notizbuch bearbeiten können, bearbeiten diese Datei; Mitglieder mit Leserechten sehen sie – auch wer später beitritt.", stateFollow: "Folgt den Mitgliederrollen des Notizbuchs", stateCustom: "Eigene Zugriffsrechte für diese Datei" },
  },
  fr: {
    office: { tier: "Suivre les rôles des membres de l’espace de travail", hint: "Les membres qui peuvent modifier l’espace de travail modifient ce fichier ; ceux qui peuvent l’afficher le consultent, y compris les personnes qui rejoignent plus tard.", stateFollow: "Suit les rôles des membres de l’espace de travail", stateCustom: "Accès personnalisé pour ce fichier" },
    note: { tier: "Suivre les rôles des membres du carnet", hint: "Les membres qui peuvent modifier le carnet modifient ce fichier ; ceux qui peuvent l’afficher le consultent, y compris les personnes qui rejoignent plus tard.", stateFollow: "Suit les rôles des membres du carnet", stateCustom: "Accès personnalisé pour ce fichier" },
  },
  es: {
    office: { tier: "Seguir los roles de los miembros del espacio de trabajo", hint: "Los miembros que pueden editar el espacio de trabajo editan este archivo; los que pueden verlo lo leen, también quienes se unan después.", stateFollow: "Sigue los roles de los miembros del espacio de trabajo", stateCustom: "Acceso personalizado para este archivo" },
    note: { tier: "Seguir los roles de los miembros del cuaderno", hint: "Los miembros que pueden editar el cuaderno editan este archivo; los que pueden verlo lo leen, también quienes se unan después.", stateFollow: "Sigue los roles de los miembros del cuaderno", stateCustom: "Acceso personalizado para este archivo" },
  },
  it: {
    office: { tier: "Segui i ruoli dei membri dello spazio di lavoro", hint: "I membri che possono modificare lo spazio di lavoro modificano questo file; quelli che possono visualizzarlo lo leggono, anche chi entra in seguito.", stateFollow: "Segue i ruoli dei membri dello spazio di lavoro", stateCustom: "Accesso personalizzato per questo file" },
    note: { tier: "Segui i ruoli dei membri del blocco note", hint: "I membri che possono modificare il blocco note modificano questo file; quelli che possono visualizzarlo lo leggono, anche chi entra in seguito.", stateFollow: "Segue i ruoli dei membri del blocco note", stateCustom: "Accesso personalizzato per questo file" },
  },
  "pt-BR": {
    office: { tier: "Seguir as funções dos membros do espaço de trabalho", hint: "Membros que podem editar o espaço de trabalho editam este arquivo; os que podem ver, leem — inclusive quem entrar depois.", stateFollow: "Segue as funções dos membros do espaço de trabalho", stateCustom: "Acesso personalizado para este arquivo" },
    note: { tier: "Seguir as funções dos membros do caderno", hint: "Membros que podem editar o caderno editam este arquivo; os que podem ver, leem — inclusive quem entrar depois.", stateFollow: "Segue as funções dos membros do caderno", stateCustom: "Acesso personalizado para este arquivo" },
  },
  nl: {
    office: { tier: "Ledenrollen van de werkruimte volgen", hint: "Leden die de werkruimte kunnen bewerken, bewerken dit bestand; leden die mogen bekijken, lezen het – ook wie later lid wordt.", stateFollow: "Volgt de ledenrollen van de werkruimte", stateCustom: "Eigen toegang voor dit bestand" },
    note: { tier: "Ledenrollen van het notitieboek volgen", hint: "Leden die het notitieboek kunnen bewerken, bewerken dit bestand; leden die mogen bekijken, lezen het – ook wie later lid wordt.", stateFollow: "Volgt de ledenrollen van het notitieboek", stateCustom: "Eigen toegang voor dit bestand" },
  },
  da: {
    office: { tier: "Følg medlemsrollerne i arbejdsområdet", hint: "Medlemmer, der kan redigere arbejdsområdet, redigerer denne fil; medlemmer, der kan se det, kan læse den – også dem, der kommer med senere.", stateFollow: "Følger medlemsrollerne i arbejdsområdet", stateCustom: "Tilpasset adgang til denne fil" },
    note: { tier: "Følg medlemsrollerne i notesbogen", hint: "Medlemmer, der kan redigere notesbogen, redigerer denne fil; medlemmer, der kan se den, kan læse filen – også dem, der kommer med senere.", stateFollow: "Følger medlemsrollerne i notesbogen", stateCustom: "Tilpasset adgang til denne fil" },
  },
  no: {
    office: { tier: "Følg medlemsrollene i arbeidsområdet", hint: "Medlemmer som kan redigere arbeidsområdet, redigerer denne filen; medlemmer som kan se det, leser den – også de som blir med senere.", stateFollow: "Følger medlemsrollene i arbeidsområdet", stateCustom: "Egendefinert tilgang til denne filen" },
    note: { tier: "Følg medlemsrollene i notatboken", hint: "Medlemmer som kan redigere notatboken, redigerer denne filen; medlemmer som kan se den, leser filen – også de som blir med senere.", stateFollow: "Følger medlemsrollene i notatboken", stateCustom: "Egendefinert tilgang til denne filen" },
  },
  sv: {
    office: { tier: "Följ medlemsrollerna i arbetsytan", hint: "Medlemmar som kan redigera arbetsytan redigerar den här filen; medlemmar som kan visa den läser filen – även de som går med senare.", stateFollow: "Följer medlemsrollerna i arbetsytan", stateCustom: "Anpassad åtkomst för den här filen" },
    note: { tier: "Följ medlemsrollerna i anteckningsboken", hint: "Medlemmar som kan redigera anteckningsboken redigerar den här filen; medlemmar som kan visa den läser filen – även de som går med senare.", stateFollow: "Följer medlemsrollerna i anteckningsboken", stateCustom: "Anpassad åtkomst för den här filen" },
  },
  fi: {
    office: { tier: "Noudata työtilan jäsenten rooleja", hint: "Jäsenet, jotka voivat muokata työtilaa, muokkaavat tätä tiedostoa; lukuoikeuden jäsenet voivat lukea sen – myös myöhemmin liittyvät.", stateFollow: "Noudattaa työtilan jäsenten rooleja", stateCustom: "Tämän tiedoston omat käyttöoikeudet" },
    note: { tier: "Noudata muistikirjan jäsenten rooleja", hint: "Jäsenet, jotka voivat muokata muistikirjaa, muokkaavat tätä tiedostoa; lukuoikeuden jäsenet voivat lukea sen – myös myöhemmin liittyvät.", stateFollow: "Noudattaa muistikirjan jäsenten rooleja", stateCustom: "Tämän tiedoston omat käyttöoikeudet" },
  },
  id: {
    office: { tier: "Ikuti peran anggota ruang kerja", hint: "Anggota yang dapat mengedit ruang kerja mengedit file ini; anggota yang hanya dapat melihat bisa membacanya — termasuk yang bergabung nanti.", stateFollow: "Mengikuti peran anggota ruang kerja", stateCustom: "Akses khusus untuk file ini" },
    note: { tier: "Ikuti peran anggota buku catatan", hint: "Anggota yang dapat mengedit buku catatan mengedit file ini; anggota yang hanya dapat melihat bisa membacanya — termasuk yang bergabung nanti.", stateFollow: "Mengikuti peran anggota buku catatan", stateCustom: "Akses khusus untuk file ini" },
  },
  vi: {
    office: { tier: "Theo quyền thành viên của không gian làm việc", hint: "Thành viên có thể chỉnh sửa không gian làm việc sẽ chỉnh sửa được tệp này; thành viên chỉ xem thì đọc được — kể cả người tham gia sau.", stateFollow: "Đang theo quyền thành viên của không gian làm việc", stateCustom: "Quyền truy cập riêng cho tệp này" },
    note: { tier: "Theo quyền thành viên của sổ tay", hint: "Thành viên có thể chỉnh sửa sổ tay sẽ chỉnh sửa được tệp này; thành viên chỉ xem thì đọc được — kể cả người tham gia sau.", stateFollow: "Đang theo quyền thành viên của sổ tay", stateCustom: "Quyền truy cập riêng cho tệp này" },
  },
  th: {
    office: { tier: "ใช้สิทธิ์ตามสมาชิกพื้นที่ทำงาน", hint: "สมาชิกที่แก้ไขพื้นที่ทำงานได้จะแก้ไขไฟล์นี้ได้ สมาชิกที่ดูได้อย่างเดียวจะอ่านได้ รวมถึงผู้ที่เข้าร่วมภายหลัง", stateFollow: "ปัจจุบัน: ใช้สิทธิ์ตามสมาชิกพื้นที่ทำงาน", stateCustom: "ปัจจุบัน: สิทธิ์เฉพาะของไฟล์นี้" },
    note: { tier: "ใช้สิทธิ์ตามสมาชิกสมุดบันทึก", hint: "สมาชิกที่แก้ไขสมุดบันทึกได้จะแก้ไขไฟล์นี้ได้ สมาชิกที่ดูได้อย่างเดียวจะอ่านได้ รวมถึงผู้ที่เข้าร่วมภายหลัง", stateFollow: "ปัจจุบัน: ใช้สิทธิ์ตามสมาชิกสมุดบันทึก", stateCustom: "ปัจจุบัน: สิทธิ์เฉพาะของไฟล์นี้" },
  },
};

export function pageFollowText(
  key: PageFollowTextKey,
  locale: string,
  product: "office" | "note",
): string {
  return TEXT[officeLanguage(locale)][product][key];
}
