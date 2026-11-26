/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Words that name the product, as each surface says them (lib/product-surface):
 * MR Note's words, and on MR Office's host MR Office's. Both are rendered and
 * styles/globals.css shows the one `html[data-surface]` names, as the brand
 * lockup's word (components/brand/BrandLockup): the bootstrap stamps the
 * surface before the first paint, so a statically rendered page (the sign-in
 * pages) never reads "MR Note" on MR Office, not even for a moment.
 *
 * `office` defaults to `note` with the name swapped; without a name in it,
 * the words are the same on both and rendered once.
 */
export function SurfaceWords({ note, office }: { readonly note: string; readonly office?: string }) {
  const officeWords = office ?? note.replace(/MR ?Note/g, "MR Office");
  if (officeWords === note) return <>{note}</>;
  return (
    <>
      <span data-brand-word="note">{note}</span>
      <span data-brand-word="office">{officeWords}</span>
    </>
  );
}
