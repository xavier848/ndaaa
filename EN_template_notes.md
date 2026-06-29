# English NDA template — what changed and what to review

## What was built
1. **`HEMA_NDA_master_EN.docx`** — an English version of the NDA master template. Same layout, same numbering, same `{{tokens}}` and `[[IF]]`/`[[END]]`/`[[REMOVE]]` markers as the Dutch master, with all fixed legal text translated to English.
2. **`app.py`** — updated so the generator picks the English template when the intake's `Language` field is English, and so the values the code injects (entity type, party designation, signatory title) come out in English too. Dutch behaviour is unchanged.

## How the language switch works
- The form already sends a `Language` field. The generator now reads it: anything starting with "en" or containing "engels" → English template; everything else (including a missing field) → Dutch. So existing Dutch submissions are completely unaffected.
- It selects `HEMA_NDA_master_EN.docx` instead of `HEMA_NDA_master.docx`.
- It produces English equivalents for the strings the code fills in itself:
  - entity type: "a private limited company" (was "een besloten vennootschap")
  - individual: "a natural person" (was "een natuurlijk persoon")
  - individual signing role: "on their own behalf" (was "namens zichzelf")
  - designation: "Counterparty" for mutual / "Supplier" for one-sided (was "Wederpartij" / "Leverancier")
  - default HEMA title: "unit manager" (was "unitmanager")

## Deployment (what you need to do in the repo)
1. Commit **both** `HEMA_NDA_master_EN.docx` and the updated `app.py` to `xavier848/ndaaa`. The English template must be in the repo root next to the Dutch one so it deploys to Render.
2. That's it — no new dependencies, no env-var changes. Render redeploys on push.
3. Confirm the form's language field value actually reaches the generator through Power Automate. The code looks for a field named `Language`. If your form/flow uses a different name, either rename it or tell me and I'll adjust the field name in `_is_english`.

## For Daan — focus the legal review here (these are judgement calls, not typos)
The translation mirrors the Dutch clause-for-clause. The spots genuinely worth a lawyer's eye:
1. **Entity description line.** English reads "a private limited liability company (besloten vennootschap met beperkte aansprakelijkheid)". The Dutch legal form is kept in brackets on purpose — standard practice in English-language Dutch contracts, since the entity really is a Dutch BV. Confirm HEMA wants it phrased this way.
2. **Defined terms.** "Group Company", "Confidential Information", "Disclosing Party", "Receiving Party", "Intellectual Property Rights", "Purpose" — confirm these are the English defined terms HEMA wants to standardise on.
3. **Counterparty designation.** Mutual → "Counterparty", one-sided → "Supplier". Confirm "Supplier" is right for the one-sided case (it maps to the Dutch "Leverancier"); if one-sided NDAs aren't always with suppliers, a neutral "Disclosing Party"/"Receiving Party" framing may be preferable.
4. **Governing law clause.** Kept as Dutch law + Amsterdam court, translated. Same substance as the Dutch master.
5. **VAB clause.** "the franchisees' association VAB (Vereniging Aangesloten Bedrijven)" — uses the confirmed full name.

## Note
This is a translation of the existing approved Dutch wording, not new legal drafting. The goal was to give Daan a complete English draft to review rather than to write from scratch.
