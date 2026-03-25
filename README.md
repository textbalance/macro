# TextBalance

A **TextBalance** egy Microsoft Word bővítmény (makró), amely segít strukturált, kiegyensúlyozott dokumentumok írásában. Vizuális visszajelzést ad a dokumentum szerkezetéről, a szakaszok hosszáról és arányairól, hogy az írásod minél arányosabb és átláthatóbb legyen.

---

## Kinek készült?

A TextBalance szinte bárkinek hasznos lehet, aki hosszabb (5+ oldalas) dokumentumokon dolgozik:
- **Írók és szerzők** – könyvek, cikkek, blogok fejezeteinek kiegyensúlyozásához.
- **Kutatók és egyetemi hallgatók** – házi dolgozatok, szakdolgozatok, disszertációk szerkezeti felépítéséhez.
- **Előadók és lelkészek** – prezentációk, beszédek, prédikációk időzítésének és súlypontjainak megtervezéséhez.
- **Üzleti szakemberek** – jelentések, javaslatok, üzleti tervek arányosításához.

---

## Hogyan működik?

1. **Címsor-alapú elemzés** – A bővítmény a dokumentumban lévő `Címsor 1` (Heading 1) és `Címsor 2` (Heading 2) stílusokat használja kiindulópontként.
2. **Összefoglaló táblázat** – Létrehoz egy áttekintő táblázatot a dokumentum elején az összes fejezetről.
3. **Vizuális jelzések** – Minden címsorhoz automatikusan hozzáadja a karakterszámot, annak százalékos arányát, és mindezt színkóddal is ellátja:
   - **Zöld** = megfelelő hosszúság
   - **Narancs** = lassan közeledik a javasolt határhoz
   - **Piros** = túl hosszú a szakasz, érdemes lehet bontani (vagy rövidíteni)
4. **Célkövetés** – Te adhatod meg a dokumentum kívánt hosszát, és a makró ahhoz viszonyítva számolja ki a százalékokat. Továbbá lehetőség van bizonyos címsorok ("Exclude Headings" funkció) kihagyására a számításból.

---

## Telepítés és használat

A projekt tartalmaz egy telepítőt (`.exe`), amellyel könnyedén integrálhatod a TextBalance-t a Microsoft Wordbe.  
A telepítést követően a Word szalagján (Ribbon) fel fog tűnni egy különálló fül vagy ikon, amivel egyszerűen elindíthatod az elemzést.

> **Dokumentáció:**  
> A részletesebb útmutatókért, a telepítési lépésekért és a pontos működésért látogass el a projekt `tutorial` mappájába, ahol részletes leírást (angolul és magyarul) találsz a Microsoft Word címsorainak helyes használatáról is.

---

## Licenc

Ez a projekt a **MIT Licenc** feltételei szerint szabadon felhasználható, módosítható és terjeszthető (részleteket lásd a `LICENSE` fájlban).
